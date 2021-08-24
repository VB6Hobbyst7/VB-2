Attribute VB_Name = "S_COMSUB"
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
'  S0SUB_BIRTHDAY (ByVal PassPort_Id As String) As String
'  S0SUB_CCUR (src1 As String) As currency
'  S0SUB_CDBL (src1 As String, Slen as integer) As double
'  S0SUB_CDNAME_GET(length As Integer) As string
'  S0SUB_CINT (src1 As String) As integer
'  S0SUB_CLNG (src1 As String) As long
'  S0SUB_COM_NAME ()
'  S0SUB_CSNG (src1 As String, Slen as integer) As single
'  S0SUB_DATE_6TO8 (ByVal para As String, ByVal req As String) As String
'  S0SUB_DATE_FORMAT (req As String, delimeter As String, date1 As String, date2 As String, weekday1 As String, iret As Integer)
'  S0SUB_DELETEFILE (DirName As String, SaveDate As String)
'  S0SUB_EXIST_FILE (para As String) As Integer
'  S0SUB_Exist_RECORD (para As String) As Integer
'  S0SUB_FCDBL (src1 as string, slen integer) as string
'  S0SUB_FCDBL2(src1 as string) as string
'  S0SUB_FORMAT (para As String, defValue As String) As String
'  S0SUB_GRIDGETROW (grd As Grid, Col As Integer, para As String) As Integer
'  S0SUB_GRIDTOPROW (grd As Grid, Row As Integer, disRow As Integer)
'  S0SUB_LSPACE (src1 As String, len as integer) As string
'  S0SUB_LOGINID()
'  S0SUB_MASK_DISP (Dest_Ctl As Control, w_data As String)
'  S0SUB_MASK_DISP(ctrl As control, src1 As string)
'  S0SUB_MESSAGE (ByVal para As Integer) As String
'  S0SUB_NULL_CHECK (para As Variant) As String
'  S0SUB_NUM_FORMAT (src1 as string, fmat as string, gbcd as string)
'  S0SUB_OPEN (ByVal frmhWnd As Integer, Index As Integer)
'  S0SUB_POSITION (Frm As Form, xpos As Long, YPos As Long)
'  S0SUB_RSPACE (src1 As String, len as integer) As string
'  S0SUB_SPREADCLEAR (spd As vaSpread, DispRow As String)
'  S0SUB_SPREADGETCOL (spd As vaSpread, Row As Integer, para As String) As Integer
'  S0SUB_SPREADGETROW (spd As vaSpread, Col As Integer, para As String) As Integer
'  S0SUB_SPREADHIGHLIGHT (spd As vaSpread, Row As Integer, OldRow As Integer) As Integer
'  S0SUB_SPREADTOPROW (spd As vaSpread, Row As Integer, disRow As Integer)

'*******************************
'***  Report 관련 지역변수
'*******************************
Dim iVLs As Integer                 'Vertical Line의 갯수
Dim iPageStartTop As Integer        'Page의 시작위치를 Setting
'**********************************************************
'** 투여약물현황 보기.                                  ***
'** para    : 접수구분(1=외래, 2=병동, 3=응급실)        ***
'** par1    : 병원구분                                  ***
'** par2    : 접수일자                                  ***
'** par3    : 진찰권번호(외래)/입원번호(병동)           ***
'** frm     : Control name                              ***
'** ctr     : Control name                              ***
'**********************************************************
Sub S0SUB_SELECT_MEDICINES(frm As Form, ctr As Control, para As String, par1 As String, par2 As String, par3 As String)
  
    Dim SqlConn As Long
    Dim SqlDoc  As String: Dim sql_ret  As Integer
    Dim code()  As String
    
    ctr = ""

    If para = "1" Then   '외래
        If par1 = "3" Then
            '--- 심혈관(외래)SERVER Open
            sql_ret = S0SUB_Open(S0COM_SERVER06, frm.hWnd, SqlConn)
            If sql_ret <> QSQL_SUCCESS Then
                Exit Sub
            End If
        Else
            '--- 본원(외래)SERVER Open
            sql_ret = S0SUB_Open(S0COM_SERVER04, frm.hWnd, SqlConn)
            If sql_ret <> QSQL_SUCCESS Then
                Exit Sub
            End If
        End If
        SqlDoc = "SELECT OrdNm, OrdCapa, OrdUnit, OrdNum, OrdMeth, OrdCnt"
        SqlDoc = SqlDoc + "  FROM CL01A_DB..CL01A03M_TBL"
        SqlDoc = SqlDoc + " WHERE UnitNo = " & Chr(39) & par3 & Chr(39)
        SqlDoc = SqlDoc + "   AND (RtnYn <> 'Y' or RtnYn = null)"
        SqlDoc = SqlDoc + "   AND CalcYn = 'Y'"
        SqlDoc = SqlDoc + "   AND OrdYmd IN ( SELECT MAX(OrdYmd) FROM CL01A_DB..CL01A03M_TBL"
        SqlDoc = SqlDoc + "                    WHERE UnitNo  = " & Chr(39) & par3 & Chr(39)
        SqlDoc = SqlDoc + "                      AND OrdYmd <= " & Chr(39) & Mid$(par2, 3, 6) & Chr(39) & ")"
    Else
        '본 화면에서 사용할 Index Open
        sql_ret = S0SUB_Open(S0COM_SERVER02, frm.hWnd, SqlConn)
        If sql_ret <> QSQL_SUCCESS Then
            Exit Sub
        End If
        SqlDoc = "SELECT B.OrdNm, B.OrdCapa, B.OrdUnit, A.FrcyDay, A.DosMeth, A.NumDay"
        SqlDoc = SqlDoc + "  FROM WD01A_DB..WD1A030M_TBL A, WD01A_DB..WD1A031M_TBL B"
        SqlDoc = SqlDoc + " WHERE A.AdmiNo = B.AdmiNo"
        SqlDoc = SqlDoc + "   AND A.OrdNo  = B.OrdNo"
        SqlDoc = SqlDoc + "   AND A.AdmiNo = " & Chr(39) & par3 & Chr(39)
        SqlDoc = SqlDoc + "   AND (A.CnclYn <> 'Y' or A.CnclYn = null)"
        SqlDoc = SqlDoc + "   AND A.OrdYmd IN ( SELECT MAX(OrdYmd) FROM WD01A_DB..WD1A030M_TBL"
        SqlDoc = SqlDoc + "                    WHERE AdmiNo  = " & Chr(39) & par3 & Chr(39)
        SqlDoc = SqlDoc + "                      AND OrdYmd <= " & Chr(39) & par2 & Chr(39) & ")"
        SqlDoc = SqlDoc + "   AND A.OrdNo  IN ( SELECT MAX(OrdNo)  FROM WD01A_DB..WD1A030M_TBL"
        SqlDoc = SqlDoc + "                    WHERE AdmiNo  = " & Chr(39) & par3 & Chr(39)
        SqlDoc = SqlDoc + "                      AND OrdYmd <= " & Chr(39) & par2 & Chr(39) & ")"
    End If
    
    sql_ret = QSqlDBExec(SqlDoc, SqlConn)
    If sql_ret = QSQL_SUCCESS Then
        Do Until QSqlGetRow(record, SqlConn) <> QSQL_SUCCESS
                   
            QSqlGetField 6, record, code()
            
            ctr = ctr + Trim$(code(1)) + " " + Trim$(code(2)) + " " + Trim$(code(3)) + " " + Trim$(code(4)) + " " + Trim$(code(5)) + " for " + Trim$(code(6)) + "day" + Chr(13)

        Loop
    End If

    If Trim$(ctr) <> "" Then ctr = Mid$(ctr, 1, Len(ctr) - 1)
    
    sql_ret = QSqlSelectFree(SqlConn)
    sql_ret = Qsqlclose(SqlConn, ONECLOSE)

End Sub




'/* 환자명 얻기...
'/* para : 접수구분(1=외래, 2=병동, 3=응급실)
Sub S0SUB_SELECT_PATIENT(frm As Form, para As String)
  
    Dim SqlDoc  As String: Dim sql_ret  As Integer
    Dim patient()  As String
    
    S0COM_name = ""
    
    If para = "1" Then   '외래
        '본 화면에서 사용할 Index Open
        sql_ret = S0SUB_Open(S0COM_SERVER03, frm.hWnd, QsqlCod2)
        If sql_ret <> QSQL_SUCCESS Then
            Exit Sub
        End If
        SqlDoc = "SELECT PatNm FROM AC01B_DB..AC01B01M_TBL"
        SqlDoc = SqlDoc + " WHERE UnitNo = " & Chr(39) & S0COM_code & Chr(39)
    Else
        '본 화면에서 사용할 Index Open
        sql_ret = S0SUB_Open(S0COM_SERVER02, frm.hWnd, QsqlCod2)
        If sql_ret <> QSQL_SUCCESS Then
            Exit Sub
        End If
        SqlDoc = "SELECT PatNm FROM AD01A_DB..AD1A020M_TBL"
        SqlDoc = SqlDoc + " WHERE UnitNo = " & Chr(39) & S0COM_code & Chr(39)
    End If
    
    sql_ret = QSqlDBExec(SqlDoc, QsqlCod2)
    If sql_ret <> QSQL_SUCCESS Then
        sql_ret = QSqlSelectFree(QsqlCod2)
        sql_ret = Qsqlclose(QsqlCod2, ONECLOSE)
        Exit Sub
    End If
    
    If QSqlGetRow(record, QsqlCod2) = QSQL_SUCCESS Then
                   
        QSqlGetField 1, record, patient()
        
        S0COM_name = patient(1)

        S0COM_ret = QSQL_SUCCESS
    Else
        S0COM_ret = -1
    End If

    sql_ret = QSqlSelectFree(QsqlCod2)
    sql_ret = Qsqlclose(QsqlCod2, ONECLOSE)
 
End Sub
'/* 진료과와 주치의 명 얻기...
'/* S0COM_CODE : 진료과코드+주치의코드
Sub S0SUB_SELECT_AC01A10M(frm As Form)
  
    Dim SqlDoc  As String: Dim sql_ret  As Integer
    Dim dept()  As String
    
    S0COM_name = "": S0COM_name1 = ""
    
    sql_ret = S0SUB_Open(S0COM_SERVER04, frm.hWnd, QsqlCod2)
    If sql_ret <> QSQL_SUCCESS Then
        MsgBox "OCS서버 연결 Error!!", 0
        Exit Sub
    End If

    SqlDoc = "SELECT DeptNm, DrNm FROM AC01A_DB..AC01A10M_TBL"
    SqlDoc = SqlDoc + " WHERE DeptCd = " & Chr(39) & Mid$(S0COM_code, 1, 2) & Chr(39)
    SqlDoc = SqlDoc + "   AND DrCd   = " & Chr(39) & Mid$(S0COM_code, 3, 1) & Chr(39)
    sql_ret = QSqlDBExec(SqlDoc, QsqlCod2)
    If sql_ret <> QSQL_SUCCESS Then
        sql_ret = QSqlSelectFree(QsqlCod2)
        sql_ret = Qsqlclose(QsqlCod2, ONECLOSE)
        Exit Sub
    End If
    
    If QSqlGetRow(record, QsqlCod2) = QSQL_SUCCESS Then
                   
        QSqlGetField 2, record, dept()
        
        S0COM_name = dept(1)
        S0COM_name1 = dept(2)

        S0COM_ret = QSQL_SUCCESS
    Else
        S0COM_ret = -1
    End If

    sql_ret = QSqlSelectFree(QsqlCod2)                 ' 코드 구분 column명 set
        
    sql_ret = Qsqlclose(QsqlCod2, ONECLOSE)

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                               *
'*   생년월일로 나이를 계산                      *
'*   passport_id   :  생년월일 변환대상 data     *
'*                                               *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_BIRTHDAY(ByVal PassPort_Id As String) As String

    Dim cDte     As String
    Dim yy       As String
    Dim age      As Integer

    On Error GoTo S0SUB_BIRTHDAY

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
        
    S0SUB_BIRTHDAY = Trim(Str$(age))
        
    On Error GoTo 0
    Exit Function
S0SUB_BIRTHDAY:

    Resume Next
        
End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String값을 currency mode로 변환               *
'*    wf_src1   :  변환 대상 data                *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CCur(wf_src1 As String) As Currency
    On Error GoTo S0SUB_CCUR_ERROR

    If Len(Trim(wf_src1)) = 0 Or IsNull(wf_src1) Then                 ' NULL ?
        wf_src1 = "0"                              ' zero set
    End If
    S0SUB_CCur = CCur(wf_src1)                     ' currency(money) mode 변환
    Exit Function

S0SUB_CCUR_ERROR:
    
    S0SUB_CCur = 0@                                ' error의 경우 zero set
    Exit Function

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String값을 double mode로 변환                 *
'*    wf_src1   :  변환 대상 data                *
'*    slen      :  소수점 이하 자리수            *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CDBL(wf_src1 As String, slen As Integer) As Double
    On Error GoTo S0SUB_CDBL_ERROR

    If Len(Trim(wf_src1)) = 0 Then               ' NULL ?
       If slen = 0 Then                          ' 소수점 이하 없음 ?
          wf_src1 = "0"
       ElseIf slen = 1 Then                      ' 소수점 이하 1 자리 ?
          wf_src1 = "0.0"
       ElseIf slen = 2 Then                      ' 소수점 이하 2 자리 ?
          wf_src1 = "0.00"
       ElseIf slen = 3 Then                      ' 소수점 이하 3 자리 ?
          wf_src1 = "0.000"
       ElseIf slen = 4 Then                      ' 소수점 이하 4 자리 ?
          wf_src1 = "0.0000"
       ElseIf slen = 5 Then                      ' 소수점 이하 5 자리 ?
          wf_src1 = "0.00000"
       End If
    End If
    S0SUB_CDBL = CDbl(wf_src1)                   ' double mode 변환
    Exit Function

S0SUB_CDBL_ERROR:
    
    S0SUB_CDBL = 0#                              ' error 의경우 zero set
    Exit Function

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* Table 로부터 코드명칭을 SELECT 하여           *
'*                       명칭 Control 에 표시    *
'*    wf_src1   :  변환 대상 data                *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Sub S0SUB_CDNAME_GET()


    Dim ret%, tStr$
    Dim nCols As Integer
    Dim record As String
    Dim SData() As String

    SqlStr = "SELECT " & S0COM_name_col & " FROM " & S0COM_table
    SqlStr = SqlStr & " WHERE " & S0COM_code_col & " = '" & Trim(S0COM_code) & "' "

    ret = QSqlDBExec(SqlStr, QsqlCode): If ret <> QSQL_SUCCESS Then GoTo QsqlFail
    ret = QSqlGetRow(record, QsqlCode): If ret <> QSQL_SUCCESS Then GoTo QsqlFail

    QSqlGetField 1, record, SData()
    
    S0COM_name = S0SUB_RSPACE(SData(1), S0COM_length)
    S0COM_ret = ret                                 ' QSql의 Return값 Setting
    ret% = QSqlSelectFree(QsqlCode)                 ' 코드 구분 column명 set

    Exit Sub

QsqlFail:
    S0COM_name = ""                                 ' Error시 Null Return
    S0COM_ret = ret                                 ' QSql의 Return값 Setting
    ret% = QSqlSelectFree(QsqlCode)                 ' 코드 구분 column 명 set

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String값을 integer mode로 변환                *
'*    wf_src1   :  변환 대상 data                *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CINT(wf_src1 As String) As Integer
    On Error GoTo S0SUB_CINT_ERROR

    If Len(Trim(wf_src1)) = 0 Then                 ' NULL ?
        wf_src1 = "0"                              ' zero set
    End If
    S0SUB_CINT = CInt(wf_src1)                     ' integer mode 변환
    Exit Function

S0SUB_CINT_ERROR:
    
    S0SUB_CINT = 0                                 ' error 의경우 zero set
    Exit Function

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String값을 long integer mode로 변환           *
'*    wf_src1   :  변환 대상 data                *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CLng(wf_src1 As String) As Long
    On Error GoTo S0SUB_CLNG_ERROR

    If Len(Trim(wf_src1)) = 0 Then                 ' NULL ?
        wf_src1 = "0"                              ' zero set
    End If
    S0SUB_CLng = CLng(wf_src1)                     ' long integer mode 변환
    Exit Function

S0SUB_CLNG_ERROR:
    
    S0SUB_CLng = 0&                                ' error 의경우 zero set
    Exit Function

End Function

Function S0SUB_COM_NAME()

    Dim pos As Integer
    Dim FName, FNum, line_str    ' Declare variables.

    On Error Resume Next

    FName = "C:\LANMAN.DOS\LANMAN.INI"
    Open FName For Input As #1   ' Open file.
    If Err <> 0 Then
        Beep
        MsgBox Trim(Err) & " : " & Error & "(" & FName & ")"
        S0SUB_COM_NAME = "NOTID"
        Close #1
        Exit Function
    End If

    Do While Not EOF(1)
        Input #1, line_str

        pos = InStr(line_str, "computername = ")

        If pos <> 0 Then
            S0SUB_COM_NAME = Mid(line_str, pos + 15, Len(line_str))
            Exit Do
        End If
    Loop

    Close #1  ' Close all files.

End Function

Sub S0SUB_CREATE_CONFIG(para As String)
   
    Dim ConDB           As Database
    ReDim ConTB(1 To 7) As New TableDef
    ReDim ConFD(1 To 7) As New Field
    ReDim FD1(1 To 3) As New Field
    ReDim FD2(1 To 3) As New Field
    ReDim FD3(1 To 3) As New Field
    ReDim FD4(1 To 3) As New Field
    ReDim FD5(1 To 3) As New Field
    ReDim FD6(1 To 3) As New Field
    ReDim Que(1 To 6) As QueryDef
    ReDim idx(1 To 7) As New Index
    
    Dim i       As Integer

    Set ConDB = CreateDatabase(para, DB_LANG_GENERAL)
    
'검사기기 환경 테이블
'-------------------------------------------
    ConTB(1).name = "CONFIG_TBL"
    
    ConFD(1).name = "NAME": ConFD(1).Type = DB_TEXT: ConFD(1).Size = 20
    ConTB(1).Fields.Append ConFD(1)
    
    ConFD(2).name = "PORTNO": ConFD(2).Type = DB_TEXT: ConFD(2).Size = 1
    ConTB(1).Fields.Append ConFD(2)
    
    ConFD(3).name = "BAUDRA": ConFD(3).Type = DB_TEXT: ConFD(3).Size = 4
    ConTB(1).Fields.Append ConFD(3)
    
    ConFD(4).name = "DATABIT": ConFD(4).Type = DB_TEXT: ConFD(4).Size = 1
    ConTB(1).Fields.Append ConFD(4)
    
    ConFD(5).name = "STOPBIT": ConFD(5).Type = DB_TEXT: ConFD(5).Size = 1
    ConTB(1).Fields.Append ConFD(5)
    
    ConFD(6).name = "PARITY": ConFD(6).Type = DB_TEXT: ConFD(6).Size = 1
    ConTB(1).Fields.Append ConFD(6)
    
    ConFD(7).name = "SAVEDAT": ConFD(7).Type = DB_TEXT: ConFD(7).Size = 2
    ConTB(1).Fields.Append ConFD(7)
        
    idx(1).name = "PrimaryKey"
    idx(1).Fields = "NAME"
    idx(1).Primary = True
    ConTB(1).Indexes.Append idx(1)
    
    ConDB.TableDefs.Append ConTB(1)

'Hitachi 검사항목 설정
'-------------------------------------------
    ConTB(2).name = "HITAC_TBL"
    
    FD1(1).name = "CODE": FD1(1).Type = DB_TEXT: FD1(1).Size = 3
    ConTB(2).Fields.Append FD1(1)
    
    FD1(2).name = "EXAM": FD1(2).Type = DB_TEXT: FD1(2).Size = 10
    ConTB(2).Fields.Append FD1(2)
    
    FD1(3).name = "NAME": FD1(3).Type = DB_TEXT: FD1(3).Size = 10
    ConTB(2).Fields.Append FD1(3)
    
    idx(2).name = "PrimaryKey"
    idx(2).Fields = "CODE"
    idx(2).Primary = True
    ConTB(2).Indexes.Append idx(2)
    
    ConDB.TableDefs.Append ConTB(2)
    
'KODAK 검사항목 설정
'-------------------------------------------
    ConTB(3).name = "KODAK_TBL"
    
    FD2(1).name = "CODE": FD2(1).Type = DB_TEXT: FD2(1).Size = 3
    ConTB(3).Fields.Append FD2(1)
    
    FD2(2).name = "EXAM": FD2(2).Type = DB_TEXT: FD2(2).Size = 10
    ConTB(3).Fields.Append FD2(2)
    
    FD2(3).name = "NAME": FD2(3).Type = DB_TEXT: FD2(3).Size = 10
    ConTB(3).Fields.Append FD2(3)
    
    idx(3).name = "PrimaryKey"
    idx(3).Fields = "CODE"
    idx(3).Primary = True
    ConTB(3).Indexes.Append idx(3)
    
    ConDB.TableDefs.Append ConTB(3)
    
'STRATUS 1 검사항목 설정
'-------------------------------------------
    ConTB(4).name = "STUR1_TBL"
    
    FD3(1).name = "CODE": FD3(1).Type = DB_TEXT: FD3(1).Size = 3
    ConTB(4).Fields.Append FD3(1)
    
    FD3(2).name = "EXAM": FD3(2).Type = DB_TEXT: FD3(2).Size = 10
    ConTB(4).Fields.Append FD3(2)
    
    FD3(3).name = "NAME": FD3(3).Type = DB_TEXT: FD3(3).Size = 10
    ConTB(4).Fields.Append FD3(3)
    
    idx(4).name = "PrimaryKey"
    idx(4).Fields = "CODE"
    idx(4).Primary = True
    ConTB(4).Indexes.Append idx(4)
    
    ConDB.TableDefs.Append ConTB(4)
    
'STRATUS 2 검사항목 설정
'-------------------------------------------
    ConTB(5).name = "STUR2_TBL"
    
    FD4(1).name = "CODE": FD4(1).Type = DB_TEXT: FD4(1).Size = 3
    ConTB(5).Fields.Append FD4(1)
    
    FD4(2).name = "EXAM": FD4(2).Type = DB_TEXT: FD4(2).Size = 10
    ConTB(5).Fields.Append FD4(2)
    
    FD4(3).name = "NAME": FD4(3).Type = DB_TEXT: FD4(3).Size = 10
    ConTB(5).Fields.Append FD4(3)
    
    idx(5).name = "PrimaryKey"
    idx(5).Fields = "CODE"
    idx(5).Primary = True
    ConTB(5).Indexes.Append idx(5)
    
    ConDB.TableDefs.Append ConTB(5)
    
'Coulter STKS 검사항목 설정
'-------------------------------------------
    ConTB(6).name = "CSTKS_TBL"
    
    FD5(1).name = "CODE": FD5(1).Type = DB_TEXT: FD5(1).Size = 3
    ConTB(6).Fields.Append FD5(1)
    
    FD5(2).name = "EXAM": FD5(2).Type = DB_TEXT: FD5(2).Size = 10
    ConTB(6).Fields.Append FD5(2)
    
    FD5(3).name = "NAME": FD5(3).Type = DB_TEXT: FD5(3).Size = 10
    ConTB(6).Fields.Append FD5(3)
    
    idx(6).name = "PrimaryKey"
    idx(6).Fields = "CODE"
    idx(6).Primary = True
    ConTB(6).Indexes.Append idx(6)
    
    ConDB.TableDefs.Append ConTB(6)
    
'Coulter T-540 검사항목 설정
'-------------------------------------------
    ConTB(7).name = "CT540_TBL"
    
    FD6(1).name = "CODE": FD6(1).Type = DB_TEXT: FD6(1).Size = 3
    ConTB(7).Fields.Append FD6(1)
    
    FD6(2).name = "EXAM": FD6(2).Type = DB_TEXT: FD6(2).Size = 10
    ConTB(7).Fields.Append FD6(2)
    
    FD6(3).name = "NAME": FD6(3).Type = DB_TEXT: FD6(3).Size = 10
    ConTB(7).Fields.Append FD6(3)
    
    idx(7).name = "PrimaryKey"
    idx(7).Fields = "CODE"
    idx(7).Primary = True
    ConTB(7).Indexes.Append idx(7)
    
    ConDB.TableDefs.Append ConTB(7)
    
    Set Que(1) = ConDB.CreateQueryDef("DEL_HITAC", "DELETE FROM HITAC_TBL")
    Set Que(2) = ConDB.CreateQueryDef("DEL_KODAK", "DELETE FROM KODAK_TBL")
    Set Que(3) = ConDB.CreateQueryDef("DEL_STUR1", "DELETE FROM STUR1_TBL")
    Set Que(4) = ConDB.CreateQueryDef("DEL_STUR2", "DELETE FROM STUR2_TBL")
    Set Que(5) = ConDB.CreateQueryDef("DEL_CSTKS", "DELETE FROM CSTKS_TBL")
    Set Que(6) = ConDB.CreateQueryDef("DEL_CT540", "DELETE FROM CT540_TBL")
    
    Que(1).Close: Que(2).Close: Que(3).Close
    Que(4).Close: Que(5).Close: Que(6).Close
    
    ConDB.Close

End Sub

'****************************************************
'*                                                  *
'*  Hitachi Interface 결과받기 테이블 생성          *
'*  DB NAME = para                                  *
'*  TABLE NAME = test, receive_seq                  *
'*                                                  *
'****************************************************
Sub S0SUB_CREATE_RET06DB(para As String)
   
    Dim RetDB           As Database
    Dim RetTB           As New TableDef
    Dim SEQNO           As New Field
    Dim GEMGB           As New Field
    Dim CHECK           As New Field
    ReDim Item(1 To 32) As New Field
    Dim idx             As New Index
    
    Dim i       As Integer

    Set RetDB = CreateDatabase(para, DB_LANG_GENERAL)
    
    RetTB.name = "RET060_TBL"
    
     '검사구분 : 1=본사 종합건강진단서비스 2=계약자서비스, 고계약자서비스
    GEMGB.name = "TotalTest": GEMGB.Type = DB_TEXT: GEMGB.Size = 1: RetTB.Fields.Append GEMGB
    SEQNO.name = "DeviceSno": SEQNO.Type = DB_TEXT: SEQNO.Size = 5: RetTB.Fields.Append SEQNO
        
    For i = 1 To 32
        Item(i).name = "ITEM" + Trim$(Str$(i))
        Item(i).Type = DB_TEXT
        Item(i).Size = 7
        RetTB.Fields.Append Item(i)
    Next
    
    CHECK.name = "RegChk": CHECK.Type = DB_BOOLEAN: RetTB.Fields.Append CHECK
    
    idx.name = "PrimaryKey"
    idx.Fields = "TotalTest;DeviceSno"
    idx.Primary = True
    RetTB.Indexes.Append idx
    
    RetDB.TableDefs.Append RetTB
    
    RetDB.Close

End Sub

'********************************************************
'*                                                      *
'*  Kodak 결과 테이블 생성                              *
'*  para      : 파일명                                  *
'*  ExamCount : 검사항목 갯수                           *
'*                                                      *
'********************************************************
Sub S0SUB_CREATE_RET07DB(para As String)
   
    Dim RetDB   As Database
    Dim RetTB   As New TableDef
    
    Dim SEQNO   As New Field
    Dim GEMGB   As New Field
    Dim CHECK   As New Field
    ReDim Item(1 To 18) As New Field
    Dim idx     As New Index
    
    Dim i       As Integer

    Set RetDB = CreateDatabase(para, DB_LANG_GENERAL)
    
    RetTB.name = "RET070_TBL"
    
     '검사구분 : 1=본사 종합건강진단서비스 2=계약자서비스, 고계약자서비스
    GEMGB.name = "GEMGB": GEMGB.Type = DB_TEXT: GEMGB.Size = 1: RetTB.Fields.Append GEMGB
    SEQNO.name = "SEQNO": SEQNO.Type = DB_TEXT: SEQNO.Size = 5: RetTB.Fields.Append SEQNO
    CHECK.name = "CHECK": CHECK.Type = DB_BOOLEAN: RetTB.Fields.Append CHECK
        
    For i = 1 To 18
        Item(i).name = "ITEM" + Format$(i, "00")
        Item(i).Type = DB_TEXT
        Item(i).Size = 8
        RetTB.Fields.Append Item(i)
    Next
    
    idx.name = "PrimaryKey"
    idx.Fields = "GEMGB;SEQNO"
    idx.Primary = True
    RetTB.Indexes.Append idx
    
    RetDB.TableDefs.Append RetTB
    
    RetDB.Close

End Sub

'****************************************************
'*                                                  *
'*  Stratus 1,2 Interface 결과받기 테이블 생성      *
'*  DB NAME = para                                  *
'*  TABLE NAME = test, receive_seq                  *
'*                                                  *
'****************************************************
Sub S0SUB_CREATE_RET08DB(para As String)
   
    Dim RetDB           As Database
    Dim RetTB           As New TableDef
    Dim SEQNO           As New Field
    Dim GEMGB           As New Field
    Dim CHECK           As New Field
    ReDim Item(1 To 32) As New Field
    Dim idx             As New Index
    
    Dim i       As Integer

    Set RetDB = CreateDatabase(para, DB_LANG_GENERAL)
    
    RetTB.name = "RET080_TBL"
    
     '검사구분 : 1=본사 종합건강진단서비스 2=계약자서비스, 고계약자서비스
    GEMGB.name = "TotalTest": GEMGB.Type = DB_TEXT: GEMGB.Size = 1: RetTB.Fields.Append GEMGB
    SEQNO.name = "DeviceSno": SEQNO.Type = DB_TEXT: SEQNO.Size = 5: RetTB.Fields.Append SEQNO
        
    For i = 1 To 16
        Item(i).name = "ITEM" + Trim$(Str$(i))
        Item(i).Type = DB_TEXT
        Item(i).Size = 7
        RetTB.Fields.Append Item(i)
    Next
    
    CHECK.name = "RegChk": CHECK.Type = DB_BOOLEAN: RetTB.Fields.Append CHECK
    
    idx.name = "PrimaryKey"
    idx.Fields = "TotalTest;DeviceSno"
    idx.Primary = True
    RetTB.Indexes.Append idx
    
    RetDB.TableDefs.Append RetTB
    
    RetDB.Close

End Sub

'******************************************
'*  Coulter STKS
'*******************************************
Sub S0SUB_CREATE_RET09DB(para As String)
   
    Dim RetDB           As Database
    Dim RetTB           As New TableDef
    Dim SEQNO           As New Field
    Dim GEMGB           As New Field
    Dim CHECK           As New Field
    ReDim Item(1 To 32) As New Field
    Dim idx             As New Index
    
    Dim i       As Integer

    Set RetDB = CreateDatabase(para, DB_LANG_GENERAL)
    
    RetTB.name = "RET090_TBL"
    
     '검사구분 : 1=본사 종합건강진단서비스 2=계약자서비스, 고계약자서비스
    GEMGB.name = "TotalTest": GEMGB.Type = DB_TEXT: GEMGB.Size = 1: RetTB.Fields.Append GEMGB
    SEQNO.name = "DeviceSno": SEQNO.Type = DB_TEXT: SEQNO.Size = 5: RetTB.Fields.Append SEQNO

    For i = 1 To 30
        Item(i).name = "ITEM" + Trim$(Str$(i))
        Item(i).Type = DB_TEXT
        Item(i).Size = 7
        RetTB.Fields.Append Item(i)
    Next
    
    CHECK.name = "RegChk": CHECK.Type = DB_BOOLEAN: RetTB.Fields.Append CHECK
    
    idx.name = "PrimaryKey"
    idx.Fields = "TotalTest;DeviceSno"
    idx.Primary = True
    RetTB.Indexes.Append idx
    
    RetDB.TableDefs.Append RetTB
    
    RetDB.Close

End Sub

Sub S0SUB_CREATE_RET10DB(para As String)
   
    Dim RetDB           As Database
    Dim RetTB           As New TableDef
    Dim SEQNO           As New Field
    Dim GEMGB           As New Field
    Dim CHECK           As New Field
    ReDim Item(1 To 32) As New Field
    Dim idx             As New Index
    
    Dim i       As Integer

    Set RetDB = CreateDatabase(para, DB_LANG_GENERAL)
    
    RetTB.name = "RET100_TBL"
    
     '검사구분 : 1=본사 종합건강진단서비스 2=계약자서비스, 고계약자서비스
    GEMGB.name = "TotalTest": GEMGB.Type = DB_TEXT: GEMGB.Size = 1: RetTB.Fields.Append GEMGB
    SEQNO.name = "DeviceSno": SEQNO.Type = DB_TEXT: SEQNO.Size = 5: RetTB.Fields.Append SEQNO

    For i = 1 To 30
        Item(i).name = "ITEM" + Trim$(Str$(i))
        Item(i).Type = DB_TEXT
        Item(i).Size = 7
        RetTB.Fields.Append Item(i)
    Next
    
    CHECK.name = "RegChk": CHECK.Type = DB_BOOLEAN: RetTB.Fields.Append CHECK
    
    idx.name = "PrimaryKey"
    idx.Fields = "TotalTest;DeviceSno"
    idx.Primary = True
    RetTB.Indexes.Append idx
    
    RetDB.TableDefs.Append RetTB
    
    RetDB.Close

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String값을 single mode로 변환                 *
'*    wf_src1   :  변환 대상 data                *
'*    slen      :  소수점 이하 자리수            *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CSNG(wf_src1 As String, slen As Integer) As Single
    On Error GoTo S0SUB_CSNG_ERROR

    If Len(Trim(wf_src1)) = 0 Then               ' NULL ?
       If slen = 0 Then                          ' 소수점 이하 없음 ?
          wf_src1 = "0"
       ElseIf slen = 1 Then                      ' 소수점 이하 1 자리 ?
          wf_src1 = "0.0"
       ElseIf slen = 2 Then                      ' 소수점 이하 2 자리 ?
          wf_src1 = "0.00"
       ElseIf slen = 3 Then                      ' 소수점 이하 3 자리 ?
          wf_src1 = "0.000"
       ElseIf slen = 4 Then                      ' 소수점 이하 4 자리 ?
          wf_src1 = "0.0000"
       End If
    End If
    S0SUB_CSNG = CDbl(wf_src1)                   ' single mode 변환
    Exit Function

S0SUB_CSNG_ERROR:
    
    S0SUB_CSNG = 0!                              ' error 의경우 zero set
    Exit Function

End Function

'*-------------------------------------------------*
'* req      : "1" ==> yyyymmdd return              *
'*          : "2" ==> yyyy-mm-dd return            *
'* para     : value                    (i)         *
'*-------------------------------------------------*
Function S0SUB_DATE_6TO8(ByVal para As String, ByVal req As String) As String

    Dim temp As String

    temp = Left(para, 2)                    '년도 저장
    If temp > "70" And temp <= "99" Then
        temp = "19" & para
    Else
        temp = "20" & para
    End If

    If req = "1" Then
        S0SUB_DATE_6TO8 = temp
    Else
        S0SUB_DATE_6TO8 = Format$(temp, "####-##-##")
    End If

End Function

'*-------------------------------------------------*
'* 일자 check 및 format 처리                       *
'* req      : "1" ==> yyyy check                   *
'*          : "2" ==> yyyymm check                 *
'*          : "3" ==> yyyymmdd check               *
'* delimeter: "/" or "-" or "."        (i)         *
'* date1    : yyyymmdd                 (i/o)       *
'* date2    : yyyy-mm-dd               (i/o)       *
'* iret     : 1 ==> succeed            (o)         *
'*          : -1 ==> parameter error   (o)         *
'*          : -2 ==> 일자 error        (o)         *
'*-------------------------------------------------*
Sub S0SUB_DATE_FORMAT(req As String, delimeter As String, Date1 As String, date2 As String, iRet As Integer)

    Dim slen%

    iRet = 1                                     ' 정상 처리 set
    If (req = "1") And (delimeter <> "/" And delimeter <> "-" And delimeter <> ".") Then
        iRet = -1                                ' parameter error
        Exit Sub
    End If
    Select Case Left$(Trim(req), 1)              ' req parameter check
        Case "1"
            If IsNumeric(Date1) And Len(Date1) = 4 Then
                If (S0SUB_CINT(Date1) < 1990) Or (S0SUB_CINT(Date1) > 2100) Then
                    iRet = -2
                    Exit Sub
                End If
                Exit Sub
            Else
                iRet = -2
                Exit Sub
            End If
        Case "2"
            If IsNumeric(Date1) And Len(Date1) = 6 Then
                If (S0SUB_CINT(Left$(Date1, 4)) < 1990) Or (S0SUB_CINT(Left$(Date1, 4)) > 2100) Then
                    iRet = -2
                    Exit Sub
                End If
                If (S0SUB_CINT(Mid$(Date1, 5, 2)) < 0) Or (S0SUB_CINT(Mid$(Date1, 5, 2)) > 12) Then
                    iRet = -2
                    Exit Sub
                End If
                Exit Sub
            Else
                iRet = -2
                Exit Sub
            End If
        Case "3"
            If IsNumeric(Date1) And Len(Date1) = 8 Then
                If IsDate(Left$(Date1, 4) + " " + Mid$(Date1, 5, 2) + " " + Right$(Date1, 2)) = False Then
                    iRet = -2
                    Exit Sub
                End If
                Exit Sub
            Else
                iRet = -2
                Exit Sub
            End If
        Case Else
            iRet = -1                            ' parameter error
            Exit Sub
    End Select

    If IsDate(date2) = False Then                ' 일자 check error ?
        iRet = -2                                ' 일자 error
        Exit Sub
    End If

End Sub

'********************************************************
'*                                                      *
'*  저장기간이 지난 interface file를 지운다.            *
'*  DirName  : 지우고자 하는 Directory Name             *
'*  SaveDate : 저장기간                                 *
'*                                                      *
'********************************************************
Sub S0SUB_DELETEFILE(DirName As String, SaveDate As String)

    Dim DelFile As String
    Dim FileName    As String
    Dim FileDate    As String

    Const ATTR_NORMAL = 0

    On Error GoTo S0SUB_DELETEFILE

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
S0SUB_DELETEFILE:
    
    Resume Next

End Sub

Sub S0SUB_DISPLAY_PART(ctr As Control)

    Dim SqlDoc  As String, sql_ret As Integer
    Dim record  As String
    Dim SqlData()   As String
    
    SqlDoc = "SELECT DISTINCT PARTCD, SLIPNM"
    SqlDoc = SqlDoc + "  FROM LAB01_DB..SLA010M"
    SqlDoc = SqlDoc + " WHERE SLIPCD = " & Chr(39) & "" & Chr(39)
    
    ctr.Clear
    
    ret = QSqlDBExec(SqlDoc, QsqlCode)
    If ret = QSQL_SUCCESS Then
        Do Until QSqlGetRow(record, QsqlCode) <> QSQL_SUCCESS
            
            QSqlGetField 2, record, SqlData()
            ctr.AddItem SqlData(1) + "  " + SqlData(2)

        Loop
    End If

    ret% = QSqlSelectFree(QsqlCode)                 ' 코드 구분 column명 set
    
    'ctr.ListIndex = 0

End Sub

'*  *   *   *   *   *   *   *   *   *   *   *
'*                                          *
'*  파일의 존재여부을 파악한다.             *
'*  para : 파일명(경로명 포함)              *
'*  Return Value : true, false              *
'*                                          *
'*  *   *   *   *   *   *   *   *   *   *   *
Function S0SUB_EXIST_FILE(para As String) As Integer

    If Dir$(para) <> "" Then
        S0SUB_EXIST_FILE = True
    Else
        S0SUB_EXIST_FILE = False
    End If

End Function

'*------------------------------------------------------*
'*                                                      *
'*  Record의 존재하면 True, 아니면 False                *
'*  para : SQL 문                                       *
'*                                                      *
'*------------------------------------------------------*
Function S0SUB_EXIST_RECORD(para As String) As Integer
    
    Dim status  As Integer
    Dim Row     As Integer
    Dim sStr    As String
    Dim tData() As String

    status = QSqlDBExec(para, QsqlCode)
    If status <> QSQL_SUCCESS Then
        status = QSqlSelectFree(QsqlCode)
        S0SUB_EXIST_RECORD = False
        Exit Function
    End If

    status = QSqlGetRow(sStr, QsqlCode)
    If status <> QSQL_SUCCESS Then
        status = QSqlSelectFree(QsqlCode)
        S0SUB_EXIST_RECORD = False
        Exit Function
    End If

    QSqlGetField 1, sStr, tData()

    If Val(tData(1)) = 0 Then
        S0SUB_EXIST_RECORD = False
    Else
        S0SUB_EXIST_RECORD = True
    End If
    
    status = QSqlSelectFree(QsqlCode)

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*  수치 영역의 값을 control 왼쪽 1 col.부터 표시*
'*    wf_src1   :  대상 변수                     *
'*    slen      :  소수점이하 자리수 (double경우)*
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_FCDBL(wf_src1 As String, slen As Integer) As String
    On Error GoTo S0SUB_FCDBL_ERROR

    If Len(Trim(wf_src1)) = 0 Then               ' NULL ?
       If slen = 0 Then
          wf_src1 = "0"
       ElseIf slen = 1 Then
          wf_src1 = "0.0"
       ElseIf slen = 2 Then
          wf_src1 = "0.00"
       ElseIf slen = 3 Then
          wf_src1 = "0.000"
       ElseIf slen = 4 Then
          wf_src1 = "0.0000"
       ElseIf slen = 5 Then
          wf_src1 = "0.00000"
       End If
    End If

    If slen = 0 Then
       S0SUB_FCDBL = RTrim(Format(CDbl(wf_src1), "0"))
    ElseIf slen = 1 Then
       S0SUB_FCDBL = RTrim(Format(CDbl(wf_src1), "0.0"))
    ElseIf slen = 2 Then
       S0SUB_FCDBL = RTrim(Format(CDbl(wf_src1), "0.00"))
    ElseIf slen = 3 Then
       S0SUB_FCDBL = RTrim(Format(CDbl(wf_src1), "0.000"))
    ElseIf slen = 4 Then
       S0SUB_FCDBL = RTrim(Format(CDbl(wf_src1), "0.0000"))
    ElseIf slen = 5 Then
       S0SUB_FCDBL = RTrim(Format(CDbl(wf_src1), "0.00000"))
    End If

    Exit Function

S0SUB_FCDBL_ERROR:
    
    S0SUB_FCDBL = "0"
    Exit Function

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*  수치 영역의 값을 control 왼쪽 1 col.부터 표시*
'*    wf_src1   :  대상 변수                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_FCDBL2(ByVal wf_src1 As String) As String

    Dim pos As Integer
    Dim dbl_para As String

    If Trim(wf_src1) = "" Then
        S0SUB_FCDBL2 = ""
        Exit Function
    End If

    pos = InStr(Trim(wf_src1), ".")
    If pos = 0 Then                                 '정수인경우
        S0SUB_FCDBL2 = RTrim(Format(CDbl(wf_src1), "0"))
        Exit Function
    End If

    dbl_para = Trim(Mid(wf_src1, pos + 1, Len(Trim(wf_src1))))
    If dbl_para = "" Or Val(dbl_para) = 0 Then
        S0SUB_FCDBL2 = RTrim(Format(CDbl(wf_src1), "0"))
        Exit Function
    End If

    Select Case Len(dbl_para)
       Case 1: S0SUB_FCDBL2 = RTrim(Format(CDbl(wf_src1), "0.0"))
       Case 2: S0SUB_FCDBL2 = RTrim(Format(CDbl(wf_src1), "0.00"))
       Case 3: S0SUB_FCDBL2 = RTrim(Format(CDbl(wf_src1), "0.000"))
       Case 4: S0SUB_FCDBL2 = RTrim(Format(CDbl(wf_src1), "0.0000"))
       Case 5: S0SUB_FCDBL2 = RTrim(Format(CDbl(wf_src1), "0.00000"))
    End Select

End Function

'****************************************************
'*                                                  *
'*  자리수를 참고치 자리수 만큼 변환                *
'*  para : 변환하고자 하는 값                       *
'*  defValue : 참고치 값                            *
'*                                                  *
'****************************************************
Function S0SUB_FORMAT(para As String, defValue As String) As String

    Dim strLen      As Integer: Dim decilen As Integer
    Dim strForamt   As String:  Dim i As Integer

    strLen = Len(Trim$(defValue))
    decilen = InStr(defValue, ".")

    strForamt = ""
    For i = 1 To decilen
        strForamt = strForamt + "#"
    Next i
    strForamt = Left$(strForamt, decilen - 1) + "0"

    If decilen = 0 Then
        S0SUB_FORMAT = Format$(Val(para), strForamt)
    Else
        strForamt = strForamt + "."
        For i = 1 To strLen - decilen
            strForamt = strForamt + "0"
        Next
        
        S0SUB_FORMAT = Format$(Val(para), strForamt)
    End If

End Function

'****************************************************
'*                                                  *
'*  배지코드/배지명을 찾는다.                       *
'*  para    : 찾고자하는 값                         *
'*  chk     : 1 = 배지코드, 2 = 배지명              *
'*  revalue : return value                          *
'*                                                  *
'****************************************************
Function S0SUB_GETCLTCODEPOS(para As String, chk As Integer, revalue As String) As Integer

    Dim pos As Integer

    For pos = 1 To S0COM_CLTCOUNT
        If chk = 1 Then
            If Trim$(para) = Left$(S0COM_CLTCODE(pos), 1) Then
                revalue = Mid$(S0COM_CLTCODE(pos), 4)
                S0SUB_GETCLTCODEPOS = pos
                Exit Function
            End If
        Else
            If Trim$(para) = Trim$(Mid$(S0COM_CLTCODE(pos), 3)) Then
                revalue = Mid$(S0COM_CLTCODE(pos), 1, 1)
                S0SUB_GETCLTCODEPOS = pos
                Exit Function
            End If
        End If
    Next
    
    S0SUB_GETCLTCODEPOS = 0

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
Function S0SUB_GRIDGETROW(grd As Object, Col As Integer, para As String) As Integer
    
    Dim code  As String

    Dim Row%

    For Row = 1 To grd.Rows - 1
        grd.Row = Row
        grd.Col = Col
        code = grd.Text

        If Trim$(code) = Trim$(para) Then
            S0SUB_GRIDGETROW = Row
            
            grd.HighLight = True                                            '입력라인 반전
            grd.SelStartRow = Row: grd.SelEndRow = Row
            grd.SelStartCol = grd.FixedCols: grd.SelEndCol = grd.Cols - 1
            
            Exit Function
        End If
    Next

    S0SUB_GRIDGETROW = -1

End Function

'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                      *
'*  grid에서 특정행를 화면에 보여준다..                 *
'*  grd : grid Name                                     *
'*  Row : Row                                           *
'*  disRow : 화면에 보여줄 수 있는 최대 행의 수         *
'*                                                      *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
Sub S0SUB_GRIDTOPROW(grd As Object, Row As Integer, disRow As Integer)
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
'*  SYSTEM LOGON ID를 SETTING                          *
'*  (USER-ID,TEMINAL-ID,CURRENCY-DATE,CURRENCY-TIME)   *
'*                                                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Sub S0SUB_LOGONID()

    Dim sql_ret As Integer
    Dim record  As String, SqlData()    As String
    
    '현재의 SYSTEM일자와 시간을 읽어온다.
    SqlStr = "select convert(char(12),getdate(),102), convert(char(12),getdate(),108)"
    sql_ret = QSqlDBExec(SqlStr, QsqlConn)
    If sql_ret <> QSQL_SUCCESS Then
        sql_ret = QSqlSelectFree(QsqlConn)
        Exit Sub
    End If

    sql_ret = QSqlGetRow(record, QsqlConn)
    If sql_ret <> QSQL_SUCCESS Then
        Beep
        MsgBox QSqlError(sql_ret%)
        sql_ret = QSqlSelectFree(QsqlConn)
        Exit Sub
    End If

    QSqlGetField 2, record, SqlData()

    S0COM_SYSDATE = Mid$(SqlData(1), 1, 4) & Mid$(SqlData(1), 6, 2) & Mid$(SqlData(1), 9, 2)
    S0COM_SYSTIME = Format$(SqlData(2), "HHMMDD")

    'S0COM_LOGINID = S0COM_USERID & "|" & S0COM_TERMID & "|" & S0COM_SYSDATE
    'S0COM_LOGINID = S0COM_LOGONID & "|" & S0COM_SYSTIME

    sql_ret = QSqlSelectFree(QsqlConn)

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String값이 Length보다 짧으면 Left Space 채움 *
'*    길면 자름(한글 마지막 글자 처리)           *
'*    w_text    :  표시 대상 data                *
'*    w_len     :  표시 길이                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_LSPACE(w_text As String, w_len As Integer) As String
    
    Dim s_len As Integer
    Dim ch As Integer, i As Integer, st As Integer

    s_len = Len(w_text)                              ' length 계산
    
    If s_len <= w_len Then                           ' Left SPACE 채움
        S0SUB_LSPACE = Space$(w_len - s_len) + w_text
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
            S0SUB_LSPACE = Left$(w_text, w_len)      ' 마지막 한글 정상 set
        Else                                         ' 마지막 한글 자름
            S0SUB_LSPACE = " " + Left$(w_text, w_len - 1)
        End If
    End If
            
End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*  masked edit에 data type display              *
'*    dest_ctl  :  표시대상 control              *
'*    w_data    :  표시 data                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Sub S0SUB_mask_disp(Dest_Ctl As Control, w_data As String)
    
    Dim w_chk1  As Integer                       ' change flag

    w_chk1 = 0                                   ' change flag reset

    If Dest_Ctl.PromptInclude = True Then        ' control 속성 = true ?
        Dest_Ctl.PromptInclude = False           ' control 속성 false set
        w_chk1 = 1                               ' change flag set
    End If

    If Len(Trim(w_data)) = 0 Then                ' 표시 data NULL ?
        Dest_Ctl.Text = ""                       ' NULL set
    Else
        Dest_Ctl.Text = w_data                   ' 표시 data set
    End If
    
    If w_chk1 = 1 Then                           ' change flag set ?
        Dest_Ctl.PromptInclude = True            ' control 속성 true reset
    End If

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                               *
'*    해당메세지를  선택하여 화면에 표시하기     *
'*    para  :  해당메세지 번호                   *
'*                                               *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_MESSAGE(ByVal para As Integer) As String

    Select Case para
        Case 1: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 2: S0SUB_MESSAGE = "정상적으로 수정되었습니다."
        Case 3: S0SUB_MESSAGE = "정상적으로 삭제되었습니다."
        Case 4: S0SUB_MESSAGE = "정상적으로 조회되었습니다."
        Case 5: S0SUB_MESSAGE = "정상적으로 인쇄되었습니다."
        Case 6: S0SUB_MESSAGE = "해당 자료가 존재하지 않습니다."
        Case 7: S0SUB_MESSAGE = "키값이 변경되었습니다! 확인바랍니다."
        Case 8: S0SUB_MESSAGE = "삭제가 취소되었습니다."
        Case 9: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 10: S0SUB_MESSAGE = """ ' "" 는 입력할 수 없는 문자입니다."
        Case 11: S0SUB_MESSAGE = "항목을 선택하여 주십시오."
        Case 12: S0SUB_MESSAGE = "정상적으로 처리되었습니다.."
        Case 13: S0SUB_MESSAGE = "입력된 USER-ID는 권한이 없으므로 등록할 수 없습니다.."
        Case 14: S0SUB_MESSAGE = "입력된 USER-ID는 권한이 없으므로 수정할 수 없습니다.."


        Case 101: S0SUB_MESSAGE = "선행번호 보다 클수 없읍니다."
        Case 102: S0SUB_MESSAGE = "날짜입력이 틀립니다.  확인하세요."

        Case 103: S0SUB_MESSAGE = "거래선코드를 입력하여 주십시오."
        Case 104: S0SUB_MESSAGE = "거래선코드를 정확하게 입력하여 주십시오."
        Case 105: S0SUB_MESSAGE = "사용 담당자를 입력하여 주십시오."
        Case 106: S0SUB_MESSAGE = "사용 담당자를 정확하게 입력하여 주십시오."
        Case 107: S0SUB_MESSAGE = "거래처 담당자를 입력하여 주십시오."
        Case 108: S0SUB_MESSAGE = "수정시간을 입력하여 주십시오."
        Case 109: S0SUB_MESSAGE = "수정시간을 정확하게 입력하여 주십시오."
        Case 110: S0SUB_MESSAGE = "다중참고치를 입력하여 주십시오."
        Case 111: S0SUB_MESSAGE = "다중참고치를 정확하게 입력하여 주십시오."
        Case 112: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 113: S0SUB_MESSAGE = "항목을 선택하여 주십시오."
        Case 114: S0SUB_MESSAGE = "사용자번호가 변경되었습니다. 확인하십시오."
        Case 115: S0SUB_MESSAGE = "수정, 삭제전에 다중참고치를 삭제 하십시오."
        Case 116: S0SUB_MESSAGE = "수정, 삭제전에 SUB항목을 삭제 하십시오."
        Case 117: S0SUB_MESSAGE = " ""000""는 입력할 수 업습니다."
        Case 118: S0SUB_MESSAGE = "학부코드를 입력하여 주십시오."
        Case 119: S0SUB_MESSAGE = "학부코드를 정확하게 입력하여 주십시오."
        Case 120: S0SUB_MESSAGE = "집단코드를 입력하여 주십시오."
        Case 121: S0SUB_MESSAGE = "집단코드를 정확하게 입력하여 주십시오."
        Case 122: S0SUB_MESSAGE = "검사항목코드를 입력하여 주십시오."
        Case 123: S0SUB_MESSAGE = "검사항목코드를 정확하게 입력하여 주십시오."
        Case 124: S0SUB_MESSAGE = "검사숫가를 입력하여 주십시오."
        Case 125: S0SUB_MESSAGE = "사용자번호를 입력하여 주십시오."
        Case 126: S0SUB_MESSAGE = "사용자명을 입력하여 주십시오."
        Case 127: S0SUB_MESSAGE = "비밀번호를 입력하여 주십시오."
        Case 128: S0SUB_MESSAGE = "비밀번호는 6자리 미만이 될수 없습니다."
        Case 129: S0SUB_MESSAGE = "분소/지사 코드를 입력하여 주십시오."
        Case 130: S0SUB_MESSAGE = "분소/지사 명칭을 입력하여 주십시오."
        Case 131: S0SUB_MESSAGE = "사원담당자를 입력하여 주십시오."
        Case 132: S0SUB_MESSAGE = "사원담당자를 정확하게 입력하여 주십시오."
        Case 133: S0SUB_MESSAGE = "분소담당자를 입력하여 주십시오."
        Case 134: S0SUB_MESSAGE = "수정마감시간을 입력하여 주십시오."
        Case 135: S0SUB_MESSAGE = "숫자를 입력하여 주십시오."
        Case 136: S0SUB_MESSAGE = "거래선명을 입력하여 주십시오."
        Case 137: S0SUB_MESSAGE = "참고문자를 입력하여 주십시오."
        Case 138: S0SUB_MESSAGE = "참고치를 입력하여 주십시오."
        Case 139: S0SUB_MESSAGE = "앞의 값보다 작을수 없습니다."
        Case 140: S0SUB_MESSAGE = "화면상의 자료를 입력또는 수정후 다음 작업을 하십시오."


        Case 201: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 202: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 203: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 204: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 205: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 206: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 207: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 208: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 209: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 210: S0SUB_MESSAGE = "정상적으로 입력되었습니다."


        Case 301: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 302: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 303: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 304: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 305: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 306: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 307: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 308: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 309: S0SUB_MESSAGE = "정상적으로 입력되었습니다."
        Case 310: S0SUB_MESSAGE = "정상적으로 입력되었습니다."


        Case 401: S0SUB_MESSAGE = "소그룹코드를 입력하여 주십시오."
        Case 402: S0SUB_MESSAGE = "소그룹코드를 정확하게 입력하여 주십시오."
        Case 403: S0SUB_MESSAGE = "최대항목을 초과하여 입력할 수 없습니다."
        Case 404: S0SUB_MESSAGE = "검사명을 입력하여 주십시오."
        Case 405: S0SUB_MESSAGE = "참고치를 입력하여 주십시오."
        Case 406: S0SUB_MESSAGE = "앞의 값보다 작을수 없습니다."
        Case 407: S0SUB_MESSAGE = "참고문자를 입력하여 주십시오."
        Case 408: S0SUB_MESSAGE = "검사분류를 입력하여 주십시오."
        Case 409: S0SUB_MESSAGE = "검사분류를 ""1"" ~ ""3"" 중 선택하여 입력하십시오."
        Case 410: S0SUB_MESSAGE = "숫자값를 입력하여 주십시오."
        Case 411: S0SUB_MESSAGE = "대그룹코드는 ""700""번부터 입력이 가능합니다."
        Case 412: S0SUB_MESSAGE = "기능구분코드를 입력하여 주십시오."
        Case 413: S0SUB_MESSAGE = "적용산식코드를 입력하여 주십시오."
        Case 414: S0SUB_MESSAGE = "적용산식코드를 정확하게 입력하여 주십시오."
        Case 415: S0SUB_MESSAGE = "산술식코드를 입력하여 주십시오."
        Case 416: S0SUB_MESSAGE = "CHART 일련번호를 입력하여 주십시오."
        Case 417: S0SUB_MESSAGE = "환자명을 입력하여 주십시오."
        Case 418: S0SUB_MESSAGE = "일자를 입력하여 주십시오."
        Case 419: S0SUB_MESSAGE = "일자를 정확하게 입력하여 주십시오."
        Case 420: S0SUB_MESSAGE = " ""1"" ~ ""2""만 입력이 가능합니다."
        Case 421: S0SUB_MESSAGE = "일련번호를 정확히 입력하여 주십시오."
        Case 422: S0SUB_MESSAGE = "CHART 번호를 입력하여 주십시오."
        Case 423: S0SUB_MESSAGE = "CHART 번호를 정확히 입력하여 주십시오."
        Case 424: S0SUB_MESSAGE = "대그룹코드를 입력하여 주십시오."
        Case 425: S0SUB_MESSAGE = "대그룹코드를 정확히 입력하여 주십시오."
        Case 426: S0SUB_MESSAGE = "금액을 입력하여 주십시오."
        Case 427: S0SUB_MESSAGE = "결제구분을 입력하여 주십시오."
        Case 428: S0SUB_MESSAGE = "접수번호를 정확히 입력하여 주십시오."
        Case 429: S0SUB_MESSAGE = "입력후 부분검사를 입력하여 주십시오."
        Case 430: S0SUB_MESSAGE = "검사항목 입력자료가 없습니다."
        Case 431: S0SUB_MESSAGE = "결과코드를 입력하여 주십시오."
        Case 432: S0SUB_MESSAGE = "결과코드를 정확히 입력하여 주십시오."
        Case 433: S0SUB_MESSAGE = "소견코드를 입력하여 주십시오."
        Case 434: S0SUB_MESSAGE = "소견코드를 정확히 입력하여 주십시오."
        Case 435: S0SUB_MESSAGE = "환자정보가 없습니다. 확인바랍니다!"
        Case 436: S0SUB_MESSAGE = "자료를 입력하여 주십시오."
        Case 437: S0SUB_MESSAGE = "자료를 입력하여 주십시오."
        Case 438: S0SUB_MESSAGE = "출력구분은 ""1"", ""2"", ""3"", ""4"", ""9""를 입력하여 주십시오."
        Case 439: S0SUB_MESSAGE = "소견코드 중복입니다. 확인바랍니다!"
        Case 440: S0SUB_MESSAGE = "종합소견코드 중복입니다. 확인바랍니다!"
        Case 441: S0SUB_MESSAGE = "학부코드를 선택하여 주십시오."
        Case 442: S0SUB_MESSAGE = "검사항목을 선택하여 주십시오."
        Case 443: S0SUB_MESSAGE = "산술식을 입력하여 주십시오."
        Case 444: S0SUB_MESSAGE = "점수를 입력하여 주십시오."
        Case 445: S0SUB_MESSAGE = "주민등록번호를 정확히 입력하여 주십시오."
        Case 446: S0SUB_MESSAGE = "중복된 자료입니다. 확인바랍니다!"
        Case 447: S0SUB_MESSAGE = "검사항목을 1개이상 입력하여 주십시오."
        Case 448: S0SUB_MESSAGE = "문자값를 입력하여 주십시오."

        Case 501: S0SUB_MESSAGE = "분소코드 또는 거래선코드가 없읍니다. "
        Case 502: S0SUB_MESSAGE = "날짜를 확인하세요."
        Case 503: S0SUB_MESSAGE = "수정 오류!!  수정권한이 없는 사용자입니다."
        Case 504: S0SUB_MESSAGE = "삭제 오류!!  삭제권한이 없는 사용자입니다."
        Case 505: S0SUB_MESSAGE = "선행 날짜보다 앞설 수 없습니다."
        Case 506: S0SUB_MESSAGE = "선행 코드보다 앞설 수 없습니다."
        Case 507: S0SUB_MESSAGE = "코드의 유효범위가 아닙니다."
        Case 508: S0SUB_MESSAGE = "해당 분소코드가 존재하지 않습니다.  확인하여 주세요."
        Case 509: S0SUB_MESSAGE = "해당 거래처코드가 존재하지 않습니다.  확인하여 주세요."
        Case 510: S0SUB_MESSAGE = "수금액이 현미수금액보다 클수 없습니다.  확인하여 주세요."
        
        Case 536: S0SUB_MESSAGE = "DB OPEN이 되어있지 않습니다."

    End Select

    
End Function

'****************************************************
'*                                                  *
'*  파일의 존재여부를 파악한다.                     *
'*  para : 파일명                                   *
'*                                                  *
'****************************************************
Function S0SUB_NULL_CHECK(para As Variant) As String

    If IsNull(para) Then
        S0SUB_NULL_CHECK = ""
    Else
        S0SUB_NULL_CHECK = para
    End If

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*  변환 대상 data를 format 형태로 변환          *
'*    src       :  변환 대상 data                *
'*    fmat      :  변환 형태                     *
'*    gbcd      :  변환 구분("T":버림,"R":반올림 *
'*                           "X":올림)           *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_num_format(src As String, fmat As String, gbcd As String) As String
    Dim i       As Integer                       ' for loop index
    Dim j       As Integer                       ' fmat 의 소숫점 이하 자리수
    Dim II      As Integer                       ' src  의 소수점 위치
    Dim JJ      As Integer
    Dim xx      As Integer                       ' src 의 정수자리수
    Dim yy      As Integer                       ' fmat 의 정수자리수

    Dim Text    As String                        ' "," 없는 src data
    Dim maxlen  As Integer                       ' fmat 길이
    Dim maxlen1 As Integer                       ' src  길이
    Dim slen    As Integer                       ' 소수점 이하 자리수
    Dim mid1    As String                        ' return value
    Dim mid2    As Double                        ' src value
    Dim mid3    As Double                        ' src 절대값
    Dim src1    As String                        ' src 의 소수점 이하 data
    Dim chk     As String                        ' 음수 flag
    Dim chk1    As String                        ' 부동소수점 flag
    On Error GoTo FORMAT_ERROR

    src = Trim(src)                              ' NULL cut
    If Len(src) = 0 Then                         ' NULL ?
        src = "0"                                ' zero set
    End If
    maxlen = Len(fmat)                           ' fmat 길이 set
    maxlen1 = Len(src)                           ' src  길이 set
    Text = ""                                    ' return data 초기처리
    For i = 1 To maxlen1                         ' src 에서 "," 를 제외
        If Mid(src, i, 1) = "," Then
        Else
           Text = Text & Mid(src, i, 1)
        End If
    Next

    Text = Trim(Text)

    yy = 0
    For i = 1 To maxlen                          ' fmat에서 "," 를 제외하고
        If Mid(fmat, i, 1) = "." Then            '   정수부분자리수(yy)를 계산
           Exit For
        End If
        If Mid(fmat, i, 1) = "," Then
        Else
           yy = yy + 1
        End If
    Next
    xx = 0
    For i = 1 To maxlen1                         ' src에서 ","를 제외하고
        If Mid(src, i, 1) = "." Then             '   정수부분 자리수(xx)를 계산
           Exit For
        End If
        If Mid(src, i, 1) = "," Then
        Else
           xx = xx + 1
        End If
    Next

    If xx > yy Then                              ' src 정수부 > fmat 정수부 ?
    Else
       mid1 = Format(S0SUB_CDBL(src, 0), fmat)   '   double 변환 후 fmat 변환
    End If

    chk1 = "N"                                   ' 부동소수점 flag reset
    j = 0                                        ' fmat의 소숫점 이하 자리수 reset
    For i = 1 To maxlen
        If chk1 = "Y" Then
           j = j + 1                             ' fmat의 소숫점 이하 자리수 계산
        End If
        If Mid(fmat, i, 1) = "." Then            ' 부동소수점 data ?
           chk1 = "Y"                            ' 부동소수점 flag set
        End If
    Next
                                                 ' src가fmat보다 정수길이가길때 fmat의 자리를기준으로
    If xx > yy Then                              '   src의자료(Text)를 앞에서부터 자른다.
       src = Mid(Text, xx - yy + 1, yy) & "." & Mid(Text, xx + 2, j)
       mid1 = Format(S0SUB_CDBL(src, 0), fmat)
    End If


    slen = j                                     ' 소숫점 이하 자리수

    mid2 = S0SUB_CDBL(src, j)                    ' src double 변환
    If mid2 < 0 Then                             ' 음수 ?
       chk = "Y"                                 '   음수 flag set
    End If
    mid3 = Abs(mid2)                             ' 절대값 계산


    If gbcd = "T" Then                           ' 버림의 경우
        For II = 1 To maxlen1                    '   src의 소수점 위치 계산
            If Mid(src, II, 1) = "." Then
                Exit For
            End If
        Next

        If II > maxlen1 Then GoTo format_m       '   정수부분만 있을때
        src1 = Mid(src, II + 1, maxlen1 - II)    '   소숫점이하 data
        If Val(src1) = 0 Then GoTo format_m      '   소수점이하가 zero 일때
    End If

    If gbcd = "T" And mid3 <> 0 Then             ' 버림의 경우
       If slen = 0 Then                          '   소수점이하 자리수 처리
          mid3 = mid3 - 0.5
       ElseIf slen = 1 Then
          mid3 = mid3 - 0.05
       ElseIf slen = 2 Then
          mid3 = mid3 - 0.005
       ElseIf slen = 3 Then
          mid3 = mid3 - 0.0005
       ElseIf slen = 4 Then
          mid3 = mid3 - 0.00005
       End If
    ElseIf gbcd = "X" Then                       ' 올림의 경우
       If slen = 0 Then                          '   소수점이하 자리수 처리
          mid3 = mid3 + 0.4
       ElseIf slen = 1 Then
          mid3 = mid3 + 0.04
       ElseIf slen = 2 Then
          mid3 = mid3 + 0.004
       ElseIf slen = 3 Then
          mid3 = mid3 + 0.0004
       ElseIf slen = 4 Then
          mid3 = mid3 + 0.00004
       End If
    Else
    End If

format_m:
    mid3 = mid3 + 0.0000000001
    If chk = "Y" Then                            ' 음수의 경우
       mid3 = mid3 * -1                          '   음수 처리
    End If

    mid1 = Format(mid3, fmat)
    mid1 = Space(maxlen - Len(mid1)) & mid1      ' leading space set
FORMAT_END:
    S0SUB_num_format = mid1                      ' return
    Exit Function
FORMAT_ERROR:
    mid1 = "0"                                   ' error 의 경우 zero set
    Beep
    Resume FORMAT_END
End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  S0SUB_Open = DB Open 처리                                   *
'*                                                              *
'*    특정 Index를 지정하여 open할 경우 사용                    *
'*                                                              *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_Open(Server As String, ByVal frmhWnd As Integer, Index As Long)
 
    Dim ret%                    'Qsqlclose RETURN CD

    Index = -1

   'SERVER 연결상태확인
    ret% = QSqlOpen(Server, frmhWnd, Index)

    If ret% <> QSQL_SUCCESS Then
        Beep
        MsgBox QSqlError(ret%)
    End If

    S0SUB_Open = ret%

End Function

' 접수구분
'
Public Function S0SUB_SET_RECCHK(para1 As Integer, para2 As String) As String

    If para1 = 1 Then
    
        If para2 = "1" Then
            S0SUB_SET_RECCHK = "외래"
        ElseIf para2 = "2" Then
            S0SUB_SET_RECCHK = "입원"
        ElseIf para2 = "3" Then
            S0SUB_SET_RECCHK = "응급실"
        Else
            S0SUB_SET_RECCHK = "수탁"
        End If
        
    Else
        If para2 = "외래" Then
            S0SUB_SET_RECCHK = "1"
        ElseIf para2 = "입원" Then
            S0SUB_SET_RECCHK = "2"
        ElseIf para2 = "응급실" Then
            S0SUB_SET_RECCHK = "3"
        End If
        
    End If
         
End Function

' 병원구분
'
'
Public Function S0SUB_SET_HOSCHK(para1 As Integer, para2 As String) As String
    
    If para1 = 1 Then
        
        If para2 = "1" Then
            S0SUB_SET_HOSCHK = "본원"
    
        ElseIf para2 = "2" Then
            S0SUB_SET_HOSCHK = "별관"
        
        ElseIf para2 = "3" Then
            S0SUB_SET_HOSCHK = "심혈관센터"
        
        ElseIf para2 = "4" Then
            S0SUB_SET_HOSCHK = "재활원"
        
        ElseIf para2 = "5" Then
            S0SUB_SET_HOSCHK = "암센터"
        
        ElseIf para2 = "F" Then
            S0SUB_SET_HOSCHK = "안이병원"
        
        ElseIf para2 = "E" Then
            S0SUB_SET_HOSCHK = "응급실"
        
        End If
        
    Else
        
        If para2 = "본원" Then
            S0SUB_SET_HOSCHK = "1"
        
        ElseIf para2 = "별관" Then
            S0SUB_SET_HOSCHK = "2"
        
        ElseIf para2 = "심혈관센터" Then
            S0SUB_SET_HOSCHK = "3"
        
        ElseIf para2 = "재활원" Then
            S0SUB_SET_HOSCHK = "4"
        
        ElseIf para2 = "암센터" Then
            S0SUB_SET_HOSCHK = "5"
        
        ElseIf para2 = "안이병원" Then
            S0SUB_SET_HOSCHK = "F"
        
        ElseIf para2 = "응급실" Then
            S0SUB_SET_HOSCHK = "E"
        
        End If
            
    End If
    
End Function


'*********************************************************
'** 코드/명칭 help 화면 표시위치 계산 및 이동          ***
'** xpos      : help field 의 left position            ***
'** ypos      : help field 의 top position             ***
'** return    : 0:FAIL 1:TRUE                          ***
'*********************************************************
Sub S0SUB_POSITION(frm As Form, xpos As Long, YPos As Long)

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
Function S0SUB_RSPACE(w_text As String, w_len As Integer) As String
    Dim s_len As Integer
    Dim ch As Integer, i As Integer, st As Integer

    s_len = Len(w_text)                              ' length 계산
    
    If s_len <= w_len Then                           ' right SPACE 채움
        S0SUB_RSPACE = w_text + Space$(w_len - s_len)
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
            S0SUB_RSPACE = Left$(w_text, w_len)      ' 마지막 한글 정상 set
        Else                                         ' 마지막 한글 자름
            S0SUB_RSPACE = Left$(w_text, w_len - 1) + " "
        End If
    End If

End Function

Function S0SUB_Len(w_text As String) As String
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
    
    S0SUB_Len = Str(s_len)
    
    
    MsgBox Str(s_len)
    's_len = Len(w_text)                              ' length 계산
    
    'If s_len <= w_len Then                           ' right SPACE 채움
    '    S0SUB_RSPACE = w_text + Space$(w_len - s_len)
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
    '        S0SUB_RSPACE = Left$(w_text, w_len)      ' 마지막 한글 정상 set
    '    Else                                         ' 마지막 한글 자름
    '        S0SUB_RSPACE = Left$(w_text, w_len - 1) + " "
    '    End If
    'End If

End Function

'****************************************************
'*                                                  *
'*  배지코드/배지명을 배열에 넣기                   *
'*                                                  *
'****************************************************
Sub S0SUB_SETCLTCODE()

    S0COM_CLTCODE(1) = "A  BA   "
    S0COM_CLTCODE(2) = "B  CF   "
    S0COM_CLTCODE(3) = "C  TCBS "
    S0COM_CLTCODE(4) = "D  MAC  "
    S0COM_CLTCODE(5) = "E  CHO  "
    S0COM_CLTCODE(6) = "F  TM   "
    S0COM_CLTCODE(7) = "G  SS   "
    S0COM_CLTCODE(8) = "H  PE   "
    S0COM_CLTCODE(9) = "I  Thio "
    S0COM_CLTCODE(10) = "J  SB   "
    S0COM_CLTCODE(11) = "K  Sab  "
    S0COM_CLTCODE(12) = "L  Other"

End Sub

'*----------------------------------------------------------*
'*                                                          *
'*  vaSpread에 데이타을 Clear한다.                       *
'*  spd : vaSpread명,  DispRow : vaSpread에 라인수    *
'*                                                          *
'*----------------------------------------------------------*
Sub S0SUB_Spread_Clear(spd As vaSpread, DispRow As Integer)
    
    spd.Col = 1: spd.col2 = spd.MaxCols
    spd.Row = 1: spd.row2 = spd.MaxRows

    spd.BlockMode = True

    spd.Action = SS_ACTION_CLEAR_TEXT
    spd.MaxRows = DispRow

    spd.BlockMode = False

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
Function S0SUB_SPREADGETCOL(spd As vaSpread, Row As Long, para As Variant) As Integer
    
    Dim code    As String
    Dim Col     As Integer, sp As Boolean

    For Col = 1 To spd.MaxCols
        sp = spd.GetText(Col, Row, CVar(code))
        
        If Trim$(code) = Trim$(para) Then
            S0SUB_SPREADGETCOL = Col
            Exit Function
        End If
    Next

    S0SUB_SPREADGETCOL = -1

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
Function S0SUB_SPREADGETROW(spd As vaSpread, Col As Long, para As String) As Integer
    
    Dim code  As Variant
    Dim Row As Long, sp As Boolean

    For Row = 1 To spd.MaxRows
        
        sp = spd.GetText(Col, Row, code)
        
        If Trim$(code) = Trim$(para) Then
            S0SUB_SPREADGETROW = Row
            Exit Function
        End If
    Next

    S0SUB_SPREADGETROW = -1

End Function

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
Function S0SUB_SPREAD_HIGHLIGHT(spd As vaSpread, Row As Integer, OldRow As Integer) As Integer
    
    spd.Col = 1: spd.col2 = spd.MaxCols
    spd.Row = Val(spd.Tag): spd.row2 = OldRow
    
    spd.BlockMode = True

    spd.BackColor = &HFFFFFF     '흰색

    spd.BlockMode = False
    
    spd.Col = 1: spd.col2 = spd.MaxCols
    spd.Row = Row: spd.row2 = Row
    
    spd.BlockMode = True

    spd.BackColor = &HC0FFC0   '

    spd.BlockMode = False

    S0SUB_SPREAD_HIGHLIGHT = Row

End Function

'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                      *
'*  vaSpread에서 특정행를 화면에 보여준다..          *
'*  spd : vaSpread Name                              *
'*  Row : Row                                           *
'*  disRow : 화면에 보여줄 수 있는 최대 행의 수         *
'*                                                      *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
Sub S0SUB_SPREADTOPROW(spd As vaSpread, Row As Integer, disRow As Integer)
    
    On Error GoTo SpreadTopRowErr

    spd.Col = 1: spd.Row = Row - disRow + 1
    
    spd.Action = SS_ACTION_GOTO_CELL
    
    On Error GoTo 0
    
    Exit Sub

SpreadTopRowErr:
    'spd.Col = 1
    'spd.Row = Row
    'spd.Action = SS_ACTION_GOTO_CELL

    Resume Next

End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'*  DELTA CHECK MODULE                                                                                            *
'*  CurRst : 결과값                                                                                                          *
'*  OrdCd  : 검사코드                                                                                                      *
'*  UnitNo : 진찰권번호                                                                                                   *
'*  CurDate : 결과일자+SYSTIME                                                                                     *
'*  NextDelta : 넘어온 값이 "True"이면 검사하고 다음 결과의 Delta값을 넘겨준다.           *
'*  NextLabNo : 다음결과의 LabDate+SlipCd+LabSqNo 를 넘겨준다.                                 *
'*  TabelName : Select할 결과 Table
'*  < 다음은 "Delta Check"(SLPC110F)를 위함>                                                                *
'*  LstRst : 이전결과값                                                                                                    *
'*  DeltaGbn : Delta Check 방법                                                                                        *
'*  DeltaInterval : 이전결과와 현재결과의 시간간격(Day)(소숫점 2자리까지 값을 구함)      *
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Delta_Check(CurRst As Variant, OrdCd As String, UnitNo As String, CurDate As String, NextDelta As String, NextLabNo As String, Optional TableName As Variant, Optional LstRst As Variant, Optional DeltaGbn As Variant, Optional DeltaInterval As Variant) As String

    Dim c As Integer
    Dim Tmp As Long
    Dim LstDate As String
    Dim NxtDate As String
    Dim NxtRst As Single
    Dim DeltaVal As Single
    Dim DeltaHi As Single
    Dim DeltaLo As Single
    Dim QsqlIndex As Long
    Dim ret%
    Dim SqlData() As String
        
    On Error GoTo Delta_Check_Error
    
    CurRst = Trim(CurRst)
    Delta_Check = ""
    
    If CurRst = "" Then Exit Function
    If IsMissing(DeltaGbn) Then DeltaGbn = ""
    If IsMissing(TableName) Then
        TableName = "LAB01_DB..SLC010M"
        QsqlIndex = QsqlCode
    Else
        QsqlIndex = QsqlCon1
    End If
    
    For c = 1 To Len(CurRst)
        Tmp = Asc(Mid(CurRst, c, 1))
        If Not (Tmp <> 47 And (Tmp >= 46 And Tmp <= 57)) Then Exit Function
    Next
    
    If DeltaGbn = "" Then
    
        SqlStr = "SELECT DISTINCT DLTCHK FROM LAB01_DB..SLA030M " _
                & " WHERE ORDCD = '" & OrdCd & Chr$(39)
                
        ret = QSqlDBExec(SqlStr, QsqlCode)
        If ret = QSQL_SUCCESS Then
            If QSqlGetRow(record, QsqlCode) = QSQL_SUCCESS Then
        
                QSqlGetField 1, record, SqlData()
               
                DeltaGbn = Trim(SqlData(1))
                
            End If
        End If
        ret = QSqlSelectFree(QsqlCode)
    End If
    
    If DeltaGbn <> "" Then
    
        SqlStr = "SELECT DELTAHI, DELTALO  FROM LAB01_DB..SLA031M " _
                & " WHERE ORDCD = '" & OrdCd & Chr$(39) _
                & " AND DLTCHK = '" & DeltaGbn & Chr$(39)
                
        ret = QSqlDBExec(SqlStr, QsqlCode)
        If ret = QSQL_SUCCESS Then
            If QSqlGetRow(record, QsqlCode) = QSQL_SUCCESS Then
        
                QSqlGetField 2, record, SqlData()
               
                DeltaHi = SqlData(1)
                DeltaLo = SqlData(2)
                
           End If
        End If
        ret = QSqlSelectFree(QsqlCode)
                
        '--------------
        '   현재 입력된 결과의 DeltaCheck
        '--------------
        If IsMissing(LstRst) Then
            SqlStr = "SELECT MAX(RSTDATE+SYSTIME) " _
                    & " FROM " & TableName _
                    & " WHERE CSTIDNO = '" & UnitNo & Chr$(39) _
                    & " AND ORDCD = '" & OrdCd & Chr$(39) _
                    & " AND RSTDATE+SYSTIME < '" & CurDate & Chr$(39)
                    
            ret = QSqlDBExec(SqlStr, QsqlIndex)
            If ret = QSQL_SUCCESS Then
                If QSqlGetRow(record, QsqlIndex) = QSQL_SUCCESS Then
                    
                    QSqlGetField 1, record, SqlData()
                    LstDate = SqlData(1)
                    
                End If
            End If
            ret = QSqlSelectFree(QsqlIndex)
            
            If LstDate <> "" Then
            
                DeltaInterval = (DateDiff("n", CDate(Format(LstDate, "0000-00-00 00:00:00")), Format(CurDate, "0000-00-00 00:00:00"))) / 60
                DeltaInterval = Format(DeltaInterval / 24, "0.00")
                
                SqlStr = "SELECT DISTINCT RSTVAL1  " _
                        & " FROM  " & TableName _
                        & " WHERE CSTIDNO = '" & UnitNo & Chr$(39) _
                        & " AND ORDCD = '" & OrdCd & Chr$(39) _
                        & " AND RSTDATE+SYSTIME = '" & LstDate & Chr$(39)
                
                ret = QSqlDBExec(SqlStr, QsqlIndex)
                If ret = QSQL_SUCCESS Then
                    If QSqlGetRow(record, QsqlIndex) = QSQL_SUCCESS Then
                        
                        QSqlGetField 1, record, SqlData()
                        LstRst = SqlData(1)
                    End If
                End If
                ret = QSqlSelectFree(QsqlIndex)
            End If
        End If
        
        
        If Not IsMissing(LstRst) Then
        
            LstRst = Val(LstRst)
            DeltaVal = CurRst - LstRst
            Select Case DeltaGbn
            Case "2"
                DeltaVal = (DeltaVal / LstRst) / 100
            Case "3"
                If DeltaInterval <> 0 Then DeltaVal = DeltaVal / DeltaInterval
            Case "4"
                If DeltaInterval <> 0 Then DeltaVal = ((DeltaVal / CurRst) / 100) / DeltaInterval
            End Select
            
            If DeltaVal > DeltaHi Or DeltaVal < DeltaLo Then
                Delta_Check = "D"
            End If
        End If
        
        If NextDelta = "True" Then
        
            NextDelta = ""
            '---------
            '   다음결과의 DeltaCheck
            '--------
            SqlStr = " SELECT  MIN(RSTDATE+SYSTIME) " _
                    & " FROM " & TableName _
                    & " WHERE CSTIDNO = '" & UnitNo & Chr$(39) _
                    & " AND ORDCD = '" & OrdCd & Chr$(39) _
                    & " AND RSTDATE+SYSTIME > '" & CurDate & Chr$(39)
            
            ret = QSqlDBExec(SqlStr, QsqlIndex)
            If ret = QSQL_SUCCESS Then
                If QSqlGetRow(record, QsqlIndex) = QSQL_SUCCESS Then
                    
                    QSqlGetField 1, record, SqlData()
                    NxtDate = SqlData(1)
                    
                End If
            End If
            ret = QSqlSelectFree(QsqlIndex)
            
            If NxtDate <> "" Then
            
                DeltaInterval = (DateDiff("n", CDate(Format(CurDate, "0000-00-00 00:00:00")), Format(NxtDate, "0000-00-00 00:00:00"))) / 60
                DeltaInterval = Format(DeltaInterval / 24, "0.00")
                
                SqlStr = " SELECT LABDATE+SLIPCD+LABSQNO, RSTVAL1 " _
                        & " FROM " & TableName _
                        & " WHERE CSTIDNO = '" & UnitNo & Chr$(39) _
                        & " AND ORDCD = '" & OrdCd & Chr$(39) _
                        & " AND RSTDATE+SYSTIME ='" & NxtDate & Chr$(39)
                        
                ret = QSqlDBExec(SqlStr, QsqlIndex)
                If ret = QSQL_SUCCESS Then
                    If QSqlGetRow(record, QsqlIndex) = QSQL_SUCCESS Then
                        
                        QSqlGetField 2, record, SqlData()
                        NextLabNo = SqlData(1)
                        If SqlData(2) <> "" Then NxtRst = CSng(SqlData(2))
                    End If
                End If
                ret = QSqlSelectFree(QsqlIndex)
            
            End If
            
            If NextLabNo <> "" Then
                
                DeltaVal = NxtRst - CurRst
                Select Case DeltaGbn
                Case "2"
                    DeltaVal = (DeltaVal / CurRst) / 100
                Case "3"
                    If DeltaInterval <> 0 Then DeltaVal = DeltaVal / DeltaInterval
                Case "4"
                    If DeltaInterval <> 0 Then DeltaVal = ((DeltaVal / CurRst) / 100) / DeltaInterval
                End Select
                
                If DeltaVal > DeltaHi Or DeltaVal < DeltaLo Then
                    NextDelta = "D"
                End If
            End If
        End If
        
    End If
    
Delta_Check_Error:
    If Err <> 0 Then
        MsgBox Err & " ; " & Err.Description

        If Err = 6 Then         'Over Flow
            Resume Next
        End If
    End If
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * *
'*  결과값 판정 MODULE                  *
'*  참고치                                         *
'*  RetVal  : 결과값                           *
'*  OrdCd   : 검사코드                       *
'*  SubSqNo : SUB검사코드              *
'*  age     : 나이                                *
'*  Sex     : 성별                               *
'* * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Ref_Judg(RetVal As Variant, OrdCd As String, SubSqNo As String, age As Integer, sex As String) As String

    Dim SqlData() As String
    'S_REF = ""              '결과지 출력시 참고치 출력을 위해
    
    If RetVal = "" Then Exit Function
    
    If SubSqNo <> "" Then

        SqlStr = "SELECT  DISTINCT REFCHK, REFCHAR2, REFNUMHI, REFNUMLO" _
                 & " FROM LAB01_DB..SLA050M " _
                 & " WHERE ORDCD = '" & OrdCd & Chr$(39) _
                 & " AND SUBSQNO = '" & SubSqNo & Chr$(39)

    Else
        If age <= 16 Then           '소아
            SqlStr = "SELECT DISTINCT REFCHK, REFCHAR1, REFHIC, REFLOC " _
                    & " FROM LAB01_DB..SLA030M " _
                    & " WHERE ORDCD LIKE '" & OrdCd & Chr$(39)
                
        Else
            If sex = "여" Then
            
                SqlStr = "SELECT DISTINCT REFCHK, REFCHAR1, REFHIF, REFLOF " _
                        & "FROM LAB01_DB..SLA030M " _
                        & " WHERE ORDCD LIKE '" & OrdCd & Chr$(39)
            Else
                SqlStr = "SELECT DISTINCT REFCHK, REFCHAR1, REFHIM, REFLOM " _
                        & "FROM LAB01_DB..SLA030M " _
                        & " WHERE ORDCD LIKE '" & OrdCd & Chr$(39)
            End If
        End If
    End If

    ret = QSqlDBExec(SqlStr, QsqlCode)
    If ret = QSQL_SUCCESS Then
        If QSqlGetRow(record, QsqlCode) = QSQL_SUCCESS Then
    
            QSqlGetField 4, record, SqlData()
            
            If SqlData(1) = "1" Then               '숫자
                If Val(RetVal) > Val(SqlData(3)) Then
                    Ref_Judg = "H"
                ElseIf Val(RetVal) < Val(SqlData(4)) Then
                    Ref_Judg = "L"
                End If
    '            S_REF = SqlData(4) & " ~ " & SqlData(3)
            ElseIf SqlData(1) = "2" Then        '문자
                If SqlData(2) <> RetVal Then Ref_Judg = "*"
    '            S_REF = SqlData(2)
            End If
            
       End If
    End If
    ret = QSqlSelectFree(QsqlCode)
         
End Function
    
'* * * * * * * * * * * * * * * * * * * * * * * * *
'*  PANIC CHECK                             *
'*  RetVal  : 결과값                           *
'*  OrdCd   : 검사코드                       *
'*  Return 값  "P" Or  Space              *
'* * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Panic_Check(RetVal As Variant, OrdCd As String) As String

    Dim SqlData() As String
    
    If RetVal = "" Then Exit Function
    
    SqlStr = "SELECT DISTINCT PANICHI, PANICLO " _
            & " FROM LAB01_DB..SLA030M " _
            & " WHERE ORDCD LIKE '" & OrdCd & Chr$(39)

    ret = QSqlDBExec(SqlStr, QsqlCode)
    If ret = QSQL_SUCCESS Then
        If QSqlGetRow(record, QsqlCode) = QSQL_SUCCESS Then
    
            QSqlGetField 2, record, SqlData()
            
            If SqlData(1) <> "" And Val(SqlData(1)) <> 0 Then                  'Panic Check
                If Val(RetVal) > Val(SqlData(1)) Or Val(RetVal) < Val(SqlData(2)) Then
                    Panic_Check = "P"
                End If
            End If
            
       End If
    End If
    ret = QSqlSelectFree(QsqlCode)
         
End Function
    


