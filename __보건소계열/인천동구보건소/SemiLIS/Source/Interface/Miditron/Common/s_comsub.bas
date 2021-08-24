Attribute VB_Name = "S_COMSUB"
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  S_COMSUB = �����Լ� ���� Library                            *
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
'***  Report ���� ��������
'*******************************
Dim iVLs As Integer                 'Vertical Line�� ����
Dim iPageStartTop As Integer        'Page�� ������ġ�� Setting
'**********************************************************
'** �����๰��Ȳ ����.                                  ***
'** para    : ��������(1=�ܷ�, 2=����, 3=���޽�)        ***
'** par1    : ��������                                  ***
'** par2    : ��������                                  ***
'** par3    : �����ǹ�ȣ(�ܷ�)/�Կ���ȣ(����)           ***
'** frm     : Control name                              ***
'** ctr     : Control name                              ***
'**********************************************************
Sub S0SUB_SELECT_MEDICINES(frm As Form, ctr As Control, para As String, par1 As String, par2 As String, par3 As String)
  
    Dim SqlConn As Long
    Dim SqlDoc  As String: Dim sql_ret  As Integer
    Dim code()  As String
    
    ctr = ""

    If para = "1" Then   '�ܷ�
        If par1 = "3" Then
            '--- ������(�ܷ�)SERVER Open
            sql_ret = S0SUB_Open(S0COM_SERVER06, frm.hWnd, SqlConn)
            If sql_ret <> QSQL_SUCCESS Then
                Exit Sub
            End If
        Else
            '--- ����(�ܷ�)SERVER Open
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
        '�� ȭ�鿡�� ����� Index Open
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




'/* ȯ�ڸ� ���...
'/* para : ��������(1=�ܷ�, 2=����, 3=���޽�)
Sub S0SUB_SELECT_PATIENT(frm As Form, para As String)
  
    Dim SqlDoc  As String: Dim sql_ret  As Integer
    Dim patient()  As String
    
    S0COM_name = ""
    
    If para = "1" Then   '�ܷ�
        '�� ȭ�鿡�� ����� Index Open
        sql_ret = S0SUB_Open(S0COM_SERVER03, frm.hWnd, QsqlCod2)
        If sql_ret <> QSQL_SUCCESS Then
            Exit Sub
        End If
        SqlDoc = "SELECT PatNm FROM AC01B_DB..AC01B01M_TBL"
        SqlDoc = SqlDoc + " WHERE UnitNo = " & Chr(39) & S0COM_code & Chr(39)
    Else
        '�� ȭ�鿡�� ����� Index Open
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
'/* ������� ��ġ�� �� ���...
'/* S0COM_CODE : ������ڵ�+��ġ���ڵ�
Sub S0SUB_SELECT_AC01A10M(frm As Form)
  
    Dim SqlDoc  As String: Dim sql_ret  As Integer
    Dim dept()  As String
    
    S0COM_name = "": S0COM_name1 = ""
    
    sql_ret = S0SUB_Open(S0COM_SERVER04, frm.hWnd, QsqlCod2)
    If sql_ret <> QSQL_SUCCESS Then
        MsgBox "OCS���� ���� Error!!", 0
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

    sql_ret = QSqlSelectFree(QsqlCod2)                 ' �ڵ� ���� column�� set
        
    sql_ret = Qsqlclose(QsqlCod2, ONECLOSE)

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                               *
'*   ������Ϸ� ���̸� ���                      *
'*   passport_id   :  ������� ��ȯ��� data     *
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
'* String���� currency mode�� ��ȯ               *
'*    wf_src1   :  ��ȯ ��� data                *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CCur(wf_src1 As String) As Currency
    On Error GoTo S0SUB_CCUR_ERROR

    If Len(Trim(wf_src1)) = 0 Or IsNull(wf_src1) Then                 ' NULL ?
        wf_src1 = "0"                              ' zero set
    End If
    S0SUB_CCur = CCur(wf_src1)                     ' currency(money) mode ��ȯ
    Exit Function

S0SUB_CCUR_ERROR:
    
    S0SUB_CCur = 0@                                ' error�� ��� zero set
    Exit Function

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String���� double mode�� ��ȯ                 *
'*    wf_src1   :  ��ȯ ��� data                *
'*    slen      :  �Ҽ��� ���� �ڸ���            *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CDBL(wf_src1 As String, slen As Integer) As Double
    On Error GoTo S0SUB_CDBL_ERROR

    If Len(Trim(wf_src1)) = 0 Then               ' NULL ?
       If slen = 0 Then                          ' �Ҽ��� ���� ���� ?
          wf_src1 = "0"
       ElseIf slen = 1 Then                      ' �Ҽ��� ���� 1 �ڸ� ?
          wf_src1 = "0.0"
       ElseIf slen = 2 Then                      ' �Ҽ��� ���� 2 �ڸ� ?
          wf_src1 = "0.00"
       ElseIf slen = 3 Then                      ' �Ҽ��� ���� 3 �ڸ� ?
          wf_src1 = "0.000"
       ElseIf slen = 4 Then                      ' �Ҽ��� ���� 4 �ڸ� ?
          wf_src1 = "0.0000"
       ElseIf slen = 5 Then                      ' �Ҽ��� ���� 5 �ڸ� ?
          wf_src1 = "0.00000"
       End If
    End If
    S0SUB_CDBL = CDbl(wf_src1)                   ' double mode ��ȯ
    Exit Function

S0SUB_CDBL_ERROR:
    
    S0SUB_CDBL = 0#                              ' error �ǰ�� zero set
    Exit Function

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* Table �κ��� �ڵ��Ī�� SELECT �Ͽ�           *
'*                       ��Ī Control �� ǥ��    *
'*    wf_src1   :  ��ȯ ��� data                *
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
    S0COM_ret = ret                                 ' QSql�� Return�� Setting
    ret% = QSqlSelectFree(QsqlCode)                 ' �ڵ� ���� column�� set

    Exit Sub

QsqlFail:
    S0COM_name = ""                                 ' Error�� Null Return
    S0COM_ret = ret                                 ' QSql�� Return�� Setting
    ret% = QSqlSelectFree(QsqlCode)                 ' �ڵ� ���� column �� set

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String���� integer mode�� ��ȯ                *
'*    wf_src1   :  ��ȯ ��� data                *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CINT(wf_src1 As String) As Integer
    On Error GoTo S0SUB_CINT_ERROR

    If Len(Trim(wf_src1)) = 0 Then                 ' NULL ?
        wf_src1 = "0"                              ' zero set
    End If
    S0SUB_CINT = CInt(wf_src1)                     ' integer mode ��ȯ
    Exit Function

S0SUB_CINT_ERROR:
    
    S0SUB_CINT = 0                                 ' error �ǰ�� zero set
    Exit Function

End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'* String���� long integer mode�� ��ȯ           *
'*    wf_src1   :  ��ȯ ��� data                *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CLng(wf_src1 As String) As Long
    On Error GoTo S0SUB_CLNG_ERROR

    If Len(Trim(wf_src1)) = 0 Then                 ' NULL ?
        wf_src1 = "0"                              ' zero set
    End If
    S0SUB_CLng = CLng(wf_src1)                     ' long integer mode ��ȯ
    Exit Function

S0SUB_CLNG_ERROR:
    
    S0SUB_CLng = 0&                                ' error �ǰ�� zero set
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
    
'�˻��� ȯ�� ���̺�
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

'Hitachi �˻��׸� ����
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
    
'KODAK �˻��׸� ����
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
    
'STRATUS 1 �˻��׸� ����
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
    
'STRATUS 2 �˻��׸� ����
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
    
'Coulter STKS �˻��׸� ����
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
    
'Coulter T-540 �˻��׸� ����
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
'*  Hitachi Interface ����ޱ� ���̺� ����          *
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
    
     '�˻籸�� : 1=���� ���հǰ����ܼ��� 2=����ڼ���, �����ڼ���
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
'*  Kodak ��� ���̺� ����                              *
'*  para      : ���ϸ�                                  *
'*  ExamCount : �˻��׸� ����                           *
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
    
     '�˻籸�� : 1=���� ���հǰ����ܼ��� 2=����ڼ���, �����ڼ���
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
'*  Stratus 1,2 Interface ����ޱ� ���̺� ����      *
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
    
     '�˻籸�� : 1=���� ���հǰ����ܼ��� 2=����ڼ���, �����ڼ���
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
    
     '�˻籸�� : 1=���� ���հǰ����ܼ��� 2=����ڼ���, �����ڼ���
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
    
     '�˻籸�� : 1=���� ���հǰ����ܼ��� 2=����ڼ���, �����ڼ���
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
'* String���� single mode�� ��ȯ                 *
'*    wf_src1   :  ��ȯ ��� data                *
'*    slen      :  �Ҽ��� ���� �ڸ���            *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_CSNG(wf_src1 As String, slen As Integer) As Single
    On Error GoTo S0SUB_CSNG_ERROR

    If Len(Trim(wf_src1)) = 0 Then               ' NULL ?
       If slen = 0 Then                          ' �Ҽ��� ���� ���� ?
          wf_src1 = "0"
       ElseIf slen = 1 Then                      ' �Ҽ��� ���� 1 �ڸ� ?
          wf_src1 = "0.0"
       ElseIf slen = 2 Then                      ' �Ҽ��� ���� 2 �ڸ� ?
          wf_src1 = "0.00"
       ElseIf slen = 3 Then                      ' �Ҽ��� ���� 3 �ڸ� ?
          wf_src1 = "0.000"
       ElseIf slen = 4 Then                      ' �Ҽ��� ���� 4 �ڸ� ?
          wf_src1 = "0.0000"
       End If
    End If
    S0SUB_CSNG = CDbl(wf_src1)                   ' single mode ��ȯ
    Exit Function

S0SUB_CSNG_ERROR:
    
    S0SUB_CSNG = 0!                              ' error �ǰ�� zero set
    Exit Function

End Function

'*-------------------------------------------------*
'* req      : "1" ==> yyyymmdd return              *
'*          : "2" ==> yyyy-mm-dd return            *
'* para     : value                    (i)         *
'*-------------------------------------------------*
Function S0SUB_DATE_6TO8(ByVal para As String, ByVal req As String) As String

    Dim temp As String

    temp = Left(para, 2)                    '�⵵ ����
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
'* ���� check �� format ó��                       *
'* req      : "1" ==> yyyy check                   *
'*          : "2" ==> yyyymm check                 *
'*          : "3" ==> yyyymmdd check               *
'* delimeter: "/" or "-" or "."        (i)         *
'* date1    : yyyymmdd                 (i/o)       *
'* date2    : yyyy-mm-dd               (i/o)       *
'* iret     : 1 ==> succeed            (o)         *
'*          : -1 ==> parameter error   (o)         *
'*          : -2 ==> ���� error        (o)         *
'*-------------------------------------------------*
Sub S0SUB_DATE_FORMAT(req As String, delimeter As String, Date1 As String, date2 As String, iRet As Integer)

    Dim slen%

    iRet = 1                                     ' ���� ó�� set
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

    If IsDate(date2) = False Then                ' ���� check error ?
        iRet = -2                                ' ���� error
        Exit Sub
    End If

End Sub

'********************************************************
'*                                                      *
'*  ����Ⱓ�� ���� interface file�� �����.            *
'*  DirName  : ������� �ϴ� Directory Name             *
'*  SaveDate : ����Ⱓ                                 *
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

    ret% = QSqlSelectFree(QsqlCode)                 ' �ڵ� ���� column�� set
    
    'ctr.ListIndex = 0

End Sub

'*  *   *   *   *   *   *   *   *   *   *   *
'*                                          *
'*  ������ ���翩���� �ľ��Ѵ�.             *
'*  para : ���ϸ�(��θ� ����)              *
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
'*  Record�� �����ϸ� True, �ƴϸ� False                *
'*  para : SQL ��                                       *
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
'*  ��ġ ������ ���� control ���� 1 col.���� ǥ��*
'*    wf_src1   :  ��� ����                     *
'*    slen      :  �Ҽ������� �ڸ��� (double���)*
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
'*  ��ġ ������ ���� control ���� 1 col.���� ǥ��*
'*    wf_src1   :  ��� ����                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_FCDBL2(ByVal wf_src1 As String) As String

    Dim pos As Integer
    Dim dbl_para As String

    If Trim(wf_src1) = "" Then
        S0SUB_FCDBL2 = ""
        Exit Function
    End If

    pos = InStr(Trim(wf_src1), ".")
    If pos = 0 Then                                 '�����ΰ��
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
'*  �ڸ����� ����ġ �ڸ��� ��ŭ ��ȯ                *
'*  para : ��ȯ�ϰ��� �ϴ� ��                       *
'*  defValue : ����ġ ��                            *
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
'*  �����ڵ�/�������� ã�´�.                       *
'*  para    : ã�����ϴ� ��                         *
'*  chk     : 1 = �����ڵ�, 2 = ������              *
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
'*  grid���� Ư������ ����Ÿ�� ã�� ��ġ�� ������..             *
'*                                                              *
'*  grd          : grid Name                                    *
'*  Col          : Column                                       *
'*  para         : Data                                         *
'*  Return Value : ����Ÿ�� ��ġ ��                             *
'*                 ã���� �ϴ� ����Ÿ�� ���� ��� -1�� SETTING  *
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
            
            grd.HighLight = True                                            '�Է¶��� ����
            grd.SelStartRow = Row: grd.SelEndRow = Row
            grd.SelStartCol = grd.FixedCols: grd.SelEndCol = grd.Cols - 1
            
            Exit Function
        End If
    Next

    S0SUB_GRIDGETROW = -1

End Function

'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                      *
'*  grid���� Ư���ฦ ȭ�鿡 �����ش�..                 *
'*  grd : grid Name                                     *
'*  Row : Row                                           *
'*  disRow : ȭ�鿡 ������ �� �ִ� �ִ� ���� ��         *
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
'*  SYSTEM LOGON ID�� SETTING                          *
'*  (USER-ID,TEMINAL-ID,CURRENCY-DATE,CURRENCY-TIME)   *
'*                                                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Sub S0SUB_LOGONID()

    Dim sql_ret As Integer
    Dim record  As String, SqlData()    As String
    
    '������ SYSTEM���ڿ� �ð��� �о�´�.
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
'* String���� Length���� ª���� Left Space ä�� *
'*    ��� �ڸ�(�ѱ� ������ ���� ó��)           *
'*    w_text    :  ǥ�� ��� data                *
'*    w_len     :  ǥ�� ����                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_LSPACE(w_text As String, w_len As Integer) As String
    
    Dim s_len As Integer
    Dim ch As Integer, i As Integer, st As Integer

    s_len = Len(w_text)                              ' length ���
    
    If s_len <= w_len Then                           ' Left SPACE ä��
        S0SUB_LSPACE = Space$(w_len - s_len) + w_text
        Exit Function
    Else
        st = 0                  ' �ѱ� ©������ reset
        For i = 1 To w_len
            ch = Asc(Mid$(w_text, i, 1))
            If ch < &H80 Then                        ' �ѱ�/���� check ?
                st = 0                               ' �ѱ� ©������ reset
            Else
                st = (st + 1) Mod 2                  ' �ѱ� ©������ set
            End If
        Next i

        If st = 0 Then                               ' �ѱ� ©������ check ?
            S0SUB_LSPACE = Left$(w_text, w_len)      ' ������ �ѱ� ���� set
        Else                                         ' ������ �ѱ� �ڸ�
            S0SUB_LSPACE = " " + Left$(w_text, w_len - 1)
        End If
    End If
            
End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*  masked edit�� data type display              *
'*    dest_ctl  :  ǥ�ô�� control              *
'*    w_data    :  ǥ�� data                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Sub S0SUB_mask_disp(Dest_Ctl As Control, w_data As String)
    
    Dim w_chk1  As Integer                       ' change flag

    w_chk1 = 0                                   ' change flag reset

    If Dest_Ctl.PromptInclude = True Then        ' control �Ӽ� = true ?
        Dest_Ctl.PromptInclude = False           ' control �Ӽ� false set
        w_chk1 = 1                               ' change flag set
    End If

    If Len(Trim(w_data)) = 0 Then                ' ǥ�� data NULL ?
        Dest_Ctl.Text = ""                       ' NULL set
    Else
        Dest_Ctl.Text = w_data                   ' ǥ�� data set
    End If
    
    If w_chk1 = 1 Then                           ' change flag set ?
        Dest_Ctl.PromptInclude = True            ' control �Ӽ� true reset
    End If

End Sub

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                               *
'*    �ش�޼�����  �����Ͽ� ȭ�鿡 ǥ���ϱ�     *
'*    para  :  �ش�޼��� ��ȣ                   *
'*                                               *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_MESSAGE(ByVal para As Integer) As String

    Select Case para
        Case 1: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 2: S0SUB_MESSAGE = "���������� �����Ǿ����ϴ�."
        Case 3: S0SUB_MESSAGE = "���������� �����Ǿ����ϴ�."
        Case 4: S0SUB_MESSAGE = "���������� ��ȸ�Ǿ����ϴ�."
        Case 5: S0SUB_MESSAGE = "���������� �μ�Ǿ����ϴ�."
        Case 6: S0SUB_MESSAGE = "�ش� �ڷᰡ �������� �ʽ��ϴ�."
        Case 7: S0SUB_MESSAGE = "Ű���� ����Ǿ����ϴ�! Ȯ�ιٶ��ϴ�."
        Case 8: S0SUB_MESSAGE = "������ ��ҵǾ����ϴ�."
        Case 9: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 10: S0SUB_MESSAGE = """ ' "" �� �Է��� �� ���� �����Դϴ�."
        Case 11: S0SUB_MESSAGE = "�׸��� �����Ͽ� �ֽʽÿ�."
        Case 12: S0SUB_MESSAGE = "���������� ó���Ǿ����ϴ�.."
        Case 13: S0SUB_MESSAGE = "�Էµ� USER-ID�� ������ �����Ƿ� ����� �� �����ϴ�.."
        Case 14: S0SUB_MESSAGE = "�Էµ� USER-ID�� ������ �����Ƿ� ������ �� �����ϴ�.."


        Case 101: S0SUB_MESSAGE = "�����ȣ ���� Ŭ�� �����ϴ�."
        Case 102: S0SUB_MESSAGE = "��¥�Է��� Ʋ���ϴ�.  Ȯ���ϼ���."

        Case 103: S0SUB_MESSAGE = "�ŷ����ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 104: S0SUB_MESSAGE = "�ŷ����ڵ带 ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 105: S0SUB_MESSAGE = "��� ����ڸ� �Է��Ͽ� �ֽʽÿ�."
        Case 106: S0SUB_MESSAGE = "��� ����ڸ� ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 107: S0SUB_MESSAGE = "�ŷ�ó ����ڸ� �Է��Ͽ� �ֽʽÿ�."
        Case 108: S0SUB_MESSAGE = "�����ð��� �Է��Ͽ� �ֽʽÿ�."
        Case 109: S0SUB_MESSAGE = "�����ð��� ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 110: S0SUB_MESSAGE = "��������ġ�� �Է��Ͽ� �ֽʽÿ�."
        Case 111: S0SUB_MESSAGE = "��������ġ�� ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 112: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 113: S0SUB_MESSAGE = "�׸��� �����Ͽ� �ֽʽÿ�."
        Case 114: S0SUB_MESSAGE = "����ڹ�ȣ�� ����Ǿ����ϴ�. Ȯ���Ͻʽÿ�."
        Case 115: S0SUB_MESSAGE = "����, �������� ��������ġ�� ���� �Ͻʽÿ�."
        Case 116: S0SUB_MESSAGE = "����, �������� SUB�׸��� ���� �Ͻʽÿ�."
        Case 117: S0SUB_MESSAGE = " ""000""�� �Է��� �� �����ϴ�."
        Case 118: S0SUB_MESSAGE = "�к��ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 119: S0SUB_MESSAGE = "�к��ڵ带 ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 120: S0SUB_MESSAGE = "�����ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 121: S0SUB_MESSAGE = "�����ڵ带 ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 122: S0SUB_MESSAGE = "�˻��׸��ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 123: S0SUB_MESSAGE = "�˻��׸��ڵ带 ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 124: S0SUB_MESSAGE = "�˻������ �Է��Ͽ� �ֽʽÿ�."
        Case 125: S0SUB_MESSAGE = "����ڹ�ȣ�� �Է��Ͽ� �ֽʽÿ�."
        Case 126: S0SUB_MESSAGE = "����ڸ��� �Է��Ͽ� �ֽʽÿ�."
        Case 127: S0SUB_MESSAGE = "��й�ȣ�� �Է��Ͽ� �ֽʽÿ�."
        Case 128: S0SUB_MESSAGE = "��й�ȣ�� 6�ڸ� �̸��� �ɼ� �����ϴ�."
        Case 129: S0SUB_MESSAGE = "�м�/���� �ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 130: S0SUB_MESSAGE = "�м�/���� ��Ī�� �Է��Ͽ� �ֽʽÿ�."
        Case 131: S0SUB_MESSAGE = "�������ڸ� �Է��Ͽ� �ֽʽÿ�."
        Case 132: S0SUB_MESSAGE = "�������ڸ� ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 133: S0SUB_MESSAGE = "�мҴ���ڸ� �Է��Ͽ� �ֽʽÿ�."
        Case 134: S0SUB_MESSAGE = "���������ð��� �Է��Ͽ� �ֽʽÿ�."
        Case 135: S0SUB_MESSAGE = "���ڸ� �Է��Ͽ� �ֽʽÿ�."
        Case 136: S0SUB_MESSAGE = "�ŷ������� �Է��Ͽ� �ֽʽÿ�."
        Case 137: S0SUB_MESSAGE = "�����ڸ� �Է��Ͽ� �ֽʽÿ�."
        Case 138: S0SUB_MESSAGE = "����ġ�� �Է��Ͽ� �ֽʽÿ�."
        Case 139: S0SUB_MESSAGE = "���� ������ ������ �����ϴ�."
        Case 140: S0SUB_MESSAGE = "ȭ����� �ڷḦ �Է¶Ǵ� ������ ���� �۾��� �Ͻʽÿ�."


        Case 201: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 202: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 203: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 204: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 205: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 206: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 207: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 208: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 209: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 210: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."


        Case 301: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 302: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 303: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 304: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 305: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 306: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 307: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 308: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 309: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."
        Case 310: S0SUB_MESSAGE = "���������� �ԷµǾ����ϴ�."


        Case 401: S0SUB_MESSAGE = "�ұ׷��ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 402: S0SUB_MESSAGE = "�ұ׷��ڵ带 ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 403: S0SUB_MESSAGE = "�ִ��׸��� �ʰ��Ͽ� �Է��� �� �����ϴ�."
        Case 404: S0SUB_MESSAGE = "�˻���� �Է��Ͽ� �ֽʽÿ�."
        Case 405: S0SUB_MESSAGE = "����ġ�� �Է��Ͽ� �ֽʽÿ�."
        Case 406: S0SUB_MESSAGE = "���� ������ ������ �����ϴ�."
        Case 407: S0SUB_MESSAGE = "�����ڸ� �Է��Ͽ� �ֽʽÿ�."
        Case 408: S0SUB_MESSAGE = "�˻�з��� �Է��Ͽ� �ֽʽÿ�."
        Case 409: S0SUB_MESSAGE = "�˻�з��� ""1"" ~ ""3"" �� �����Ͽ� �Է��Ͻʽÿ�."
        Case 410: S0SUB_MESSAGE = "���ڰ��� �Է��Ͽ� �ֽʽÿ�."
        Case 411: S0SUB_MESSAGE = "��׷��ڵ�� ""700""������ �Է��� �����մϴ�."
        Case 412: S0SUB_MESSAGE = "��ɱ����ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 413: S0SUB_MESSAGE = "�������ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 414: S0SUB_MESSAGE = "�������ڵ带 ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 415: S0SUB_MESSAGE = "������ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 416: S0SUB_MESSAGE = "CHART �Ϸù�ȣ�� �Է��Ͽ� �ֽʽÿ�."
        Case 417: S0SUB_MESSAGE = "ȯ�ڸ��� �Է��Ͽ� �ֽʽÿ�."
        Case 418: S0SUB_MESSAGE = "���ڸ� �Է��Ͽ� �ֽʽÿ�."
        Case 419: S0SUB_MESSAGE = "���ڸ� ��Ȯ�ϰ� �Է��Ͽ� �ֽʽÿ�."
        Case 420: S0SUB_MESSAGE = " ""1"" ~ ""2""�� �Է��� �����մϴ�."
        Case 421: S0SUB_MESSAGE = "�Ϸù�ȣ�� ��Ȯ�� �Է��Ͽ� �ֽʽÿ�."
        Case 422: S0SUB_MESSAGE = "CHART ��ȣ�� �Է��Ͽ� �ֽʽÿ�."
        Case 423: S0SUB_MESSAGE = "CHART ��ȣ�� ��Ȯ�� �Է��Ͽ� �ֽʽÿ�."
        Case 424: S0SUB_MESSAGE = "��׷��ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 425: S0SUB_MESSAGE = "��׷��ڵ带 ��Ȯ�� �Է��Ͽ� �ֽʽÿ�."
        Case 426: S0SUB_MESSAGE = "�ݾ��� �Է��Ͽ� �ֽʽÿ�."
        Case 427: S0SUB_MESSAGE = "���������� �Է��Ͽ� �ֽʽÿ�."
        Case 428: S0SUB_MESSAGE = "������ȣ�� ��Ȯ�� �Է��Ͽ� �ֽʽÿ�."
        Case 429: S0SUB_MESSAGE = "�Է��� �κа˻縦 �Է��Ͽ� �ֽʽÿ�."
        Case 430: S0SUB_MESSAGE = "�˻��׸� �Է��ڷᰡ �����ϴ�."
        Case 431: S0SUB_MESSAGE = "����ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 432: S0SUB_MESSAGE = "����ڵ带 ��Ȯ�� �Է��Ͽ� �ֽʽÿ�."
        Case 433: S0SUB_MESSAGE = "�Ұ��ڵ带 �Է��Ͽ� �ֽʽÿ�."
        Case 434: S0SUB_MESSAGE = "�Ұ��ڵ带 ��Ȯ�� �Է��Ͽ� �ֽʽÿ�."
        Case 435: S0SUB_MESSAGE = "ȯ�������� �����ϴ�. Ȯ�ιٶ��ϴ�!"
        Case 436: S0SUB_MESSAGE = "�ڷḦ �Է��Ͽ� �ֽʽÿ�."
        Case 437: S0SUB_MESSAGE = "�ڷḦ �Է��Ͽ� �ֽʽÿ�."
        Case 438: S0SUB_MESSAGE = "��±����� ""1"", ""2"", ""3"", ""4"", ""9""�� �Է��Ͽ� �ֽʽÿ�."
        Case 439: S0SUB_MESSAGE = "�Ұ��ڵ� �ߺ��Դϴ�. Ȯ�ιٶ��ϴ�!"
        Case 440: S0SUB_MESSAGE = "���ռҰ��ڵ� �ߺ��Դϴ�. Ȯ�ιٶ��ϴ�!"
        Case 441: S0SUB_MESSAGE = "�к��ڵ带 �����Ͽ� �ֽʽÿ�."
        Case 442: S0SUB_MESSAGE = "�˻��׸��� �����Ͽ� �ֽʽÿ�."
        Case 443: S0SUB_MESSAGE = "������� �Է��Ͽ� �ֽʽÿ�."
        Case 444: S0SUB_MESSAGE = "������ �Է��Ͽ� �ֽʽÿ�."
        Case 445: S0SUB_MESSAGE = "�ֹε�Ϲ�ȣ�� ��Ȯ�� �Է��Ͽ� �ֽʽÿ�."
        Case 446: S0SUB_MESSAGE = "�ߺ��� �ڷ��Դϴ�. Ȯ�ιٶ��ϴ�!"
        Case 447: S0SUB_MESSAGE = "�˻��׸��� 1���̻� �Է��Ͽ� �ֽʽÿ�."
        Case 448: S0SUB_MESSAGE = "���ڰ��� �Է��Ͽ� �ֽʽÿ�."

        Case 501: S0SUB_MESSAGE = "�м��ڵ� �Ǵ� �ŷ����ڵ尡 �����ϴ�. "
        Case 502: S0SUB_MESSAGE = "��¥�� Ȯ���ϼ���."
        Case 503: S0SUB_MESSAGE = "���� ����!!  ���������� ���� ������Դϴ�."
        Case 504: S0SUB_MESSAGE = "���� ����!!  ���������� ���� ������Դϴ�."
        Case 505: S0SUB_MESSAGE = "���� ��¥���� �ռ� �� �����ϴ�."
        Case 506: S0SUB_MESSAGE = "���� �ڵ庸�� �ռ� �� �����ϴ�."
        Case 507: S0SUB_MESSAGE = "�ڵ��� ��ȿ������ �ƴմϴ�."
        Case 508: S0SUB_MESSAGE = "�ش� �м��ڵ尡 �������� �ʽ��ϴ�.  Ȯ���Ͽ� �ּ���."
        Case 509: S0SUB_MESSAGE = "�ش� �ŷ�ó�ڵ尡 �������� �ʽ��ϴ�.  Ȯ���Ͽ� �ּ���."
        Case 510: S0SUB_MESSAGE = "���ݾ��� ���̼��ݾ׺��� Ŭ�� �����ϴ�.  Ȯ���Ͽ� �ּ���."
        
        Case 536: S0SUB_MESSAGE = "DB OPEN�� �Ǿ����� �ʽ��ϴ�."

    End Select

    
End Function

'****************************************************
'*                                                  *
'*  ������ ���翩�θ� �ľ��Ѵ�.                     *
'*  para : ���ϸ�                                   *
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
'*  ��ȯ ��� data�� format ���·� ��ȯ          *
'*    src       :  ��ȯ ��� data                *
'*    fmat      :  ��ȯ ����                     *
'*    gbcd      :  ��ȯ ����("T":����,"R":�ݿø� *
'*                           "X":�ø�)           *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_num_format(src As String, fmat As String, gbcd As String) As String
    Dim i       As Integer                       ' for loop index
    Dim j       As Integer                       ' fmat �� �Ҽ��� ���� �ڸ���
    Dim II      As Integer                       ' src  �� �Ҽ��� ��ġ
    Dim JJ      As Integer
    Dim xx      As Integer                       ' src �� �����ڸ���
    Dim yy      As Integer                       ' fmat �� �����ڸ���

    Dim Text    As String                        ' "," ���� src data
    Dim maxlen  As Integer                       ' fmat ����
    Dim maxlen1 As Integer                       ' src  ����
    Dim slen    As Integer                       ' �Ҽ��� ���� �ڸ���
    Dim mid1    As String                        ' return value
    Dim mid2    As Double                        ' src value
    Dim mid3    As Double                        ' src ���밪
    Dim src1    As String                        ' src �� �Ҽ��� ���� data
    Dim chk     As String                        ' ���� flag
    Dim chk1    As String                        ' �ε��Ҽ��� flag
    On Error GoTo FORMAT_ERROR

    src = Trim(src)                              ' NULL cut
    If Len(src) = 0 Then                         ' NULL ?
        src = "0"                                ' zero set
    End If
    maxlen = Len(fmat)                           ' fmat ���� set
    maxlen1 = Len(src)                           ' src  ���� set
    Text = ""                                    ' return data �ʱ�ó��
    For i = 1 To maxlen1                         ' src ���� "," �� ����
        If Mid(src, i, 1) = "," Then
        Else
           Text = Text & Mid(src, i, 1)
        End If
    Next

    Text = Trim(Text)

    yy = 0
    For i = 1 To maxlen                          ' fmat���� "," �� �����ϰ�
        If Mid(fmat, i, 1) = "." Then            '   �����κ��ڸ���(yy)�� ���
           Exit For
        End If
        If Mid(fmat, i, 1) = "," Then
        Else
           yy = yy + 1
        End If
    Next
    xx = 0
    For i = 1 To maxlen1                         ' src���� ","�� �����ϰ�
        If Mid(src, i, 1) = "." Then             '   �����κ� �ڸ���(xx)�� ���
           Exit For
        End If
        If Mid(src, i, 1) = "," Then
        Else
           xx = xx + 1
        End If
    Next

    If xx > yy Then                              ' src ������ > fmat ������ ?
    Else
       mid1 = Format(S0SUB_CDBL(src, 0), fmat)   '   double ��ȯ �� fmat ��ȯ
    End If

    chk1 = "N"                                   ' �ε��Ҽ��� flag reset
    j = 0                                        ' fmat�� �Ҽ��� ���� �ڸ��� reset
    For i = 1 To maxlen
        If chk1 = "Y" Then
           j = j + 1                             ' fmat�� �Ҽ��� ���� �ڸ��� ���
        End If
        If Mid(fmat, i, 1) = "." Then            ' �ε��Ҽ��� data ?
           chk1 = "Y"                            ' �ε��Ҽ��� flag set
        End If
    Next
                                                 ' src��fmat���� �������̰��涧 fmat�� �ڸ�����������
    If xx > yy Then                              '   src���ڷ�(Text)�� �տ������� �ڸ���.
       src = Mid(Text, xx - yy + 1, yy) & "." & Mid(Text, xx + 2, j)
       mid1 = Format(S0SUB_CDBL(src, 0), fmat)
    End If


    slen = j                                     ' �Ҽ��� ���� �ڸ���

    mid2 = S0SUB_CDBL(src, j)                    ' src double ��ȯ
    If mid2 < 0 Then                             ' ���� ?
       chk = "Y"                                 '   ���� flag set
    End If
    mid3 = Abs(mid2)                             ' ���밪 ���


    If gbcd = "T" Then                           ' ������ ���
        For II = 1 To maxlen1                    '   src�� �Ҽ��� ��ġ ���
            If Mid(src, II, 1) = "." Then
                Exit For
            End If
        Next

        If II > maxlen1 Then GoTo format_m       '   �����κи� ������
        src1 = Mid(src, II + 1, maxlen1 - II)    '   �Ҽ������� data
        If Val(src1) = 0 Then GoTo format_m      '   �Ҽ������ϰ� zero �϶�
    End If

    If gbcd = "T" And mid3 <> 0 Then             ' ������ ���
       If slen = 0 Then                          '   �Ҽ������� �ڸ��� ó��
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
    ElseIf gbcd = "X" Then                       ' �ø��� ���
       If slen = 0 Then                          '   �Ҽ������� �ڸ��� ó��
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
    If chk = "Y" Then                            ' ������ ���
       mid3 = mid3 * -1                          '   ���� ó��
    End If

    mid1 = Format(mid3, fmat)
    mid1 = Space(maxlen - Len(mid1)) & mid1      ' leading space set
FORMAT_END:
    S0SUB_num_format = mid1                      ' return
    Exit Function
FORMAT_ERROR:
    mid1 = "0"                                   ' error �� ��� zero set
    Beep
    Resume FORMAT_END
End Function

'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  S0SUB_Open = DB Open ó��                                   *
'*                                                              *
'*    Ư�� Index�� �����Ͽ� open�� ��� ���                    *
'*                                                              *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_Open(Server As String, ByVal frmhWnd As Integer, Index As Long)
 
    Dim ret%                    'Qsqlclose RETURN CD

    Index = -1

   'SERVER �������Ȯ��
    ret% = QSqlOpen(Server, frmhWnd, Index)

    If ret% <> QSQL_SUCCESS Then
        Beep
        MsgBox QSqlError(ret%)
    End If

    S0SUB_Open = ret%

End Function

' ��������
'
Public Function S0SUB_SET_RECCHK(para1 As Integer, para2 As String) As String

    If para1 = 1 Then
    
        If para2 = "1" Then
            S0SUB_SET_RECCHK = "�ܷ�"
        ElseIf para2 = "2" Then
            S0SUB_SET_RECCHK = "�Կ�"
        ElseIf para2 = "3" Then
            S0SUB_SET_RECCHK = "���޽�"
        Else
            S0SUB_SET_RECCHK = "��Ź"
        End If
        
    Else
        If para2 = "�ܷ�" Then
            S0SUB_SET_RECCHK = "1"
        ElseIf para2 = "�Կ�" Then
            S0SUB_SET_RECCHK = "2"
        ElseIf para2 = "���޽�" Then
            S0SUB_SET_RECCHK = "3"
        End If
        
    End If
         
End Function

' ��������
'
'
Public Function S0SUB_SET_HOSCHK(para1 As Integer, para2 As String) As String
    
    If para1 = 1 Then
        
        If para2 = "1" Then
            S0SUB_SET_HOSCHK = "����"
    
        ElseIf para2 = "2" Then
            S0SUB_SET_HOSCHK = "����"
        
        ElseIf para2 = "3" Then
            S0SUB_SET_HOSCHK = "����������"
        
        ElseIf para2 = "4" Then
            S0SUB_SET_HOSCHK = "��Ȱ��"
        
        ElseIf para2 = "5" Then
            S0SUB_SET_HOSCHK = "�ϼ���"
        
        ElseIf para2 = "F" Then
            S0SUB_SET_HOSCHK = "���̺���"
        
        ElseIf para2 = "E" Then
            S0SUB_SET_HOSCHK = "���޽�"
        
        End If
        
    Else
        
        If para2 = "����" Then
            S0SUB_SET_HOSCHK = "1"
        
        ElseIf para2 = "����" Then
            S0SUB_SET_HOSCHK = "2"
        
        ElseIf para2 = "����������" Then
            S0SUB_SET_HOSCHK = "3"
        
        ElseIf para2 = "��Ȱ��" Then
            S0SUB_SET_HOSCHK = "4"
        
        ElseIf para2 = "�ϼ���" Then
            S0SUB_SET_HOSCHK = "5"
        
        ElseIf para2 = "���̺���" Then
            S0SUB_SET_HOSCHK = "F"
        
        ElseIf para2 = "���޽�" Then
            S0SUB_SET_HOSCHK = "E"
        
        End If
            
    End If
    
End Function


'*********************************************************
'** �ڵ�/��Ī help ȭ�� ǥ����ġ ��� �� �̵�          ***
'** xpos      : help field �� left position            ***
'** ypos      : help field �� top position             ***
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
'* String���� Length���� ª���� right space ä�� *
'*    ��� �ڸ�(�ѱ� ������ ���� ó��)           *
'*    w_text    :  ǥ�� ��� data                *
'*    w_len     :  ǥ�� ����                     *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Function S0SUB_RSPACE(w_text As String, w_len As Integer) As String
    Dim s_len As Integer
    Dim ch As Integer, i As Integer, st As Integer

    s_len = Len(w_text)                              ' length ���
    
    If s_len <= w_len Then                           ' right SPACE ä��
        S0SUB_RSPACE = w_text + Space$(w_len - s_len)
        Exit Function
    Else
        st = 0                                       ' �ѱ� ©������ reset
        For i = 1 To w_len
            ch = Asc(Mid$(w_text, i, 1))
            If ch < &H80 Then                        ' �ѱ�/���� check ?
                st = 0                               ' �ѱ� ©������ reset
            Else
                st = (st + 1) Mod 2                  ' �ѱ� ©������ set
            End If
        Next i

        If st = 0 Then                               ' �ѱ� ©������ check ?
            S0SUB_RSPACE = Left$(w_text, w_len)      ' ������ �ѱ� ���� set
        Else                                         ' ������ �ѱ� �ڸ�
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
    's_len = Len(w_text)                              ' length ���
    
    'If s_len <= w_len Then                           ' right SPACE ä��
    '    S0SUB_RSPACE = w_text + Space$(w_len - s_len)
    '    Exit Function
    'Else
     '   st = 0                                       ' �ѱ� ©������ reset
    '    For i = 1 To w_len
    '        ch = Asc(Mid$(w_text, i, 1))
    '        If ch < &H80 Then                        ' �ѱ�/���� check ?
    '            st = 0                               ' �ѱ� ©������ reset
    '        Else
    '            st = (st + 1) Mod 2                  ' �ѱ� ©������ set
    '        End If
    '    Next i

    '    If st = 0 Then                               ' �ѱ� ©������ check ?
    '        S0SUB_RSPACE = Left$(w_text, w_len)      ' ������ �ѱ� ���� set
    '    Else                                         ' ������ �ѱ� �ڸ�
    '        S0SUB_RSPACE = Left$(w_text, w_len - 1) + " "
    '    End If
    'End If

End Function

'****************************************************
'*                                                  *
'*  �����ڵ�/�������� �迭�� �ֱ�                   *
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
'*  vaSpread�� ����Ÿ�� Clear�Ѵ�.                       *
'*  spd : vaSpread��,  DispRow : vaSpread�� ���μ�    *
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
'*  vaSpread���� Ư������ ����Ÿ�� ã�� ��ġ�� ������..      *
'*                                                              *
'*  spd          : vaSpread Name                             *
'*  row          : Row                                          *
'*  para         : Data                                         *
'*  Return Value : ����Ÿ�� ��ġ ��                             *
'*                 ã���� �ϴ� ����Ÿ�� ���� ��� -1�� SETTING  *
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
'*  vaSpread���� Ư������ ����Ÿ�� ã�� ��ġ�� ������..      *
'*                                                              *
'*  spd          : vaSpread Name                             *
'*  Col          : Column                                       *
'*  para         : Data                                         *
'*  Return Value : ����Ÿ�� ��ġ ��                             *
'*                 ã���� �ϴ� ����Ÿ�� ���� ��� -1�� SETTING  *
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
'*  �ش� vaSpread���� �����࿡ HighLight�Ѵ�.                *
'*                                                              *
'*  spd          : vaSpread Name                             *
'*  row          : Row Value                                    *
'*  oldrow       : ���� Row Value                               *
'*  Return Value : HighList Row Value                           *
'*                                                              *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Function S0SUB_SPREAD_HIGHLIGHT(spd As vaSpread, Row As Integer, OldRow As Integer) As Integer
    
    spd.Col = 1: spd.col2 = spd.MaxCols
    spd.Row = Val(spd.Tag): spd.row2 = OldRow
    
    spd.BlockMode = True

    spd.BackColor = &HFFFFFF     '���

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
'*  vaSpread���� Ư���ฦ ȭ�鿡 �����ش�..          *
'*  spd : vaSpread Name                              *
'*  Row : Row                                           *
'*  disRow : ȭ�鿡 ������ �� �ִ� �ִ� ���� ��         *
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
'*  CurRst : �����                                                                                                          *
'*  OrdCd  : �˻��ڵ�                                                                                                      *
'*  UnitNo : �����ǹ�ȣ                                                                                                   *
'*  CurDate : �������+SYSTIME                                                                                     *
'*  NextDelta : �Ѿ�� ���� "True"�̸� �˻��ϰ� ���� ����� Delta���� �Ѱ��ش�.           *
'*  NextLabNo : ��������� LabDate+SlipCd+LabSqNo �� �Ѱ��ش�.                                 *
'*  TabelName : Select�� ��� Table
'*  < ������ "Delta Check"(SLPC110F)�� ����>                                                                *
'*  LstRst : ���������                                                                                                    *
'*  DeltaGbn : Delta Check ���                                                                                        *
'*  DeltaInterval : ��������� �������� �ð�����(Day)(�Ҽ��� 2�ڸ����� ���� ����)      *
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
        '   ���� �Էµ� ����� DeltaCheck
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
            '   ��������� DeltaCheck
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
'*  ����� ���� MODULE                  *
'*  ����ġ                                         *
'*  RetVal  : �����                           *
'*  OrdCd   : �˻��ڵ�                       *
'*  SubSqNo : SUB�˻��ڵ�              *
'*  age     : ����                                *
'*  Sex     : ����                               *
'* * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Ref_Judg(RetVal As Variant, OrdCd As String, SubSqNo As String, age As Integer, sex As String) As String

    Dim SqlData() As String
    'S_REF = ""              '����� ��½� ����ġ ����� ����
    
    If RetVal = "" Then Exit Function
    
    If SubSqNo <> "" Then

        SqlStr = "SELECT  DISTINCT REFCHK, REFCHAR2, REFNUMHI, REFNUMLO" _
                 & " FROM LAB01_DB..SLA050M " _
                 & " WHERE ORDCD = '" & OrdCd & Chr$(39) _
                 & " AND SUBSQNO = '" & SubSqNo & Chr$(39)

    Else
        If age <= 16 Then           '�Ҿ�
            SqlStr = "SELECT DISTINCT REFCHK, REFCHAR1, REFHIC, REFLOC " _
                    & " FROM LAB01_DB..SLA030M " _
                    & " WHERE ORDCD LIKE '" & OrdCd & Chr$(39)
                
        Else
            If sex = "��" Then
            
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
            
            If SqlData(1) = "1" Then               '����
                If Val(RetVal) > Val(SqlData(3)) Then
                    Ref_Judg = "H"
                ElseIf Val(RetVal) < Val(SqlData(4)) Then
                    Ref_Judg = "L"
                End If
    '            S_REF = SqlData(4) & " ~ " & SqlData(3)
            ElseIf SqlData(1) = "2" Then        '����
                If SqlData(2) <> RetVal Then Ref_Judg = "*"
    '            S_REF = SqlData(2)
            End If
            
       End If
    End If
    ret = QSqlSelectFree(QsqlCode)
         
End Function
    
'* * * * * * * * * * * * * * * * * * * * * * * * *
'*  PANIC CHECK                             *
'*  RetVal  : �����                           *
'*  OrdCd   : �˻��ڵ�                       *
'*  Return ��  "P" Or  Space              *
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
    


