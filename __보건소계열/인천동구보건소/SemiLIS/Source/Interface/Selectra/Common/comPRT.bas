Attribute VB_Name = "comPRT"
Option Explicit

Public Const D0PRT_FontName = "����ü"

Public D0PRT_FontSize         As Integer

Public D0PRT_Title00          As String   '��� Title
Public D0PRT_Title01          As String   'Left Title
Public D0PRT_Title02          As String   'Left Title

Public D0PRT_SubTitle()       As String   '�׸� title
Public D0PRT_SubTitleCount    As Integer  '�׸� title�� ����
Public D0PRT_CurrentX()       As Long   '�����ġ

Public D0PRT_Page             As Integer  'Page Number
Public D0PRT_LineLen          As Integer  'Line ����
Public D0PRT_LeftSpace        As Integer  '���� ����
Public D0PRT_SpaceLen         As Integer

Public D0PRT_Item01           As String
Public D0PRT_Item02           As String
Public D0PRT_Item03           As String
Public D0PRT_ITem04           As String
Public D0PRT_ITem05           As String



Function D0SUB_Print_BaseCode_DJA080M(ByVal SqlConn As Integer, ByVal code As String)

    Dim sql_ret As Integer, SqlDoc  As String
    Dim record  As String, Cod1()   As String
    Dim Line_Count  As Integer
    Dim curY    As Long

    Dim slipcd  As String
    
    Line_Count = 1
    
    SqlDoc = "SELECT DISTINCT A7.ORDCD,  A7.SEQNO, A6.ORDNM, A3.CDGBNM" _
           + "  FROM LAB01_DB..DJA070M A7, LAB01_DB..DJA060M A6" _
           + "     , LAB01_DB..DJA030M A3" _
           + " WHERE A7.CDGBN = 'P' AND A7.RTNCD = " & Chr(39) & code & Chr(39) _
           + "   AND A7.ORDCD = A6.ORDCD" _
           + "   AND A3.LCGBCD = '02'" _
           + "   AND SUBSTRING(A7.ORDCD,1,2) = A3.SCGBCD"
    sql_ret = QSqlDBExec(SqlDoc, SqlConn)
    If sql_ret = QSQL_SUCCESS Then
        Do Until QSqlGetRow(record, SqlConn) <> QSQL_SUCCESS

            '/* ����Ÿ �б�
            QSqlGetField 4, record, Cod1()
            
            '/* Title�μ�
            If Printer.currenty > Printer.ScaleHeight - Printer.TextHeight(" ") * 7 _
               Or D0PRT_Page = 1 Then
                curY = D0SUB_Print_Title(5, 4)
                Line_Count = 1
            End If
                
            '/* 5 Line �μ��� Line Skip
            If (Line_Count Mod 5) = 1 And Not (curY = Printer.currenty Or curY = 0) Then Printer.Print
                
            If slipcd <> Mid$(Cod1(1), 1, 2) Then
                slipcd = Mid$(Cod1(1), 1, 2)
            Else
                Cod1(4) = ""
            End If
            
            Printer.currentx = D0PRT_CurrentX(4):   Printer.Print Cod1(2);
            Printer.currentx = D0PRT_CurrentX(5):   Printer.Print Cod1(4);
            Printer.currentx = D0PRT_CurrentX(6):   Printer.Print Cod1(1) + "  " + Cod1(3)
            
            Line_Count = Line_Count + 1
        Loop
    End If
    Call QSqlSelectFree(SqlConn)

    
End Function


'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*                                                                          *
'*  ����� : 1 Line�� ����� �� �ִ� �ڵ���¸� ���                      *
'*  ReDim D0PRT_SUBTITLE(1 To 3) As String                                  *
'*  D0PRT_LineLen = 90 (line�� ���� ����)                                   *
'*  D0PRT_TITLE = "��ü�ڵ� ����Ʈ" (Title ����)                            *
'*  D0PRT_SUBTITLE(1) = "�ڵ�"      (sub title �׸� : �ڸ�������  )         *
'*  D0PRT_SUBTITLE(2) = "��ü���"  (sub title �׸� : �ڸ�������)           *
'*  D0PRT_SUBTITLE(3) = "��ü��" + Space(24)  (sub title �׸� : �ڸ�������) *
'*  D0PRT_SubTitleCount = 3   (Field ���� )                                 *
'*                                                                          *
'*  flg     : 1=�������,    3=�Ϲ��ڵ�, 4=�����ڵ�, 5=�ο��ڵ�, 6=�˻��ڵ� *
'*            7=Routine�ڵ�, 8=����ڵ�, 9=�����ȣ                         *
'*  sqldoc  : sql����                                                       *
'*  sqlconn : Server Open Connect                                           *
'*                                                                          *
'*  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Function D0SUB_Print_BaseCode(frm As Form, ByVal FLG As String, ByVal SqlCode As Integer _
                            , Optional ByVal SqlDoc As String _
                            , Optional ByVal sql_where _
                            , Optional ByVal check As Variant) _

    Dim sql_ret As Integer, SqlConn As Integer
    Dim record  As String, code()   As String
    Dim col     As Integer
    Dim Line_Count  As Integer
    Dim curY    As Long

    On Error GoTo PRINT_BASE_CODE

    D0PRT_Page = 1: Line_Count = 1
    D0PRT_LeftSpace = 20: D0PRT_SpaceLen = 2
    D0PRT_LineLen = 140
    
    Printer.EndDoc
    
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = 2

    If FLG = 4 Then '�����ڵ� ���
        SqlDoc = "SELECT CODEGB+SUGACD, SUGANO, SUGANM, SCOST1, SCOST2" _
               + "     , SDATE,  EDATE,  DAYFLG, PHYCHK" _
               + "  FROM LAB01_DB..DJA040M" _
               + " WHERE " & sql_where
            
        sql_ret = QSqlDBExec(SqlDoc, SqlCode)
        If sql_ret = QSQL_SUCCESS Then
            Do Until QSqlGetRow(record, SqlCode) <> QSQL_SUCCESS
    
                '/* ����Ÿ �б�
                QSqlGetField 9, record, code()
                
                If code(4) <> "" Then code(4) = Format$(Format$(code(4), "###,###"), "@@@@@@@")
                If code(5) <> "" Then code(5) = Format$(Format$(code(5), "###,###"), "@@@@@@@")
                If code(8) <> "" Then code(8) = code(8) + "��"
                
                '/* Title�μ�
                If Printer.currenty > Printer.ScaleHeight - Printer.TextHeight(" ") * 7 _
                    Or D0PRT_Page = 1 Then curY = D0SUB_Print_Title
                    
                '/* 5 Line �μ��� Line Skip
                If (Line_Count Mod 5) = 1 And Not curY = Printer.currenty Then Printer.Print
                
                '/* ����Ÿ �μ�
                Printer.Print Spc(D0PRT_LeftSpace); Format$(Line_Count, "###");
                For col = 1 To D0PRT_SubTitleCount
                    Printer.currentx = D0PRT_CurrentX(col):   Printer.Print code(col);
                Next
                Printer.Print
                
                Line_Count = Line_Count + 1
            Loop
        End If
    ElseIf FLG = 5 Then '�ο��ڵ� ���
        '�� ȭ�鿡�� ����� Index Open
        sql_ret = QSqlOpen(D0COM_SERVER01, frm.hWnd, SqlConn)
        If sql_ret <> QSQL_SUCCESS Then Exit Function
    
        SqlDoc = "SELECT DISTINCT A5.OFFCD,  A5.OFFNM, A5.SUGACD, A5.DAYFLG, A3.CDGBNM" _
               + "  FROM LAB01_DB..DJA050M A5, LAB01_DB..DJA040M A4" _
               + "     , LAB01_DB..DJA030M A3" _
               + " WHERE A3.LCGBCD  = '06'" _
               + "   AND A5.CODEGB *= A3.SCGBCD" _
               + "   AND A5.SUGACD *= A4.SUGACD"
        If IsMissing(sql_where) = False Then _
           SqlDoc = SqlDoc + " AND A5.CODEGB = " & Chr(39) & sql_where & Chr(39)
            
        sql_ret = QSqlDBExec(SqlDoc, SqlCode)
        If sql_ret = QSQL_SUCCESS Then
            Do Until QSqlGetRow(record, SqlCode) <> QSQL_SUCCESS
    
                '/* ����Ÿ �б�
                QSqlGetField 5, record, code()
                
                If code(4) <> "" Then code(4) = code(4) + "��"
                
                '/* Title�μ�
                If Printer.currenty > Printer.ScaleHeight - Printer.TextHeight(" ") * 7 _
                    Or D0PRT_Page = 1 Then curY = D0SUB_Print_Title(5, 4)
                    
                '/* ����Ÿ �μ�
                Printer.Print Spc(D0PRT_LeftSpace); Format$(Line_Count, "###");
                Printer.currentx = D0PRT_CurrentX(1):   Printer.Print code(1);
                Printer.currentx = D0PRT_CurrentX(2):   Printer.Print code(2);
                Printer.currentx = D0PRT_CurrentX(3):   Printer.Print code(3);
                If code(4) <> "" Then
                    Printer.currentx = D0PRT_CurrentX(4):   Printer.Print code(4);
                    Printer.currentx = D0PRT_CurrentX(5):   Printer.Print code(5)
                End If
                
                Call D0SUB_Print_BaseCode_DJA080M(SqlConn, sql_where + code(1))
                
                Line_Count = Line_Count + 1
            Loop
        End If
        Call Qsqlclose(SqlConn, ONECLOSE)
    ElseIf FLG = 6 Then '�˻��ڵ� ���
    
        D0PRT_Title00 = "�˻��׸��ڵ� ���"
        D0PRT_Title01 = "": D0PRT_Title02 = "�� : "
        D0PRT_Item01 = "": D0PRT_Item02 = ""
        
        ReDim D0PRT_SubTitle(1 To 9) As String
        ReDim D0PRT_CurrentX(1 To 9) As Long
        
        D0PRT_SubTitle(1) = "����"
        D0PRT_SubTitle(2) = "�ڵ�" + Space(7)
        D0PRT_SubTitle(3) = "SLIP��" + Space(14)
        D0PRT_SubTitle(4) = "�˻��" + Space(24)
        D0PRT_SubTitle(5) = "���� "
        D0PRT_SubTitle(6) = "�˻����" + Space(2)
        D0PRT_SubTitle(7) = "����ġ" + Space(17)
        D0PRT_SubTitle(8) = "Panic��"
        D0PRT_SubTitle(9) = "Delta��"
    
        D0PRT_SubTitleCount = 9
        
        Dim slipcd  As String
        Dim deptcd  As String
        
        SqlDoc = "SELECT A6.DEPTCD, A6.ORDCHK, SUBATRING(A6.ORDCD,1,2), SUBSTRING(A6.ORDCD,3,5), A6.SUBCD" _
               + "     , A6.ORDNM,  A6.SUGACD, A6.ORDUNIT, A6.REFCHK,  A6.REFCHAR" _
               + "     , A6.REFLOM, A6.REFHIM, A6.REFLOF,  A6.REFHIF,  A6.REFLOC" _
               + "     , A6.REFHIC, A6.PANVAL, A6.DELVAL,  A31.CDGBNM, A32.CDGBNM" _
               + "  FROM LAB01_DB..DJA060M A6, LAB01_DB..DJA030M A31" _
               + "     , LAB01_DB..DJA030M A32" _
               + " WHERE A31.LCGBCD = '01' AND A6.DEPTCD = A31.SCGBCD" _
               + "   AND A32.LCGBCD = '02' AND A6.SLIPCD = A32.SCGBCD"
        If IsMissing(sql_where) = False Then SqlDoc = SqlDoc + " AND " + sql_where
        SqlDoc = SqlDoc + " ORDER BY A6.DEPTCD, SUBSTRING(A6.ORDCD,1,2), SUBSTRING(A6.ORDCD,3,5)"
        
        sql_ret = QSqlDBExec(SqlDoc, SqlCode)
        If sql_ret = QSQL_SUCCESS Then
            Do Until QSqlGetRow(record, SqlCode) <> QSQL_SUCCESS
    
                '/* ����Ÿ �б�
                QSqlGetField 20, record, code()
                
                If code(5) <> "" Then code(5) = "-" + code(5)
                
                If deptcd <> code(1) Then
                    deptcd = code(1)
                    D0PRT_Item02 = code(19)
                    curY = D0SUB_Print_Title
                    Line_Count = 1
                End If
                
                If slipcd <> code(3) Then
                    slipcd = code(3)
                Else
                    code(20) = ""
                End If
                
                '/* Title�μ�
                If Printer.currenty > Printer.ScaleHeight - Printer.TextHeight(" ") * 7 _
                    Or D0PRT_Page = 1 Then curY = D0SUB_Print_Title(5, 4)
                    
                '/* ����Ÿ �μ�
                Printer.Print Spc(D0PRT_LeftSpace); Format$(Line_Count, "###");
                Printer.currentx = D0PRT_CurrentX(1):  Printer.Print code(2);
                Printer.currentx = D0PRT_CurrentX(2):  Printer.Print code(3) + "-" + code(4) + code(5);
                Printer.currentx = D0PRT_CurrentX(3):  Printer.Print code(20);
                Printer.currentx = D0PRT_CurrentX(4):  Printer.Print code(6);
                Printer.currentx = D0PRT_CurrentX(5):  Printer.Print code(7);
                Printer.currentx = D0PRT_CurrentX(6):  Printer.Print code(8);
                If code(9) = "C" Then
                    Printer.currentx = D0PRT_CurrentX(7):  Printer.Print Space(4) + code(10)
                ElseIf code(9) = "N" Then
                    Printer.currentx = D0PRT_CurrentX(7):  Printer.Print "�� " + Format$(code(11), "@@@@@@@@") + " - " + Format$(code(12), "@@@@@@@@");
                    Printer.currentx = D0PRT_CurrentX(8):  Printer.Print Format$(code(17), "@@@@@@");
                    Printer.currentx = D0PRT_CurrentX(9):  Printer.Print Format$(code(18), "@@@@@@")
                    
                    Printer.currentx = D0PRT_CurrentX(7):  Printer.Print "�� " + Format$(code(13), "@@@@@@@@") + " - " + Format$(code(14), "@@@@@@@@")
                    Printer.currentx = D0PRT_CurrentX(7):  Printer.Print "�� " + Format$(code(15), "@@@@@@@@") + " - " + Format$(code(16), "@@@@@@@@")
                Else
                    Printer.Print
                End If
                Printer.Print
                
                Line_Count = Line_Count + 1
            Loop
        End If
        Call Qsqlclose(SqlConn, ONECLOSE)
    
    ElseIf FLG = 7 Then '��ƾ�ڵ� ���
        
        Dim profcd  As String
        Dim cnt     As Integer
        
        cnt = 1
        SqlDoc = "SELECT A.RTNCD, B.ORDNM, A.ORDCD, C.ORDNM" _
               + "  FROM LAB01_DB..DJA070M A, LAB01_DB..DJA060M B" _
               + "     , LAB01_DB..DJA060M C" _
               + " WHERE A.CDGBN = 'R' AND A.RTNCD = B.ORDCD" _
               + "   AND A.ORDCD = C.ORDCD"
        sql_ret = QSqlDBExec(SqlDoc, SqlCode)
        If sql_ret = QSQL_SUCCESS Then
            Do Until QSqlGetRow(record, SqlCode) <> QSQL_SUCCESS
    
                '/* ����Ÿ �б�
                QSqlGetField 4, record, code()
                
                '/* Title�μ�
                If Printer.currenty > Printer.ScaleHeight - Printer.TextHeight(" ") * 7 _
                   Or D0PRT_Page = 1 Then
                     curY = D0SUB_Print_Title
                    cnt = 1
                End If
                
                If profcd <> code(1) Then
                    profcd = code(1)
                    If cnt <> 1 Then Printer.Print
                    cnt = 1
                    Line_Count = Line_Count + 1
                    '/* ����Ÿ �μ�
                    Printer.Print Spc(D0PRT_LeftSpace); Format$(Line_Count - 1, "###");
                Else
                    code(2) = ""
                End If
                
                Printer.currentx = D0PRT_CurrentX(1):   Printer.Print code(2);
                Printer.currentx = D0PRT_CurrentX(2):   Printer.Print code(3);
                Printer.currentx = D0PRT_CurrentX(3):   Printer.Print code(4)
                
                cnt = cnt + 1
                
                '/* 5 Line �μ��� Line Skip
                If (cnt Mod 5) = 1 And Not curY = Printer.currenty Then Printer.Print
            
            Loop
        End If
    
    Else
        sql_ret = QSqlDBExec(SqlDoc, SqlCode)
        If sql_ret = QSQL_SUCCESS Then
            Do Until QSqlGetRow(record, SqlCode) <> QSQL_SUCCESS
    
                '/* ����Ÿ �б�
                QSqlGetField D0PRT_SubTitleCount, record, code()
    
                '/* Title�μ�
                If Printer.currenty > Printer.ScaleHeight - Printer.TextHeight(" ") * 7 _
                    Or D0PRT_Page = 1 Then curY = D0SUB_Print_Title
    
                '/* 5 Line �μ��� Line Skip
                If (Line_Count Mod 5) = 1 And Not curY = Printer.currenty Then Printer.Print
    
                '/* ����Ÿ �μ�
                Printer.Print Spc(D0PRT_LeftSpace); Format$(Line_Count, "###");
                For col = 1 To D0PRT_SubTitleCount
                    Printer.currentx = D0PRT_CurrentX(col):   Printer.Print code(col);
                Next
                Printer.Print
                Line_Count = Line_Count + 1
            Loop
        End If
    End If
PRINT_BASE_CODE_EXIT:
    Call QSqlSelectFree(SqlCode)

    Printer.EndDoc
    Exit Function

PRINT_BASE_CODE:

    If Err = 482 Then
        MsgBox "Printer Error!!", MB_OK
        Resume PRINT_BASE_CODE_EXIT
    Else
        Resume Next
    End If

End Function

Function D0SUB_Print_Title(Optional ByVal col As Variant _
                         , Optional ByVal sPos As Variant) As Integer

    Dim idx As Integer
    Dim currentx As Long, currenty As Long
    Dim strLen  As Integer

    If D0PRT_Page <> 1 Then Printer.NewPage
    
    Printer.FontName = D0PRT_FontName
    Printer.FontSize = D0PRT_FontSize * 1.5

    Printer.FontUnderline = True
    Printer.Print Spc(D0PRT_LeftSpace); Spc(Int(D0PRT_LineLen / 4) - Int(Len(D0PRT_Title00) / 2)); D0PRT_Title00
    Printer.FontUnderline = False
    Printer.Print
    
    Printer.FontSize = D0PRT_FontSize

    strLen = Len(D0PRT_Item01)
    Printer.Print Spc(D0PRT_LeftSpace); D0PRT_Title01; D0PRT_Item01;
    Printer.Print Spc(D0PRT_LineLen - Len(D0PRT_Title01) - Len(D0PRT_Item01) - 18);
    currentx = Printer.currentx
    Printer.Print "  �� : "; Format$(D0PRT_Page, "###")
    
    Printer.Print Spc(D0PRT_LeftSpace); D0PRT_Title02; D0PRT_Item02; Spc(10); D0PRT_Item03;
    Printer.currentx = currentx
    Printer.Print "��¥ : "; Format$(Now, "YYYY.MM.DD")

    Printer.Print Spc(D0PRT_LeftSpace); String(D0PRT_LineLen, Chr(6))

    Printer.Print Spc(D0PRT_LeftSpace); "��ȣ"; Spc(D0PRT_SpaceLen);
    If IsMissing(col) Then
        For idx = 1 To D0PRT_SubTitleCount
            D0PRT_CurrentX(idx) = Printer.currentx
            Printer.Print D0PRT_SubTitle(idx); Spc(D0PRT_SpaceLen);
        Next
        Printer.Print
    Else
        For idx = 1 To col
            D0PRT_CurrentX(idx) = Printer.currentx
            Printer.Print D0PRT_SubTitle(idx); Spc(D0PRT_SpaceLen);
        Next
        Printer.Print
        For idx = col + 1 To D0PRT_SubTitleCount
            If idx = col + 1 Then
                Printer.currentx = D0PRT_CurrentX(sPos)
            Else
                D0PRT_CurrentX(sPos) = Printer.currentx
            End If
            Printer.Print D0PRT_SubTitle(idx); Spc(D0PRT_SpaceLen);
            sPos = sPos + 1
        Next
        Printer.Print
    End If
    Printer.Print Spc(D0PRT_LeftSpace); String(D0PRT_LineLen, Chr(6))
    
    currenty = Printer.currenty
    
    Printer.currenty = Printer.ScaleHeight - Printer.TextHeight(" ") * 4
    Printer.Print Spc(D0PRT_LeftSpace); String(D0PRT_LineLen, Chr(6))
    Printer.Print Spc(D0PRT_LeftSpace); Space(D0PRT_LineLen - 14); "���۱� ���Ǽ�"

    Printer.currenty = currenty
    D0SUB_Print_Title = currenty
    
    D0PRT_Page = D0PRT_Page + 1

End Function
