Attribute VB_Name = "modDbLibrary"
Option Explicit


Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim lsPid           As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strResult       As String
    Dim strEqpCd        As String
    Dim strErrMsg       As String
    
    If gMode = 0 Then
    
    Else
        With frmInterface
            SaveTransDataW = -1
    
            lsID = Trim(GetText(.spdTot, argSpcRow, colBarcode))
            lsPid = Trim(GetText(.spdTot, argSpcRow, colPID))
    
            'Local���� ȯ�ں��� ����� ��������
            ClearSpread .vasTemp
    
                  SQL = ""
                  SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX " & vbCrLf
            SQL = SQL & "  FROM PAT_RES " & vbCrLf
            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(.spdTot, argSpcRow, colOrdDate)) & "'  " & vbCrLf
            SQL = SQL & "   AND BARCODE = '" & lsID & "' "
            
            Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
    
            If Res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
    
            .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
    
            sResult1 = ""
            sResult2 = ""
    
'            cn_Ser.BeginTrans
    
            '������ ����� �����ϱ�
            For iRow = 1 To .vasTemp.DataRowCnt
                strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
                strResult = Trim(GetText(.vasTemp, iRow, 3)) '���(�������)
                If UCase(Mid(strResult, 1, 8)) = "POSITIVE" Then
                    strResult = "Positive"
                End If
                If strResult <> "" Then
                          SQL = ""
                          SQL = "Update LisiLib.Minterface " & vbCrLf
                    SQL = SQL & "   Set Result = '" & Trim(strResult) & "'," & vbCrLf
                    SQL = SQL & "       Rltflag = 'N', " & vbCrLf
                    SQL = SQL & "       Updtdate = (select substring(char(curdate()),1,4) || substring(char(curdate()),6,2) || substring(char(curdate()),9,2) || substring(char(curtime()),4,2) || substring(char(curtime()),7,2) || substring(char(curtime()),10,2) from sysibm.sysdummy1), " & vbCrLf
                    SQL = SQL & "       Testercode = '" & gUserID & "'," & vbCrLf
                    SQL = SQL & "       Flag = '2', " & vbCrLf
                    SQL = SQL & "       Updtempl = '" & gUserID & "'" & vbCrLf
                    SQL = SQL & " Where barcodeno = '" & lsID & "'" & vbCrLf
                    SQL = SQL & "   And mcode = '" & gEquip & "'" & vbCrLf
                    SQL = SQL & "   And itemcode = '" & Mid(strEqpCd, 1, 5) & "'" & vbCrLf
                    If Len(strEqpCd) > 5 Then
                       SQL = SQL & "   And dcode = '" & Mid(strEqpCd, 6) & "'"
                    End If
                    adoTextQueryExc SQL
                    
                    '��� ������ �Ϸ�Ǹ� �ش� procedure�� call �Ѵ�.
'                     batch slrtrm55p(pmach : char(3) => ����ڵ�,
'                                                perr : char(1) => ����Ȯ�� �� �����ڵ�),
'                     real  slrtrm56p(pbarc : char(12) => ���ڵ��ȣ,
'                                        pmach : char(3) => ����ڵ�,
'                                            perr : char(1) => ����Ȯ�� �� �����ڵ�)
                    strErrMsg = adoExecQuery55P("SLRTRM55P", gEquip, "")
                    
                    
                End If
            Next iRow
    
'            cn_Ser.CommitTrans
            SaveTransDataW = 1
    
        End With
    End If
    
End Function


Function SaveTransDataR(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
''������ ����Ÿ ���̽��� ����
'    Dim iRow            As Integer
'    Dim lsID            As String
'    Dim lsPid           As String
'    Dim sResult1        As String
'    Dim sResult2        As String
'
'
'    SaveTransDataR = -1
'
'    lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode))
'    lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, colPID))
'
'    'Local���� ȯ�ں��� ����� ��������
'    ClearSpread frmInterface.vasTemp
'    With frmInterface
'        SQL = ""
'        SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICFLAG,DELTAFLAG,PSEX " & vbCrLf & _
'              "  FROM PAT_RES " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              "   AND EXAMDATE = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
'              "   AND BARCODE = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
'              "   AND DISKNO = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
'              "   AND POSNO = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
'
'        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
'
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'
'        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
'
'        sResult1 = ""
'        sResult2 = ""
'
'        cn_Ser.BeginTrans
'
'        '������ ����� �����ϱ�
'        For iRow = 1 To .vasTemp.DataRowCnt
'            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '���(�����)
'            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '���(�������)
'
'            '-- ����� ġȯ
'            sResult1 = Replace(sResult1, "<", "")
'            sResult1 = Replace(sResult1, ">", "")
'
'            If sResult1 <> "" Then
'                SQL = ""
'                SQL = SQL + "UPDATE " & gDBTBL_Parm.RSLTTABLE & vbCrLf      '-- ������̺�
'                SQL = SQL & "   SET "
'                SQL = SQL & gDBCOLUMN_Parm.RESULT & " = '" & sResult1 & "', " & vbCrLf                                      '���(�����)
'                SQL = SQL & gDBCOLUMN_Parm.RESULT & " = '" & sResult2 & "', " & vbCrLf                                      '���(�������)
'                SQL = SQL & gDBCOLUMN_Parm.MACHCD & " = '" & gEquipCode & "', " & vbCrLf                                    '����ڵ�
'                SQL = SQL & gDBCOLUMN_Parm.USER & " = '" & gEquipCode & "', " & vbCrLf                                      '����Է���
'                SQL = SQL & gDBCOLUMN_Parm.RSLTDATE & " = SysDate, " & vbCrLf                                               '����Է���
'                SQL = SQL & " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & lsID & "' " & vbCrLf                              '���ڵ��ȣ
'                SQL = SQL & "   AND " & gDBCOLUMN_Parm.TESTCD & " = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' " & vbCrLf   '�˻��ڵ�
'                SQL = SQL & "   AND " & gDBCOLUMN_Parm.PID & " = '" & lsPid & "' " & vbCrLf                                 'ȯ�ڹ�ȣ
'                SQL = SQL & "   AND " & gDBCOLUMN_Parm.STATUS & " < '2' "                                                   '�������"
'                Res = SendQuery(gServer, SQL)
'
'                If Res < 0 Then
'                    SaveQuery SQL
'                    cn_Ser.RollbackTrans
'                    Exit Function
'                End If
'
'            End If
'        Next iRow
'
'    End With
'
'    cn_Ser.CommitTrans
'    SaveTransDataR = 1
    
End Function

'-- ������ ���� ��������
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    
'    Dim sBarcode As String
'    Dim sSpecNo As String
'    Dim strAge  As String
'
'    GetSampleInfoW = -1
'
'    sBarcode = Trim(GetText(frmInterface.spdorder, asRow, colBarcode))   '2 ���� ���ڵ� ��ȣ
'
'    If sBarcode = "" Then
'        Exit Function
'    End If
'
'    '���ڵ��ȣ�� ȯ������ �ҷ�����
''    SQL = ""
''    SQL = SQL + "SELECT " & gDBCOLUMN_Parm.PID & "," & gDBCOLUMN_Parm.PNAME & "," & gDBCOLUMN_Parm.PSEX & "," & gDBCOLUMN_Parm.PAGE & vbCrLf
''    SQL = SQL + "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
''    SQL = SQL + " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
''    SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
''    SQL = SQL + "   AND " & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL"
'
''      -- ���̺� ���
'          SQL = "SELECT DiSTINCT b.SCP42IDNOA, a.SCP41NAME, a.SCP41SEX, a.SCP41BIRTH,b.SCP42SUGACD "
'    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
'    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
'    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
'    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
'    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
'    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & sBarcode & "'"
'    'SQL = SQL & vbCrLf & "   AND b.SCP42SUGACD in (" & strGumCd & ")"
'    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
'
'    '-- ����
'''          SQL = "SELECT DiSTINCT IDNO, IDNAME, Sex, BIRTHDAY "
'''    SQL = SQL & vbCrLf & "  FROM vwSPMNOINFO "
'''    SQL = SQL & vbCrLf & " WHERE SPMNO = '" & sBarcode & "'"
'''    'SQL = SQL & vbCrLf & "   AND PCODE = '60' "
'''    'SQL = SQL & vbCrLf & "   AND SUGACD in (" & strGumCd & ")"
'
''vwSPMNOINFO
''
''PCODE �˻���Ʈ
''JDATE �˻���
''SPMSID  seq �ѹ�
''SPMNO ���óѹ�(���ڵ�ѹ�)
''IDNO    ��Ʈ��ȣ(�⺻ 7�ڸ� + Ÿ������ȣ 1�ڸ�, Ÿ����ȯ�ڰ� �ƴѰ�� ' ' ��)
''IDNAME ȯ���̸�
''KWA ��
''WARD ����
''Sex ����
''BIRTHDAY ����
''SUGACD �˻�����ڵ�
''SUGANM �˻������Ī
''RESULTYN �������
''SENDYN      ����뺸 ����
''SPMTIME ��ü�Ͻ�
'
'    Res = GetDBSelectColumn(gServer, SQL)
'
'    If Res = 1 Then
'        SetText frmInterface.spdorder, Trim(gReadBuf(0)), asRow, colPID    '5
'        SetText frmInterface.spdorder, Trim(gReadBuf(1)), asRow, colPName  '6
'        SetText frmInterface.spdorder, Trim(gReadBuf(2)), asRow, colSex    '7
'        strAge = Format(Now, "yyyy") - Mid(Trim(gReadBuf(3)), 1, 4)
'        SetText frmInterface.spdorder, strAge, asRow, colAge    '8
'        GetSampleInfoW = 1
'    Else
'        GetSampleInfoW = -1
'    End If

End Function

Function GetSampleInfoR(ByVal asRow As Long) As Integer
'    Dim sBarcode As String
'    Dim sSpecNo As String
'
'    GetSampleInfoR = -1
'
'    'ȯ������ ��������
'    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBarcode))   '���� ���ڵ� ��ȣ
'
'    If sBarcode = "" Then
'        Exit Function
'    End If
'
'    '���ڵ��ȣ�� ȯ������ �ҷ�����
'
'    SQL = ""
'    SQL = SQL + "SELECT " & gDBCOLUMN_Parm.PID & "," & gDBCOLUMN_Parm.PNAME & "," & gDBCOLUMN_Parm.PSEX & "," & gDBCOLUMN_Parm.PAGE + vbLf
'    SQL = SQL + "  FROM " & gDBTBL_Parm.ORDTABLE + vbLf
'    SQL = SQL + " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " + vbLf
'    SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " + vbLf
'    SQL = SQL + "   AND " & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL" + vbLf
'
'    Res = GetDBSelectColumn(gServer, SQL)
'
'    If Res = 1 Then
'        SetText frmInterface.spdorder, Trim(sSpecNo), asRow, colSpecNo
'        SetText frmInterface.spdorder, Trim(gReadBuf(0)), asRow, colPID
'        SetText frmInterface.spdorder, Trim(gReadBuf(1)), asRow, colPName
'        SetText frmInterface.spdorder, Trim(gReadBuf(2)), asRow, colSex
'        SetText frmInterface.spdorder, Trim(gReadBuf(3)), asRow, colAge
'
'        GetSampleInfoR = 1
'    Else
'
'        GetSampleInfoR = -1
'    End If
End Function

Function GetEquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String

    GetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    Res = GetDBSelectVas(gLocal, SQL, frmInterface.vasTemp1)
    
    If frmInterface.vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To frmInterface.vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
        End If
    Next i

    'SPSLHRRST
    SQL = " Select SUCD From LRESULT " & vbCr & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    Res = GetDBSelectColumn(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        GetEquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function


Function GetGetEquipExamCode_CA1500(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim strStatFg  As String
Dim sExamCd As String
Dim strItems As String
Dim strTemp As String
Dim strIntBase As String

    GetGetEquipExamCode_CA1500 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
'    argPID = "1558200030"
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    Res = GetDBSelectColumn(gServer, SQL)
    GetGetEquipExamCode_CA1500 = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            sExamCd = Trim(gReadBuf(i))
        Else
            Exit For
        End If
    Next
    
    SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and SUBSTR(exmn_cd,1,1) <> 'G'" & _
          "   and RSLT_NO IS NOT NULL"
          
    Res = GetDBSelectRow(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    Erase gReadBuf
          SQL = "Select equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    SQL = SQL & " order by equipcode    "
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strIntBase = Trim(gReadBuf(i))
            strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1) & "0" & Space$(6)
            If strIntBase <> strTemp Then
                strExamCode = strExamCode & strIntBase 'Mid(Trim(gReadBuf(i)), 1, Len(Trim(gReadBuf(i))) - 1) & "0" & Space$(6)
                strTemp = strIntBase
            End If

            'strExamCode = strExamCode & Mid(Trim(gReadBuf(i)), 1, Len(Trim(gReadBuf(i))) - 1) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    '�������� (R:Routin, E:Stat)
    'strStatFg = IIf(pAccInfo.StatFg = "1", "E", "U")
    strStatFg = "U"
    
    
'    strExamCode = STX & "S2210101" & strStatFg & Space(6) & Space(4) & mOrder.RackNo & mOrder.TubePos & mOrder.BarNo & _
                "B" & Space(15) & strExamCode & ETX
    
    strExamCode = "" & "S2210101" & strStatFg & Space(6) & Space(4) & mResult.RackNo & mResult.TubePos & mResult.BarNo & _
                "B" & Space(15) & strExamCode & ""
    
    GetGetEquipExamCode_CA1500 = strExamCode
    
End Function

Function GetOrderExamCode(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
    Dim i           As Integer
    Dim sExamCode   As String
    Dim strExamCode As String
    Dim sExamCd     As String
    Dim adoRS2 As ADODB.Recordset

    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '-- �˻��ڵ� ��������
    Set adoRS2 = New ADODB.Recordset
    Set adoRS2 = adoExecQuery51P("SLRTRM52P", Trim(argPID), gEquipCode, "")
    
    GetOrderExamCode = ""
    
    Select Case strRecordStatus
        Case "R"
            'lblStatus.Caption = Trim(argPID) & " ���ڵ� ���� ! ���ڵ��ȣ�� Ȯ���ϼ���."
            'adoRS2.Close: Set adoRS2 = Nothing ': Exit Sub
        Case "M"
            'lblStatus.Caption = Trim(argPID) & " ����ڵ� ���� !  "
            'adoRS2.Close: Set adoRS2 = Nothing ': Exit Sub
        Case "Y", "N", " "
            'lblStatus.Caption = Trim(argPID) & " �˻�����."
            If Not adoRS2.EOF Then
                Do While Not adoRS2.EOF
                    GetOrderExamCode = GetOrderExamCode & "'" & Trim$(adoRS2("ITEMCODE")) & Trim$(adoRS2("DCODE")) & "',"
                    adoRS2.MoveNext
                Loop
            End If
    End Select
    
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
    End If
    
    adoRS2.Close
    Set adoRS2 = Nothing
    
End Function

Function GetOrderExamCode_New(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

    GetOrderExamCode_New = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    argPID = Mid(argPID, 1, 10)
    
          SQL = "SELECT DiSTINCT b.SCP42SUGACD "
    SQL = SQL & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
    SQL = SQL & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
    SQL = SQL & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
    SQL = SQL & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = b.SCP42SPMNO2 "
    SQL = SQL & vbCrLf & "   AND a.SCP41SPMNO2 = '" & argPID & "'"
    SQL = SQL & vbCrLf & "   AND b.SCP42RESULT IS NULL "
    
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_New = GetOrderExamCode_New & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_New <> "" Then
        GetOrderExamCode_New = Mid(GetOrderExamCode_New, 1, Len(GetOrderExamCode_New) - 1)
    End If
    
End Function


Function GetGetEquipExamCode_E411(argEquipCode As String, argPID As String, Optional intRow As Long) As String
''��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
''�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
'Dim i As Integer
'Dim sExamCode As String
'Dim strExamCode As String
'Dim sSpecNo     As String
'Dim iRow        As Long
'Dim SpecNo      As String
'
'    GetGetEquipExamCode_E411 = ""
'
'    If Trim(argEquipCode) = "" Then
'        Exit Function
'    End If
'
'    '-- �ڰ�ü�� 11�ڸ��� ��ȸ�ϱ����Ͽ� ������ �ڸ��� ���ش�.
'    argPID = Mid(argPID, 1, 10)
'
'    If Mid(argPID, 1, 2) = "99" Then
'        'strExamCode = Proc_Order_LX_QC(argPID)
'
'        'iRow = frmInterface.spdorder.DataRowCnt
'        iRow = intRow
'
'        SpecNo = Trim(GetText(frmInterface.spdorder, iRow, colSpecNo))
'
'        SQL = "SELECT QC_EXMN_CD "
'        SQL = SQL & vbCrLf & " FROM SPSLMQMST "
'        SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(SpecNo, 3, 3) & "' "     '//// ��� ��ȣ
'        SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(SpecNo, 6, 3) & "' "     '//// �˻�� ��ȣ
'        SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(SpecNo, 9, 1) & "' "      '//// ���� ��ȣ
'        SQL = SQL & vbCrLf & "  AND QC_EXMN_CD IN (" & gAllExam & ") "
'        SQL = SQL & vbCrLf & "  AND USE_STR_DT <= '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' "
'        SQL = SQL & vbCrLf & "  AND USE_END_DT >= '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' "
'        Res = GetDBSelectRow(gServer, SQL)
'        strExamCode = ""
'
'        For i = 0 To UBound(gReadBuf)
'            If gReadBuf(i) <> "" Then
'                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
'            Else
'                Exit For
'            End If
'        Next
'
'    Else
'        '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
'        SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
'        Res = GetDBSelectColumn(gServer, SQL)
'        sSpecNo = Trim(gReadBuf(0))
'
'        '-- �˻��ڵ� ��������
'        SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
'              " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
'              "   and RSLT_NO IS NOT NULL"
'
'        Res = GetDBSelectRow(gServer, SQL)
'        strExamCode = ""
'
'        For i = 0 To UBound(gReadBuf)
'            If gReadBuf(i) <> "" Then
'                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
'            Else
'                Exit For
'            End If
'        Next
'    End If
'
'    If strExamCode = "" Then
''        MsgBox "������ ȯ��"
'        GetGetEquipExamCode_E411 = ""
'        Exit Function
'    End If
'    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
'    'GetEquipExamCode =
'
'    ClearSpread frmInterface.vasTemp1
''    sExamCode = ""
'
'    '-- ������ �˻��ڵ��� ä�� ã��
'          SQL = "Select distinct equipcode "
'    SQL = SQL & "  From EquipExam "
'    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
'    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
'
'    Res = GetDBSelectRow(gLocal, SQL)
'    strExamCode = ""
'    For i = 0 To UBound(gReadBuf)
'
'        If gReadBuf(i) <> "" Then
'            'gReadBuf(i) = Mid(gReadBuf(i), 1, Len(gReadBuf(i)) - 1)
'            If Trim(gReadBuf(i)) <> "990" Then
'                strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
'            End If
'        Else
'            Exit For
'        End If
'    Next
'
'    GetGetEquipExamCode_E411 = Mid(strExamCode, 2)
    
End Function



Function GetGetEquipExamCode_Architect(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'    Dim i As Integer
'    Dim strExamCode As String
'    Dim sBarcode     As String
'
'    GetGetEquipExamCode_Architect = ""
'
'    If Trim(argEquipCode) = "" Then
'        Exit Function
'    End If
'
'    sBarcode = Trim(GetText(frmInterface.spdorder, intRow, colBarcode))   '2 ���� ���ڵ� ��ȣ
'
'    If sBarcode = "" Then
'        Exit Function
'    End If
'
'    '-- �˻��ڵ� ��������
''    SQL = ""
''    SQL = SQL + "SELECT " & gDBCOLUMN_Parm.TESTCD & vbCrLf
''    SQL = SQL + "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
''    SQL = SQL + " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
''    SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
''    SQL = SQL + "   AND " & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL"
''
''    Res = GetDBSelectRow(gServer, SQL)
''    strExamCode = ""
''
''    For i = 0 To UBound(gReadBuf)
''        If gReadBuf(i) <> "" Then
''            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
''        Else
''            Exit For
''        End If
''    Next
''
''    If strExamCode = "" Then
''        '-- ������ȯ���̰ų� �ش���� �˻��� ����
''        GetGetEquipExamCode_Architect = ""
''        Exit Function
''    End If
''
''    '-- ������ "," �ڸ���
''    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
'
'    ClearSpread frmInterface.vasTemp1
'
'    '-- ������ �˻��ڵ��� ä�� ã��
'    SQL = "          "
'    SQL = SQL & "SELECT Distinct EQUIPCODE "
'    SQL = SQL & "  FROM EQUIPEXAM "
'    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
'    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
'
'    Res = GetDBSelectRow(gLocal, SQL)
'    strExamCode = ""
'
'    '-- �ش� ��� �°� ����ä�� �����ϱ� [ASTM Format >> Architect]
'    For i = 0 To UBound(gReadBuf)
'        If gReadBuf(i) <> "" Then
'            If Trim(gReadBuf(i)) <> "990" Then
'                strExamCode = strExamCode & Trim(gReadBuf(i))
'            End If
'        Else
'            Exit For
'        End If
'    Next
'
'    '-- ù�ڸ� "\" �ڸ���
'    GetGetEquipExamCode_Architect = strExamCode
    
End Function


Function GetGetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'    Dim i As Integer
'    Dim strExamCode As String
'    Dim sBarcode     As String
'
'    GetGetEquipExamCode_AU480 = ""
'
'    If Trim(argEquipCode) = "" Then
'        Exit Function
'    End If
'
'    sBarcode = Trim(GetText(frmInterface.spdorder, intRow, colBarcode))   '2 ���� ���ڵ� ��ȣ
'
'    If sBarcode = "" Then
'        Exit Function
'    End If
'
'
'    ClearSpread frmInterface.vasTemp1
'
'    '-- ������ �˻��ڵ��� ä�� ã��
'    SQL = "          "
'    SQL = SQL & "SELECT Distinct EQUIPCODE "
'    SQL = SQL & "  FROM EQUIPEXAM "
'    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
'    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
'
'    Res = GetDBSelectRow(gLocal, SQL)
'    strExamCode = ""
'
'    For i = 0 To UBound(gReadBuf)
'        If gReadBuf(i) <> "" Then
'            'If Trim(gReadBuf(i)) <> "990" Then
'                '                                                     dilution
'                strExamCode = strExamCode & "0" & Trim(gReadBuf(i)) & "0"
'            'End If
'        Else
'            Exit For
'        End If
'    Next
'
'    GetGetEquipExamCode_AU480 = strExamCode
    
End Function


Function GetGetEquipExamCode(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String

    GetGetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '-- �ڰ�ü�� 11�ڸ��� ��ȸ�ϱ����Ͽ� ������ �ڸ��� ���ش�.
    argPID = Mid(argPID, 1, 10)
    
    SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
          " Where SPCM_NO = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    Res = GetDBSelectColumn(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
          SQL = "Select equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(argEquipCode) & ")"
    
    Res = GetDBSelectColumn(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & Trim(gReadBuf(i)) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    GetGetEquipExamCode = strExamCode
    
End Function


