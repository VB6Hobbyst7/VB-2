Attribute VB_Name = "modDbLibrary"
Option Explicit


Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim VallsID            As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim strDate         As String
    
    
    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        VallsID = Val(lsID)
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strDate = Format(CDate(.dtpToday.Value), "yyyymmdd")
        
        '-- Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '����ڵ�
        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf                                        '�˻���
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf     '���ڵ�
        'SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf         'DISK ��ȣ
        'SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "                    'POS ��ȣ
              
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        cn_Ser.BeginTrans
        
        '-- ������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '���(�����)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '���(�������)
            
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If sResult <> "" Then
                      SQL = "UPDATE " & gDB_Parm.DB & "..JUN370_RESULTTB"
                SQL = SQL & "   SET RESULT = '" & sResult & "'"
                SQL = SQL & " WHERE WAITSEQNO = '" & lsID & "'"
                SQL = SQL & "   AND MAP2SEQNO = '" & strEqpCd & "'"
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        cn_Ser.CommitTrans
        SaveTransDataW = 1
    
    End With

End Function


Function SaveTransDataR(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'������ ����Ÿ ���̽��� ����
    Dim iRow            As Integer
    Dim lsID            As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim VallsID         As String
    Dim strDate         As String

    SaveTransDataR = -1
    
    'Local���� ȯ�ں��� ����� ��������
    ClearSpread frmInterface.vasTemp
    
    With frmInterface
        lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, 2))
        VallsID = lsID
        lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, 5))
        strDate = Format(CDate(.dtpExamDate.Value), "yyyymmdd")
        
        '-- Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                            '����ڵ�
        SQL = SQL & "   AND EXAMDATE = '" & strDate & "'  " & vbCrLf   '�˻���
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.vasRID, argSpcRow, 2)) & "' " & vbCrLf     '���ڵ�
        'SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf         'DISK ��ȣ
        'SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "                    'POS ��ȣ
                
        Res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If Res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
                
        cn_Ser.BeginTrans
        
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '���(�����)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '���(�������)
            
            '-- ���������
            If .optSaveResultR(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If sResult <> "" Then
'                If Len(VallsID) > 6 Then
'                    SQL = "Update ONIT..GUMJIN_INTERFACE" & _
'                          "   Set RESULT = '" & sResult & "'," & _
'                          "       ACT_RETURN_DATE = '" & strDate & "'" & _
'                          " Where PER_GUMJIN_DATE = '" & Mid(lsID, 1, 8) & "'" & _
'                          "   And PER_GUM_NUM = " & lsID & "" & _
'                          "   And INTERFACECODE = '" & strEqpCd & "'"
'                Else
                    SQL = "Update onit_out..jun370_resulttb" & _
                          "   Set Result = '" & sResult & "'" & _
                          " Where orderorder = '" & VallsID & "'" & _
                          "   and map2seqno = '" & strEqpCd & "'"
'                End If
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
        
            End If
        Next iRow
            
    End With
           
    cn_Ser.CommitTrans
    SaveTransDataR = 1
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim strSex  As String
    Dim strAge  As String
    
    Dim strColPtID
    
    Dim ValBarcode As String
    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '2 ���� ���ڵ� ��ȣ
    ValBarcode = Val(sBarcode)
    
    If sBarcode = "" Then
        Exit Function
    End If
    

    SQL = ""
    SQL = SQL & " SELECT Distinct a.EnterDate, c.sujinname, a.ChartNo, a.SUJINPART "
    SQL = SQL & "   FROM " & gDB_Parm.DB & "..WaitPrsnp a, " & gDB_Parm.DB & "..jun370_resulttb b, " & gDB_Parm.DB & "..pewprsnp c, " & gDB_Parm.DB & "..BAGMAP2PREF d "
    SQL = SQL & "  WHERE a.WaitSeqNo = '" & sBarcode & "' "
    SQL = SQL & "    AND d.labno in (2) "
    SQL = SQL & "    AND b.map2seqno = d.map2seqno "
    SQL = SQL & "    AND a.chartno = c.chartno "
    SQL = SQL & "    AND a.jundal = '370'"
    SQL = SQL & "    AND a.WaitSeqNo = b.WaitSeqNo"
    SQL = SQL & "    AND b.Result = '' OR b.Result IS NULL"
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colOrdDate    '1  ó������
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID        '6  �˻��ڹ�ȣ(=���Ϲ�ȣ)
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName      '7  �˻��ڸ�
        
        If Trim(gReadBuf(3)) = "62" Then
            SetText frmInterface.vasID, "����", asRow, colSex
        Else
            SetText frmInterface.vasID, "����", asRow, colSex
        End If
        GetSampleInfoW = 1
    Else
        GetSampleInfoW = -1
    End If

End Function


'-- ������ ���� ��������
Function GetSampleInfoB(ByVal pBarNo As String) As Integer
    Dim intRow      As Long
    
    If pBarNo = "" Then
        Exit Function
    End If
    
    '���ڵ��ȣ�� ȯ������ �ҷ�����
    SQL = ""
    SQL = SQL & " SELECT Distinct a.CHARTNO, c.SUJINNAME, a.SUJINPART, a.ENTERDATE, b.WAITSEQNO " & vbLf
    SQL = SQL & "   FROM " & gDB_Parm.DB & "..WAITPRSNP a, " & gDB_Parm.DB & "..JUN370_RESULTTB b, " & gDB_Parm.DB & "..PEWPRSNP c, " & gDB_Parm.DB & "..BAGMAP2PREF d " & vbLf
    SQL = SQL & "  WHERE a.JUNDAL = '370' " & vbLf
    SQL = SQL & "    AND a.WAITSEQNO = b.WAITSEQNO " & vbLf
    SQL = SQL & "    AND a.CHARTNO = c.CHARTNO " & vbLf
    SQL = SQL & "    AND d.LABNO in (2) " & vbLf
    SQL = SQL & "    AND b.SPCM_NO = '" & pBarNo & "'"
    SQL = SQL & "    AND b.MAP2SEQNO IN (" & gAllExam & ") " & vbLf
    SQL = SQL & "    AND b.MAP2SEQNO = d.MAP2SEQNO " & vbLf
    SQL = SQL & "    AND (b.RESULT = '' OR b.RESULT IS NULL)" & vbLf
    
    'SetRawData "[SQL]" & SQL
    
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
        intRow = frmInterface.vasID.MaxRows
        
        SetText frmInterface.vasID, "1", intRow, colCheckBox
        SetText frmInterface.vasID, CStr(intRow), intRow, colSeqNo
        SetText frmInterface.vasID, Trim(gReadBuf(3)), intRow, colOrdDate
        SetText frmInterface.vasID, Trim(gReadBuf(4)), intRow, colBarcode
        
        SetText frmInterface.vasID, Trim(gReadBuf(0)), intRow, colPID    '5
        SetText frmInterface.vasID, Trim(gReadBuf(1)), intRow, colPName  '6
            
        If Trim(gReadBuf(2)) = "62" Then
            SetText frmInterface.vasID, "����", intRow, colSex
        Else
            SetText frmInterface.vasID, "����", intRow, colSex
        End If
        
        GetSampleInfoB = 1
        
        Call vasActiveCell(frmInterface.vasID, intRow, colOrdDate)

        frmInterface.txtBarcode.Text = ""
        frmInterface.txtBarcode.SetFocus
    Else
        GetSampleInfoB = -1
    End If

End Function

Function GetSampleInfoR(ByVal asRow As Long) As Integer
    Dim sBarcode As String
    Dim sSpecNo As String

    GetSampleInfoR = -1
    
    '-- ȯ������ ��������
    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBarcode))   '���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- ���ڵ��ȣ�� ȯ������ �ҷ�����
          SQL = "SELECT " & gDBCOLUMN_Parm.PID & "," & gDBCOLUMN_Parm.PNAME & "," & gDBCOLUMN_Parm.PSEX & "," & gDBCOLUMN_Parm.PAGE & vbCrLf
    SQL = SQL & "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
    SQL = SQL & " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
    If gDBCOLUMN_Parm.STATUS <> "" Then
        SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
    End If
    If gDBCOLUMN_Parm.RESULT <> "" Then
        SQL = SQL + "   AND (" & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL)"
    End If
    
    Res = GetDBSelectColumn(gServer, SQL)
    
    If Res = 1 Then
        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName
        '-- ������ ������� �ֹι�ȣ�� ã��
        'strSex = IIf(Mid(Trim(gReadBuf(4)), 7, 1) = "1", "M", "F")
        'SetText frmInterface.vasID, strSex, colSex    '7  ����
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '7  ����
        '-- ���̰� ������� �ֹι�ȣ�� ã��
        'strAge = Format(Now, "yyyy") - Mid(Trim(gReadBuf(3)), 1, 4)
        'SetText frmInterface.vasID, strAge, asRow, colAge
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colSex    '8  ����
        
        GetSampleInfoR = 1
    Else
    
        GetSampleInfoR = -1
    End If
    
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
    
    SQL = " Select examcode From EQPMASTER " & vbCrLf & _
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
    SQL = SQL & "  From EQPMASTER "
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

'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Function GetOrderExamCode(argEquipCode As String, argPID As String) As String
    Dim i           As Integer
    Dim sExamCode   As String
    Dim strExamCode As String
    Dim sExamCd     As String
    Dim RS As ADODB.Recordset
    Dim ValargPID   As String
    
    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    'ValargPID = Val(argPID)

    SQL = ""
    SQL = SQL & " SELECT Distinct b.MAP2SEQNO " & vbLf
    SQL = SQL & "   FROM " & gDB_Parm.DB & "..WaitPrsnp a, " & gDB_Parm.DB & "..jun370_resulttb b, " & gDB_Parm.DB & "..pewprsnp c, " & gDB_Parm.DB & "..BAGMAP2PREF d " & vbLf
    SQL = SQL & "  WHERE b.WAITSEQNO = '" & argPID & "' "
    SQL = SQL & "    AND a.JUNDAL = '370' " & vbLf
    SQL = SQL & "    AND a.WAITSEQNO = b.WAITSEQNO " & vbLf
    SQL = SQL & "    AND a.CHARTNO = c.CHARTNO " & vbLf
    SQL = SQL & "    AND d.LABNO in (2) " & vbLf
    SQL = SQL & "    AND b.MAP2SEQNO IN (" & gAllExam & ") " & vbLf
    SQL = SQL & "    AND b.MAP2SEQNO = d.MAP2SEQNO " & vbLf
    SQL = SQL & "    AND b.RESULT = '' OR b.RESULT IS NULL" & vbLf
    SQL = SQL & "  ORDER BY b.MAP2SEQNO "
    
    Set RS = cn_Ser.Execute(SQL)
    Do Until RS.EOF
        GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields(0)) & "',"
        RS.MoveNext
    Loop
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
    End If

    
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
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim sSpecNo     As String
Dim iRow        As Long
Dim SpecNo      As String
    
    GetGetEquipExamCode_E411 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '-- �ڰ�ü�� 11�ڸ��� ��ȸ�ϱ����Ͽ� ������ �ڸ��� ���ش�.
    argPID = Mid(argPID, 1, 10)
    
    If Mid(argPID, 1, 2) = "99" Then
        'strExamCode = Proc_Order_LX_QC(argPID)
        
        'iRow = frmInterface.vasID.DataRowCnt
        iRow = intRow
        
        SpecNo = Trim(GetText(frmInterface.vasID, iRow, colSpecNo))
        
        SQL = "SELECT QC_EXMN_CD "
        SQL = SQL & vbCrLf & " FROM SPSLMQMST "
        SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(SpecNo, 3, 3) & "' "     '//// ��� ��ȣ
        SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(SpecNo, 6, 3) & "' "     '//// �˻�� ��ȣ
        SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(SpecNo, 9, 1) & "' "      '//// ���� ��ȣ
        SQL = SQL & vbCrLf & "  AND QC_EXMN_CD IN (" & gAllExam & ") "
        SQL = SQL & vbCrLf & "  AND USE_STR_DT <= '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' "
        SQL = SQL & vbCrLf & "  AND USE_END_DT >= '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "' "
        Res = GetDBSelectRow(gServer, SQL)
        strExamCode = ""
        
        For i = 0 To UBound(gReadBuf)
            If gReadBuf(i) <> "" Then
                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
            Else
                Exit For
            End If
        Next
        
    Else
        '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
        SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
        Res = GetDBSelectColumn(gServer, SQL)
        sSpecNo = Trim(gReadBuf(0))
        
        '-- �˻��ڵ� ��������
        SQL = " Select EXMN_CD From SPSLHRRST " & vbCr & _
              " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
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
    End If
    
    If strExamCode = "" Then
'        MsgBox "������ ȯ��"
        GetGetEquipExamCode_E411 = ""
        Exit Function
    End If
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'GetEquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    
    '-- ������ �˻��ڵ��� ä�� ã��
          SQL = "Select distinct equipcode "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
    
        If gReadBuf(i) <> "" Then
            'gReadBuf(i) = Mid(gReadBuf(i), 1, Len(gReadBuf(i)) - 1)
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next
    
    GetGetEquipExamCode_E411 = Mid(strExamCode, 2)
    
End Function



Function GetGetEquipExamCode_Architect(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_Architect = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 ���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '-- �˻��ڵ� ��������
'    SQL = ""
'    SQL = SQL + "SELECT " & gDBCOLUMN_Parm.TESTCD & vbCrLf
'    SQL = SQL + "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
'    SQL = SQL + " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
'    SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
'    SQL = SQL + "   AND " & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL"
'
'    Res = GetDBSelectRow(gServer, SQL)
'    strExamCode = ""
'
'    For i = 0 To UBound(gReadBuf)
'        If gReadBuf(i) <> "" Then
'            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
'        Else
'            Exit For
'        End If
'    Next
'
'    If strExamCode = "" Then
'        '-- ������ȯ���̰ų� �ش���� �˻��� ����
'        GetGetEquipExamCode_Architect = ""
'        Exit Function
'    End If
'
'    '-- ������ "," �ڸ���
'    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = "          "
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    '-- �ش� ��� �°� ����ä�� �����ϱ� [ASTM Format >> Architect]
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next
    
    '-- ù�ڸ� "\" �ڸ���
    GetGetEquipExamCode_Architect = strExamCode
    
End Function

'-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ���
Function GetGetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_AU480 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 ���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'AU480�� ��� ��񿡼� dilution ���� ���� '0'�߰�
            strExamCode = strExamCode & "0" & Trim(gReadBuf(i)) & "0"
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_AU480 = strExamCode
    
End Function


'-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ���
Function GetGetEquipExamCode_CentaurCP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_CentaurCP = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBarcode))   '2 ���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""

    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_CentaurCP = strExamCode
    
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
    SQL = SQL & "  From EQPMASTER "
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


