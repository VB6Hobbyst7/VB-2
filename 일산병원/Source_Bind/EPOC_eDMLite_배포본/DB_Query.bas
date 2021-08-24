Attribute VB_Name = "DB_Query"
Option Explicit



'��� ����� G/M/B �ڵ� ��
Public gState_G     As String       '//// �׷��ڵ�
Public gState_M     As String       '//// ��Ƽ�ڵ�
Public gState_B     As String       '//// ���͸��ڵ�


Function Insert_Data_QC(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sCnt            As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim ExamCnt         As String
    Dim ExamCode_Spec   As String
    Dim lsQC_Date       As String
    
    Dim insCnt          As String
    Dim strQCcode As String
    Dim varQcCode As Variant
        
    With frmInterface
        Insert_Data_QC = -1
        ExamCode_Spec = ""
        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        
        lsQC_Date = Format(GetDateFull, "yyyymmdd")

        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, RESDATE, EXAMDATE" & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' " & vbCrLf & _
              " And sendflag < '2' "
        
        res = db_select_Vas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            If ExamCode_Spec <> "" Then
                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            Else
                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            End If
        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

    
        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        
        SQL = "SELECT EXMN_CD FROM SPSLMQCTM "
        If lsID = "9904800111" Then
            SQL = SQL & vbCrLf & " WHERE eqpm_cd = '048'"   'ABL800 BASIC QC
        Else
            SQL = SQL & vbCrLf & " WHERE eqpm_cd = '036'"   'ABL800 BASIC QC
        End If
        res = db_select_Row(gServer, SQL)
        
        strQCcode = ""
        
        For k = 0 To UBound(gReadBuf)
            If gReadBuf(k) <> "" Then
                strQCcode = strQCcode & Trim(gReadBuf(k)) & ","
            Else
                Exit For
            End If
        Next
        
        varQcCode = Split(strQCcode, ",")

        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If Mid(sResult1, 1, 3) = "-99" Then: sResult1 = " "
            
            If sResult1 <> "" Then
                SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                SQL = SQL & vbCrLf & "  AND RSLT_VALU IS NULL "
                SQL = SQL & vbCrLf & "GROUP BY RSLT_SQNO "
                res = db_select_Col(gServer, SQL)
                sCnt = gReadBuf(0)
                
                If IsNumeric(sCnt) = True Then
                    SQL = "UPDATE SPSLHQRST "
                    SQL = SQL & vbCrLf & "  SET RSLT_VALU = '" & sResult1 & "', "                        '���(�����)
                    SQL = SQL & vbCrLf & "      RSLT_DT = sysdate, "                                     '���(�������)"
                    SQL = SQL & vbCrLf & "      RSLT_RPTR_ID = '" & gUserID & "', "                                                           'Delta üũ"
                    SQL = SQL & vbCrLf & "      AMEN_ID = '" & gUserID & "', "                                                           'Panic üũ"
                    SQL = SQL & vbCrLf & "      UPDT_DT = sysdate "                                     '����Է���"
                    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                    SQL = SQL & vbCrLf & "  AND RSLT_SQNO = '" & sCnt & "' "
                    SQL = SQL & vbCrLf & "  AND RSLT_VALU IS NULL "
                    res = SendQuery(gServer, SQL)
                    If res < 0 Then
                        SaveQuery SQL
                       ' db_RollBack gServer
                       cn_Ser.RollbackTrans
                        Exit Function
                    End If
                
                Else
                    If insCnt = "" Then
                        SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                        SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                        SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                        SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                        'SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                        SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                        res = db_select_Col(gServer, SQL)
                    
                        If gReadBuf(0) = "" Then
                            insCnt = "1"
                        Else
                            insCnt = CLng(gReadBuf(0)) + 1
                        End If
                    End If
                    
                    For k = 0 To UBound(varQcCode) - 1
                        If Trim(GetText(.vasTemp, iRow, 2)) <> "" And Trim(GetText(.vasTemp, iRow, 2)) = varQcCode(k) Then
                            SQL = ""
                            SQL = SQL & vbCrLf & "INSERT INTO SPSLHQRST(EXMN_DY   ,EQPM_CD ,SBSN_CD ,LVL_CD  "
                            SQL = SQL & vbCrLf & "                     ,RSLT_SQNO ,EXMN_CD ,RSLT_DT ,RSLT_RPTR_ID "
                            SQL = SQL & vbCrLf & "                     ,RSLT_VALU ,SPCM_NO ,DEL_YN "
                            SQL = SQL & vbCrLf & "                     ,REGI_ID   ,RGST_DT ,AMEN_ID ,UPDT_DT) "
                            SQL = SQL & vbCrLf & "               VALUES('" & Trim(GetText(.vasTemp, iRow, 9)) & "', '" & Mid(lsID, 3, 3) & "', '" & Mid(lsID, 6, 3) & "', '" & Mid(lsID, 9, 1) & "', "
                            SQL = SQL & vbCrLf & "                      " & insCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gUserID & "', "
                            'SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gUserID & "', "
                            SQL = SQL & vbCrLf & "                      '" & sResult1 & "', '" & lsID & "', 'N', "
                            SQL = SQL & vbCrLf & "                      '" & gUserID & "', sysdate, '" & gUserID & "', sysdate ) "
                            res = SendQuery(gServer, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                cn_Ser.RollbackTrans
                                Exit Function
                            End If
                            Exit For
                        End If
                    Next
                        
                End If
                
            End If
            
        Next iRow
        
        cn_Ser.CommitTrans
        Insert_Data_QC = 1
    End With
    
End Function

'-- �ش� ȯ�� �˻��� H/L, Delta, Panic �����ϱ�
Function GetDecision(ByVal argSpcRow As Integer, ByVal strBarno As String, ByVal strExamCode As String, ByVal strResult As String) As String
    Dim rs_Delta        As ADODB.Recordset
    Dim rs_DPRef        As ADODB.Recordset
    Dim strBefoRslt     As String
    Dim strDestRslt     As String
    Dim strHLVal        As String
    Dim strDelta        As String
    Dim strPanic        As String
    Dim strSex          As String
    Dim strHVal         As String
    Dim strLVal         As String
                
             
    '-- ȯ���� ����
    strSex = Trim(GetText(frmInterface.vasID, argSpcRow, colSex))
    
    '##### ���ε� ���� - 11 ##############################################
''    '-- �ش� ȯ���� ����ġ,��Ÿ,�д� ã�ƿ���
''    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW                 "
''    SQL = SQL & vbCrLf & " FROM SPSLMFBIF                                                                                                                      "
''    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= SYSDATE                                                                                                         "
''    SQL = SQL & vbCrLf & "   AND USE_END_DY >= SYSDATE                                                                                                         "
''    SQL = SQL & vbCrLf & "   and EXMN_CD = '" & Trim(strExamCode) & "' "
''    Set rs_DPRef = cn_Ser.Execute(SQL)
    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW "
    SQL = SQL & vbCrLf & " FROM SPSLMFBIF               "
    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= ?  "
    SQL = SQL & vbCrLf & "   AND USE_END_DY >= ?  "
    SQL = SQL & vbCrLf & "   and EXMN_CD = ? "
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("USE_STR_DY", adDBDate, , , gsDBDateTime)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("USE_END_DY", adDBDate, , , gsDBDateTime)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, Trim(strExamCode))
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    If AdoRs_ORACLE.BOF = False Then
        '���ڵ�� ����
        Set rs_DPRef = AdoRs_ORACLE
        Do Until rs_DPRef.EOF
            '-- ������ ������� ��
            '-- ������� ��ġ�� ��쿡�� ���Ѵ�.
            If IsNumeric(strResult) Then
                strHLVal = ""
                If strSex = "M" Then
                    If IsNumeric(rs_DPRef.Fields("MALE_HIGH")) Then
                        If CDbl(strResult) > CDbl(rs_DPRef.Fields("MALE_HIGH")) Then
                            strHLVal = "H"
                        Else
                            strHLVal = " "
                        End If
                    Else
                        strHLVal = ""
                    End If
                    
                    If IsNumeric(rs_DPRef.Fields("MALE_LOW")) Then
                        If Trim(strHLVal) = "" Then
                            If CDbl(strResult) < CDbl(rs_DPRef.Fields("MALE_LOW")) Then
                                strHLVal = "L"
                            Else
                                strHLVal = " "
                            End If
                        End If
                    Else
                        strHLVal = ""
                    End If
                
                Else
                    If IsNumeric(rs_DPRef.Fields("FEML_HIGH")) Then
                        If CDbl(strResult) > CDbl(rs_DPRef.Fields("FEML_HIGH")) Then
                            strHLVal = "H"
                        Else
                            strHLVal = " "
                        End If
                    Else
                        strHLVal = ""
                    End If
                    If IsNumeric(rs_DPRef.Fields("FEML_LOW")) Then
                        If Trim(strHLVal) = "" Then
                            If (CDbl(strResult) < CDbl(rs_DPRef.Fields("FEML_LOW"))) Then
                                strHLVal = "L"
                            Else
                                strHLVal = " "
                            End If
                        End If
                    Else
                        strHLVal = ""
                    End If
                End If
            Else
                strHLVal = ""
            End If
            
            '-- Panic ����
            '-- ������� ��ġ�� ��쿡�� ���Ѵ�.
            If IsNumeric(strResult) Then
                strPanic = ""
                Select Case Trim(rs_DPRef.Fields("PANC_DVSN"))
                    Case 0:     '0 ������
                            strPanic = ""
                    Case 1:     '1 ���Ѹ�
                            If IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
                                If CDbl(strResult) > rs_DPRef.Fields("PANC_HIGH") Then
                                    strPanic = "P"
                                Else
                                    strPanic = " "
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case 2:     '2 ���Ѹ�
                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) Then
                                If CDbl(strResult) < rs_DPRef.Fields("PANC_LOW") Then
                                    strPanic = "P"
                                Else
                                    strPanic = " "
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case 3:     '3 ��� ���
                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) And IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
                                If (CDbl(strResult) < rs_DPRef.Fields("PANC_LOW") Or _
                                    CDbl(strResult) > rs_DPRef.Fields("PANC_HIGH")) Then
                                    strPanic = "P"
                                Else
                                    strPanic = " "
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case Else:
                            strPanic = ""
                End Select
            Else
                strPanic = ""
            End If
            
    
        
            '** ������� ��ȸ ����
            '-- ��Ÿ���� ����ϱ� ���� ������� ��ȸ (�Ѵ��̳� ������� �ֱٰ��� ��ȸ�Ѵ�.)
            SQL = ""
            SQL = SQL & vbCrLf & "SELECT B.SPCM_NO           BEFO_BCNO                                                               "
            SQL = SQL & vbCrLf & "     , B.EXMN_CD           BEFO_EXMN_CD                                                            "
            SQL = SQL & vbCrLf & "     , B.REAL_RSLT         BEFO_REAL_RSLT                                                          "
            SQL = SQL & vbCrLf & "     , B.VIEW_RSLT         BEFO_VIEW_RSLT                                                          "
            SQL = SQL & vbCrLf & "     , B.LAST_RPTG_DT     BEFO_FINL_DT                                                             "
            SQL = SQL & vbCrLf & "     , (SYSDATE - B.LAST_RPTG_DT)  DELTA_TERM_DT                                                   "  '���ú����� ������� �Ⱓ�� ���Ѵ�.
            SQL = SQL & vbCrLf & "     , B.PID               PID                                                                     "
            SQL = SQL & vbCrLf & "  FROM (SELECT MAX(B.LAST_RPTG_DT) LAST_RPTG_DT                                                    "
            SQL = SQL & vbCrLf & "             , B.EXMN_CD                                                                           "
            SQL = SQL & vbCrLf & "             , B.PID                                                                               "
            SQL = SQL & vbCrLf & "          FROM SPSLHRRST A, SPSLHRRST B                                                            "
            SQL = SQL & vbCrLf & "         WHERE A.SPCM_NO   <> B.SPCM_NO                                                            "
            SQL = SQL & vbCrLf & "           AND A.PID        = B.PID                                                                "
            SQL = SQL & vbCrLf & "           AND A.EXMN_CD    = B.EXMN_CD                                                            "
            SQL = SQL & vbCrLf & "           AND A.RCPN_DT   >= B.RCPN_DT                                                            "
            SQL = SQL & vbCrLf & "           AND B.LAST_RPTG_DT IS NOT NULL                                                          "
            'SQL = SQL & vbCrLf & "           AND A.RSLT_STAT < '3'                                                                   "
            SQL = SQL & vbCrLf & "           AND A.SPCM_NO = FN_LABCVTBCNO('" & strBarno & "')                                       "
            SQL = SQL & vbCrLf & "         GROUP BY B.PID, B.EXMN_CD ) A, SPSLHRRST B                                                "
            SQL = SQL & vbCrLf & " WHERE A.PID = B.PID                                                                               "
            SQL = SQL & vbCrLf & "   AND A.LAST_RPTG_DT = B.LAST_RPTG_DT                                                             "
            SQL = SQL & vbCrLf & "   AND A.EXMN_CD = B.EXMN_CD                                                                       "
            SQL = SQL & vbCrLf & "   AND A.EXMN_CD = '" & Trim(strExamCode) & "' "         '�˻��ڵ�"
            SQL = SQL & vbCrLf & "   AND B.LAST_RPTG_DT BETWEEN (SYSDATE-30) AND SYSDATE                "           '-- 30�� �̳�
            Set rs_Delta = cn_Ser.Execute(SQL)
            Do Until rs_Delta.EOF
                strBefoRslt = rs_Delta.Fields("BEFO_REAL_RSLT")             '�������
                strDestRslt = Trim(strResult)  '������
                If IsNumeric(strBefoRslt) = False Then '///////////////////// ��������� ���ڰ� ��������
                    Do
                        If Trim(strBefoRslt) = "" Then Exit Do
                        strBefoRslt = Mid(strBefoRslt, 2)
                        If IsNumeric(Mid(strBefoRslt, 1, 1)) = True Then
                            If InStr(1, strBefoRslt, ")") > 0 Then: strBefoRslt = Mid(strBefoRslt, 1, InStr(1, strBefoRslt, ")") - 1)
                            Exit Do
                        End If
                    Loop
                End If
                
                '-- Delta ����  (�Ʒ� ������ �´��� ���� �ʿ���...��)
                '-- ������� ��ġ�� ��쿡�� ���Ѵ�.
                If IsNumeric(strDestRslt) And IsNumeric(strBefoRslt) = True Then
                    strDelta = ""
                    Select Case Trim(rs_DPRef.Fields("DELT_DVSN"))
                        Case 0:     '0 ������
                                strDelta = ""
                        Case 1:     '1 ��ȭ�� = ������ - �������
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '��ȭ��
                        Case 2:     '2 ��ȭ���� = ��ȭ�� / ������� * 100
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '��ȭ��
                                strDelta = (CDbl(strDelta) / CDbl(strBefoRslt)) * 100               '��ȭ����
                        Case 3:     '3 �Ⱓ�� ��ȭ���� = ��ȭ���� / �Ⱓ
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '��ȭ��
                                strDelta = (CDbl(strDelta) / CDbl(strBefoRslt)) * 100               '��ȭ����
                                strDelta = strDelta / CDbl(rs_Delta.Fields("DELTA_TERM_DT"))        '�Ⱓ�� ��ȭ����
                        Case 4:     '4 �Ⱓ�� ��ȭ�� = ��ȭ�� / �Ⱓ
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '��ȭ��
                                strDelta = CDbl(strDelta) / CDbl(rs_Delta.Fields("DELTA_TERM_DT"))  '�Ⱓ�� ��ȭ��
                        Case 5:     '5 ���뺯ȭ���� = ��ȭ�� / �������
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '��ȭ��
                                strDelta = CDbl(strDelta) / CDbl(strBefoRslt)                       '���뺯ȭ����
                        Case Else:
                                strDelta = ""
                    End Select
                Else
                    strDelta = ""
                End If
                '-- Delta ����
                If IsNumeric(rs_DPRef.Fields("DELT_HIGH")) And IsNumeric(rs_DPRef.Fields("DELT_LOW")) Then
                    If (CDbl(strDestRslt) > rs_DPRef.Fields("DELT_HIGH") Or CDbl(strDestRslt) < rs_DPRef.Fields("DELT_LOW")) Then
                        strDelta = "D"
                    Else
                        strDelta = " "
                    End If
                Else
                    strPanic = ""
                End If
    
                rs_Delta.MoveNext
            Loop
            
            rs_DPRef.MoveNext
        Loop
    End If
    
    GetDecision = strHLVal & "|" & strDelta & "|" & strPanic
    
    Set rs_DPRef = Nothing
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    
    '##### ���ε� ���� - 11 ##############################################
        
    GetDecision = strHLVal & "|" & strDelta & "|" & strPanic
    
End Function



Function RsltState_Check(asSpecNo As String, asExamCode As String) As String '/// ��� ���� : (�׷��ڵ�/��Ƽ�ڵ�) : ���°� �߰����� �����϶�
    Dim PRSC_CD_G       As String
    Dim EXMN_CD         As String
    Dim PRSC_CD_M       As String
    Dim PRSC_CD_B       As String
    
    Dim strExam         As Variant
    Dim varExam         As Variant
    Dim i               As Integer
    Dim strWhereSQL     As String
    
    RsltState_Check = ""
    PRSC_CD_G = " "
    PRSC_CD_M = " "
    PRSC_CD_B = " "
    
    '##### ���ε� ���� - 37 ##############################################
''    SQL = ""
''    SQL = SQL & vbCrLf & "SELECT DISTINCT "
''    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
''    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
''    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
''    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
''    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
''    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
''    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = '" & asExamCode & "' "
''    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE ('%G%') "
''    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
''    res = db_select_Col(gServer, SQL)
''    If gReadBuf(0) <> "" Then: PRSC_CD_G = gReadBuf(0)
''    gReadBuf(0) = ""

    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT R1.PRSC_CD, NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1, SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = ? "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = ? "
    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE (?) "
    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, asSpecNo)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, asExamCode)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PRSC_CD", adVarChar, , 12, "%G%")   'ó���ڵ�
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    If AdoRs_ORACLE.BOF = False Then
        PRSC_CD_G = AdoRs_ORACLE.Fields(0) & ""
    End If
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    '##### ���ε� ���� - 37 ##############################################
    
    '##### ���ε� ���� - 28 ##############################################
''    SQL = ""
''    SQL = SQL & vbCrLf & "SELECT DISTINCT "
''    SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
''    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
''    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
''    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
''    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
''    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
''    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
''    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('M') "
''    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
''    res = db_select_Col(gServer, SQL)
''    If gReadBuf(0) <> "" Then: PRSC_CD_M = gReadBuf(0)
''    gReadBuf(0) = ""
    strExam = Replace(gAllExam, "'", "")
    varExam = Split(strExam, ",")
    strWhereSQL = ""
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    SQL = SQL & vbCrLf & "       R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = ? "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN (?) "
    For i = 0 To UBound(varExam)
        strWhereSQL = strWhereSQL & "R1.EXMN_CD = ? OR "
    Next
    If strWhereSQL <> "" Then
        SQL = SQL & vbCrLf & "AND (" & Mid(strWhereSQL, 1, Len(strWhereSQL) - 3) & ")"
    End If
    
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, asSpecNo)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("CD_DVSN", adVarChar, , 5, "M")
    For i = 0 To UBound(varExam)
        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, varExam(i))
    Next
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    If AdoRs_ORACLE.BOF = False Then
        PRSC_CD_M = AdoRs_ORACLE.Fields(0) & ""
    End If
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    '##### ���ε� ���� - 28 ##############################################

    '##### ���ε� ���� - 18 ##############################################
''    SQL = ""
''    SQL = SQL & vbCrLf & "SELECT DISTINCT "
''    SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
''    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
''    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
''    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
''    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
''    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
''    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
''    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('B') "
''    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
''    res = db_select_Col(gServer, SQL)
''    If gReadBuf(0) <> "" Then: PRSC_CD_B = gReadBuf(0)
''    gReadBuf(0) = ""

    Erase varExam
    strExam = Replace(gAllExam, "'", "")
    varExam = Split(strExam, ",")
    strWhereSQL = ""
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT R1.EXMN_CD,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1, SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO = ? "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN (?) "
    For i = 0 To UBound(varExam)
        strWhereSQL = strWhereSQL & "R1.EXMN_CD = ? OR "
    Next
    If strWhereSQL <> "" Then
        SQL = SQL & vbCrLf & "AND (" & Mid(strWhereSQL, 1, Len(strWhereSQL) - 3) & ")"
    End If
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, asSpecNo)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("CD_DVSN", adVarChar, , 5, "B")
    For i = 0 To UBound(varExam)
        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, varExam(i))
    Next
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    If AdoRs_ORACLE.BOF = False Then
        PRSC_CD_B = AdoRs_ORACLE.Fields(0) & ""
    End If
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    '##### ���ε� ���� - 18 ##############################################
        
    RsltState_Check = PRSC_CD_G & "/" & PRSC_CD_M & "/" & PRSC_CD_B
    
End Function


Function Make_Remark_all(asExamCode As String, asSex As String, asResult As String)
'///////////// �ڸ�Ʈ ���� (��ü��ü)
    Dim i As Integer
    
    Dim Comment_Gubun As String
    Dim Comment_MFGubun As String
    Dim Comment_Code As String      '///////// �Ǻ�����
    Dim Comment_CodeH As String
    Dim Comment_CodeL As String

    Dim Comment_RefMH As String
    Dim Comment_RefML As String
    Dim Comment_RefFH As String
    Dim Comment_RefFL As String

    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT cmtdest, cmtflag, CMTCODE, cmtcodeSub, cmhigh, cmlow, cFhigh, cFlow "
    SQL = SQL & vbCrLf & "  FROM EQUIPEXAM "
    SQL = SQL & vbCrLf & " WHERE EXAMCODE IN (" & asExamCode & ") "
    SQL = SQL & vbCrLf & "   AND CMTDEST = '1' "
    
    res = db_select_Col(gLocal, SQL)
    
    If res < 1 Then: Exit Function
    
    Comment_Gubun = gReadBuf(0)
    Comment_MFGubun = gReadBuf(1)
    Comment_CodeH = gReadBuf(2)
    Comment_CodeL = gReadBuf(3)
    Comment_RefMH = gReadBuf(4)
    Comment_RefML = gReadBuf(5)
    Comment_RefFH = gReadBuf(6)
    Comment_RefFL = gReadBuf(7)

    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""
    gReadBuf(4) = ""
    gReadBuf(5) = ""
    gReadBuf(6) = ""
    gReadBuf(7) = ""
        
        
    '///// 0:����, 1:��/��, 2:������
    If Comment_MFGubun = "0" Then
        
        If asResult >= Comment_RefMH Then
            Comment_Code = Comment_CodeH
        ElseIf asResult <= Comment_RefML Then
            Comment_Code = Comment_CodeL
        End If
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT CNTS "
        SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
        SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_Code & "' "
        SQL = SQL & vbCrLf & ""
        res = db_select_Col(gServer, SQL)
        
        
        
        
        If InStr(1, gComment_All, gReadBuf(0)) = 0 Then
            If gComment_All = "" Then
                gComment_All = gReadBuf(0)
            Else
                gComment_All = gComment_All & chrCR & gReadBuf(0)
            End If
        End If
    ElseIf Comment_MFGubun = "1" Then
        
        If asSex = "M" Then
            If asResult >= Comment_RefMH Then
                Comment_Code = Comment_CodeH
            ElseIf asResult <= Comment_RefML Then
                Comment_Code = Comment_CodeL
            End If
        ElseIf asSex = "F" Then
            If asResult >= Comment_RefFH Then
                Comment_Code = Comment_CodeH
            ElseIf asResult <= Comment_RefFL Then
                Comment_Code = Comment_CodeL
            End If
        End If

        SQL = ""
        SQL = SQL & vbCrLf & "SELECT CNTS "
        SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
        SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_Code & "' "
        SQL = SQL & vbCrLf & ""
        res = db_select_Col(gServer, SQL)
        
        If InStr(1, gComment_All, gReadBuf(0)) = 0 Then
            If gComment_All = "" Then
                gComment_All = gReadBuf(0)
            Else
                gComment_All = gComment_All & chrCR & gReadBuf(0)
            End If
        End If
        
    ElseIf Comment_MFGubun = "2" Then
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT CNTS "
        SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
        SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_CodeH & "' "
        SQL = SQL & vbCrLf & ""
        res = db_select_Col(gServer, SQL)
        
        If InStr(1, gComment_All, gReadBuf(0)) = 0 Then
            If gComment_All = "" Then
                gComment_All = gReadBuf(0)
            Else
                gComment_All = gComment_All & chrCR & gReadBuf(0)
            End If
        End If
        
    End If

    
End Function
'''
'''Function Insert_Data(ByVal argSpcRow As Integer) As Integer
'''    Dim iRow            As Integer
'''    Dim i               As Integer
'''    Dim j               As Integer
'''    Dim lsID            As String
'''    Dim lsSpecNo        As String
'''    Dim lsPid           As String
'''    Dim sResult         As String
'''    Dim sCnt            As String
'''    Dim sResult1        As String
'''    Dim sResult2        As String
'''    Dim ExamCnt         As String
'''    Dim ExamCode_Spec   As String
'''    Dim ExamCode_Remark     As String
'''
'''    Dim State_GM    As String       '//// �׷�/��Ƽ �ڵ�
'''    Dim State_cnt   As Integer      '//// �׷�/��Ƽ �ڵ� �� ����
'''    Dim State_G     As String       '//// �׷��ڵ�
'''    Dim State_M     As String       '//// ��Ƽ�ڵ�
'''    Dim State_B     As String       '//// ���͸��ڵ�
'''
'''    Dim Send_State      As String
'''    Dim SQL_LOCAL As String
'''
'''    Dim strWhereSQL     As String
'''    Dim sqlRet          As Integer
'''
'''    With frmInterface
'''        gComment_All = ""
'''        Insert_Data = -1
'''        ExamCode_Spec = ""
'''        ExamCode_Remark = ""
'''
'''        State_GM = ""
'''        State_cnt = 0
'''        State_G = ""
'''        State_M = ""
'''        lsID = ""
'''        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
'''        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
'''        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
'''
'''        'Local���� ȯ�ں��� ����� ��������
'''        ClearSpread .vasTemp
'''        ClearSpread .vasTemp1
'''        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
'''              " From pat_res " & vbCrLf & _
'''              " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
'''              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
'''              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
'''              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
'''
'''        Save_Raw_Data "[������ȸ]  " & SQL
'''
'''        res = db_select_Vas(gLocal, SQL, .vasTemp)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            Exit Function
'''        End If
'''
'''        For i = 1 To frmInterface.vasTemp.DataRowCnt    '/// ���� �˻��� �˻��ڵ��
'''            If ExamCode_Spec <> "" Then
'''                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
'''            Else
'''                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
'''            End If
'''        Next i
'''
'''        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
'''        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
'''
'''        sCnt = ""
'''        sResult1 = ""
'''        sResult2 = ""
'''
'''
'''
'''        '/-------------------------------����ũ ó�� ������ �������̽��� ����� �ڵ�� ��ü�� ��ȸ�ؼ� ����ũ ǥ�����ٰ��� ã��(�ʿ������ ����)
'''        SQL = "SELECT EXMN_CD "
'''        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
'''        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
'''        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
'''        res = db_select_Vas(gServer, SQL, .vasTemp1)
'''
'''
'''        For i = 1 To frmInterface.vasTemp1.DataRowCnt    '/// ���� �˻��� �˻��ڵ��
'''            If ExamCode_Remark <> "" Then
'''                ExamCode_Remark = ExamCode_Remark & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
'''            Else
'''                ExamCode_Remark = "'" & Trim(GetText(frmInterface.vasTemp, i, 1)) & "'"
'''            End If
'''        Next i
'''
'''        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
'''
'''        For i = 1 To frmInterface.vasTemp.DataRowCnt
'''            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
'''        Next i
'''        '/--------------------------------------------------------------------------------------------------------------
'''
'''        cn_Ser.BeginTrans
'''        '������ ����� �����ϱ�
'''        For iRow = 1 To .vasTemp.DataRowCnt
'''
'''            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
'''            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
'''
'''            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
'''                gComment_Code = ""
'''
'''
'''                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '�˻��ڵ�"
'''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
'''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
'''
'''                Save_Raw_Data "[SQL]  " & SQL
'''
'''                res = db_select_Col(gServer, SQL)
'''
'''                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
'''
'''                sCnt = CLng(gReadBuf(0)) + 1
'''
'''                '/----------------------------- �ڵ�����ũ ó�� (�ʿ������ ����)
''''                Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
'''                '/-----------------------------
'''
'''                '-- ������� ���ڰ��� ��츸 ��Ÿ/�д� ������ �Ѵ�.
'''                sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
'''                If IsNumeric(sResult) Then
'''                    Dim strDecision     As Variant
'''                    Dim strBarcode      As String
'''
'''                    strBarcode = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
'''                    'strDecision = GetDecision(argSpcRow, strBarcode, iRow)
'''                    strDecision = GetDecision(argSpcRow, strBarcode, Trim(GetText(frmInterface.vasTemp, iRow, 2)), sResult)
'''                    strDecision = Split(strDecision, "|")
'''                Else
'''                    strDecision = "||"
'''                    strDecision = Split(strDecision, "|")
'''                End If
'''
'''
'''                SQL = "UPDATE SPSLHRRST "   '-- ������̺�
'''                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "      '���(�����)
'''                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "      '���(�������)"
'''                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & strDecision(0) & "', "                                    'H/L üũ"
'''                SQL = SQL & vbCrLf & "       DLTA_YN = '" & strDecision(1) & "', "                                      'Delta üũ"
'''                SQL = SQL & vbCrLf & "       PANC_YN = '" & strDecision(2) & "', "                                      'Panic üũ"
'''                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
'''
'''
'''                Send_State = "1" '/  <---------- ��������� �ƴ϶� ���°� 1�θ� ��
'''
'''                '/----------------------------- ��� ���� �ֱ�
'''                If Send_State = "1" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
'''                ElseIf Send_State = "2" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                 '�߰�������"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
'''                ElseIf Send_State = "3" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                     '�߰�������"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'''                    SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '����������"
'''                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
'''                End If
'''
'''                '/----------------------------- ��� ���� �ֱ�
'''
'''                '/----------------------------- �ڵ�����ũ ó�� (�ʿ������ ����)
'''                If gComment_All <> "" Or gComment_Code <> "" Then
'''                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
'''                End If
'''                '/-----------------------------
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
'''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
'''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
'''
'''                Save_Raw_Data "[SQL]  " & SQL
'''
'''                res = SendQuery(gServer, SQL)
'''                If res < 0 Then
'''                    SaveQuery SQL
'''                   ' db_RollBack gServer
'''                   cn_Ser.RollbackTrans
'''                    Exit Function
'''                End If
'''
'''                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 2)))
'''
'''                State_cnt = InStr(1, State_GM, "/")
'''                State_G = Mid(State_GM, 1, State_cnt - 1)
'''                State_GM = Mid(State_GM, State_cnt + 1)
'''                State_cnt = InStr(1, State_GM, "/")
'''                State_M = Mid(State_GM, 1, State_cnt - 1)
'''                State_B = Mid(State_GM, State_cnt + 1)
'''
'''
'''                '/------------------------------------ ������̺� �׷��ڵ� ���� ������Ʈ
'''                If Trim(State_G) <> "" Then
'''                    SQL = "UPDATE SPSLHRRST "
'''
'''                        '/////////  D/P/H �� ������ : �˻����� ��������� �ִ´�
'''                        If Send_State = "1" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "2" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                 '�߰�������"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '����������"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "3" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '�߰�������"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '����������"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        End If
'''                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
'''                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '�˻��ڵ�"
'''                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
'''                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
'''
'''                    Save_Raw_Data "[SQL]  " & SQL
'''
'''                    res = SendQuery(gServer, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        cn_Ser.RollbackTrans
'''                        Exit Function
'''                    End If
'''                End If
'''                '/------------------------------------
'''
'''                '/------------------------------------ ������̺� ��Ƽ�ڵ� ���� ������Ʈ
'''                If Trim(State_M) <> "" Then
'''                    SQL = "UPDATE SPSLHRRST "
'''
'''                        '/////////  D/P/H �� ������ : �˻����� ��������� �ִ´�
'''                        If Send_State = "1" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "2" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                 '�߰�������"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '����������"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "3" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '�߰�������"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '����������"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        End If
'''                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
'''                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_M) & "' "                                        '�˻��ڵ�"
'''                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
'''                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
'''
'''                    Save_Raw_Data "[SQL]  " & SQL
'''
'''                    res = SendQuery(gServer, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        cn_Ser.RollbackTrans
'''                        Exit Function
'''                    End If
'''                End If
'''            '/------------------------------------
'''
'''            '/------------------------------------ ������̺� �׷��ڵ� ���� ������Ʈ
'''                If Trim(State_B) <> "" Then
'''                    SQL = "UPDATE SPSLHRRST "
'''
'''                        '/////////  D/P/H �� ������ : �˻����� ��������� �ִ´�
'''                        If Send_State = "1" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "2" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                 '�߰�������"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '����������"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "3" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '����Է���"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '�߰�������"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '����������"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '���������
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        End If
'''                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
'''                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_B) & "' "                                        '�˻��ڵ�"
'''                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
'''                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
'''
'''                    Save_Raw_Data "[SQL]  " & SQL
'''
'''                    res = SendQuery(gServer, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        cn_Ser.RollbackTrans
'''                        Exit Function
'''                    End If
'''                End If
'''            '/------------------------------------
'''
'''            '/------------------------------------ �������̺� STATE ������Ʈ
'''
'''                '##### ���ε� ���� - 7 ##############################################
'''''                '////////// ���� ���̺�
'''''                SQL = "UPDATE SPSLMJBDI "
'''''                If Send_State = "1" Then
'''''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
'''''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''                ElseIf Send_State = "2" Then
'''''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
'''''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'''''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''                ElseIf Send_State = "3" Then
'''''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
'''''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'''''                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
'''''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''                End If
'''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''''                SQL = SQL & vbCrLf & "   AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "','" & Trim(State_B) & "','" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "') "                    '�˻��ڵ�"
'''''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
'''''                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
'''''                res = SendQuery(gServer, SQL)
'''
'''                If Send_State = "1" Then
'''                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ? " & vbCrLf
'''                ElseIf Send_State = "2" Then
'''                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ? " & vbCrLf
'''                ElseIf Send_State = "3" Then
'''                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ?, LAST_RPTG_DT = ? " & vbCrLf
'''                End If
'''                SQL = SQL & " WHERE SPCM_NO = ? " & vbCrLf
'''                SQL = SQL & "   AND RSLT_STAT <> ? " & vbCrLf
'''                SQL = SQL & "   AND SPCM_STAT = ? " & vbCrLf
'''                strWhereSQL = ""
'''                If State_G <> "" Then
'''                    strWhereSQL = strWhereSQL & "   AND (EXMN_CD = ? "
'''                    If State_M <> "" Then
'''                        strWhereSQL = strWhereSQL & "   OR EXMN_CD = ? )"
'''                    Else
'''                        strWhereSQL = strWhereSQL & ")"
'''                    End If
'''                Else
'''                    If State_M <> "" Then
'''                        strWhereSQL = strWhereSQL & "   AND EXMN_CD = ? "
'''                    End If
'''                End If
'''                SQL = SQL & strWhereSQL
'''
'''                Set AdoCmd_ORACLE = New ADODB.Command
'''                Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
'''
'''                AdoCmd_ORACLE.CommandType = adCmdText
'''                AdoCmd_ORACLE.CommandText = SQL
'''
'''                '-- �ý��� ��¥ �������� �Լ� : gsDBDateTime
'''
'''                If Send_State = "1" Then
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, gUserID & "")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'''                    If State_G <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, Trim(State_G))
'''                    End If
'''                    If State_M <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, Trim(State_M))
'''                    End If
'''                ElseIf Send_State = "2" Then
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 15, "2")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, gUserID & "")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , 5, gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'''                    If State_G <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 5, Trim(State_G))
'''                    End If
'''                    If State_M <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 5, Trim(State_M))
'''                    End If
'''                ElseIf Send_State = "3" Then
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, gUserID & "")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'''                    If State_G <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 5, Trim(State_G))
'''                    End If
'''                    If State_M <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 5, Trim(State_M))
'''                    End If
'''                End If
'''
'''                AdoCmd_ORACLE.Execute sqlRet
'''                Set AdoCmd_ORACLE = Nothing
'''                '##### ���ε� ���� - 7 ##############################################
'''
'''                If sqlRet < 0 Then
'''                    SaveQuery SQL
'''                    cn_Ser.RollbackTrans
'''                    Exit Function
'''                End If
'''
'''            '/------------------------------------
'''            End If
'''        Next iRow
'''
'''        '/------------------------------------ ó�����̺� STATE ������Ʈ
'''        '///////// ó�����̺�
'''        '##### ���ε� ���� - 19 ##############################################
'''''        SQL = "UPDATE SPSLMJBBI "
'''''        If Send_State = "1" Then
'''''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
'''''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''        ElseIf Send_State = "2" Then
'''''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
'''''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''        ElseIf Send_State = "3" Then
'''''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
'''''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''        End If
'''''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
'''''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
'''''        res = SendQuery(gServer, SQL)
'''
'''        SQL = "UPDATE SPSLMJBBI "
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ? "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "
'''        SQL = SQL & vbCrLf & "   AND PID = ? "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> ? "
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = ? "
'''
'''        Set AdoCmd_ORACLE = New ADODB.Command
'''        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
'''
'''        AdoCmd_ORACLE.CommandType = adCmdText
'''        AdoCmd_ORACLE.CommandText = SQL
'''
'''        If Send_State = "1" Then
'''            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
'''        ElseIf Send_State = "2" Then
'''            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
'''        ElseIf Send_State = "2" Then
'''            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''        End If
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, gUserID & "")
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 8, Trim(lsPid))
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'''
'''        AdoCmd_ORACLE.Execute sqlRet
'''        Set AdoCmd_ORACLE = Nothing
'''        '##### ���ε� ���� - 19 ##############################################
'''
'''        If sqlRet < 0 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''        '/------------------------------------
'''        'db_Commit gServer
'''        cn_Ser.CommitTrans
'''        Insert_Data = 1
'''    End With
'''
'''End Function


'//////////////��� ���� �ٲ� (2011.10.11) - ȿ��
Function Insert_Data(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sCnt            As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim ExamCnt         As String
    Dim ExamCode_Spec   As String
    Dim ExamCode_Remark As String
    
    Dim State_GM        As String       '//// �׷�/��Ƽ �ڵ�
    Dim State_cnt       As Integer      '//// �׷�/��Ƽ �ڵ� �� ����
    Dim State_G         As String       '//// �׷��ڵ�
    Dim State_M         As String       '//// ��Ƽ�ڵ�
    Dim State_B         As String       '//// ���͸��ڵ�
    
    Dim Send_State      As String
    Dim SQL_LOCAL       As String
    Dim CODE_L8_cnt     As String
    
    Dim sqlRet          As Integer
    

    With frmInterface
        gComment_All = ""
        Insert_Data = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        CODE_L8_cnt = ""
        State_GM = ""
        State_cnt = 0
        State_G = ""
        State_M = ""
        gState_G = ""
        gState_M = ""
        gState_B = ""
        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))

        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt    '/// ���� �˻��� �˻��ڵ��
            If ExamCode_Spec <> "" Then
                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            Else
                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            End If
        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        
        
        SQL = "SELECT COUNT(A.EXMN_CD), B.BLPS_ID FROM SPSLHRRST A, SPSLMJBDI B "
        SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "                                             '��ü��ȣ"
        SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
        SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "                                              '�˻��ڵ�"
        'SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
        SQL = SQL & vbCrLf & " GROUP BY  B.BLPS_ID "
        res = db_select_Col(gServer, SQL)
        
        If Val(gReadBuf(0)) = "0" Then Exit Function
        .lblUser.Caption = gReadBuf(1)

        
        
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" And Trim(GetText(.vasTemp, iRow, 2)) <> "" And lsSpecNo <> "" Then
                gComment_Code = ""
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1

'                               SQL = "UPDATE SPSLHRRST "
'                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
'                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '���(�������)"
'                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL üũ"
'                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta üũ"
'                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic üũ"
'                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
'                SQL = SQL & vbCrLf & "       AMEN_ID = '" & .lblUser.Caption & "', "                                      '���������
'                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
'                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
'
'
'                Send_State = "3" '/  <---------- ��������� �ƴ϶� ���°� 3�θ� ��
'
'                '/----------------------------- ��� ���� �ֱ�
'                If Send_State = "1" Then
'
'                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & .lblUser.Caption & "', "                                 '����Է���"
'                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
'                ElseIf Send_State = "2" Then
'
'                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & .lblUser.Caption & "', "                                 '����Է���"
'                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & .lblUser.Caption & "', "                                  '�߰�������"
'                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
'                ElseIf Send_State = "3" Then
'
'                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & .lblUser.Caption & "', "                                 '����Է���"
'                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
'                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & .lblUser.Caption & "', "                                 '�߰�������"
'                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
'                    SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & .lblUser.Caption & "', "                                 '����������"
'                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
'                End If
'
'                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
'                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
'                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
'                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
'                res = SendQuery(gServer, SQL)
'                If res < 0 Then
'                    SaveQuery SQL
'                   ' db_RollBack gServer
'                   cn_Ser.RollbackTrans
'                    Exit Function
'                End If
                

                Send_State = "3"
                
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = ?, "                                          '���(�����)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = ?, "                                          '���(�������)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = ?, "                  'HL üũ"
                SQL = SQL & vbCrLf & "       PANC_YN = ?, "                    'Delta üũ"
                SQL = SQL & vbCrLf & "       DLTA_YN = ?, "                    'Panic üũ"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = ?, "                                        '����ڵ�
                SQL = SQL & vbCrLf & "       AMEN_ID = ?, "                                      '���������
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                SQL = SQL & vbCrLf & "       RSLT_NO = ?, "                                                '�����ȣ (��� �����ÿ� ����)
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = ?, "                                 '����Է���"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = ?, "                                     '�߰�������"
                SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
                SQL = SQL & vbCrLf & "       LAST_RPTR_ID = ?, "                                 '����������"
                SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
                SQL = SQL & vbCrLf & "       RSLT_STAT = ? "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = ? "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = ? "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < ? "                                                          '�������"
                
                Set AdoCmd_ORACLE = New ADODB.Command
                Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
                AdoCmd_ORACLE.CommandType = adCmdText
                AdoCmd_ORACLE.CommandText = SQL
                
                
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("REAL_RSLT", adVarChar, , 20, sResult1)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("VIEW_RSLT", adVarChar, , 20, sResult2)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("DTRM_DVSN", adVarChar, , 20, Trim(GetText(.vasTemp, iRow, 5)))
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PANC_YN", adVarChar, , 20, Trim(GetText(.vasTemp, iRow, 6)))
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("DLTA_YN", adVarChar, , 20, Trim(GetText(.vasTemp, iRow, 7)))
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 20, gEquipCode)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_NO", adVarChar, , 20, sCnt)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 20, Send_State)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 20, lsSpecNo)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 20, Trim(GetText(.vasTemp, iRow, 2)))
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 20, lsPid)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 20, "2")
                
                AdoCmd_ORACLE.Execute sqlRet
                Set AdoCmd_ORACLE = Nothing
                
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 2)))
                
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_GM = Mid(State_GM, State_cnt + 1)
                State_cnt = InStr(1, State_GM, "/")
                State_M = Mid(State_GM, 1, State_cnt - 1)
                State_B = Mid(State_GM, State_cnt + 1)
                
                '/�ڵ� �񱳽� �̿�
                If InStr(1, gState_G, State_G) = 0 Then
                    gState_G = gState_G & "," & State_G
                End If
                If InStr(1, gState_M, State_M) = 0 Then
                    gState_M = gState_M & "," & State_M
                End If
                If InStr(1, gState_B, State_B) = 0 Then
                    gState_B = gState_B & "," & State_B
                End If
                
                If Mid(gState_G, 1, 1) = "," Then gState_G = Mid(gState_G, 2)
                
                If Mid(gState_M, 1, 1) = "," Then gState_M = Mid(gState_M, 2)
                
                If Mid(gState_B, 1, 1) = "," Then gState_B = Mid(gState_B, 2)
                
                
                '/------------------------------------ ������̺� �׷��ڵ� ���� ������Ʈ
                If Trim(State_G) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "2" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                                  '�߰������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "3" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                                  '�߰������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = ? "                                 '����������"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = ? "                                                  '���������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        End If
                        
                        SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "                                             '��ü��ȣ"
                        SQL = SQL & vbCrLf & "   AND EXMN_CD = ? "                                        '�˻��ڵ�"
                        SQL = SQL & vbCrLf & "   AND PID = ? "                                                    'ȯ�ڹ�ȣ"
                        SQL = SQL & vbCrLf & "   AND RSLT_STAT < ? "                                                          '�������"

                        Set AdoCmd_ORACLE = New ADODB.Command
                        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
                        
                        AdoCmd_ORACLE.CommandType = adCmdText
                        AdoCmd_ORACLE.CommandText = SQL
                        
                        '-- �ý��� ��¥ �������� �Լ� : gsDBDateTime
                        
                        If Send_State = "1" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "2" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "3" Then

                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        End If
                            
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_NO", adVarChar, , 1, sCnt)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 15, lsPid)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                        
                        AdoCmd_ORACLE.Execute sqlRet
                        Set AdoCmd_ORACLE = Nothing
                End If
                '/------------------------------------
                
                '/------------------------------------ ������̺� ��Ƽ�ڵ� ���� ������Ʈ
                If Trim(State_M) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "2" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                                  '�߰������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "3" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                                  '�߰������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = ? "                                 '����������"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = ? "                                                  '���������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        End If
                        
                        SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "                                             '��ü��ȣ"
                        SQL = SQL & vbCrLf & "   AND EXMN_CD = ? "                                        '�˻��ڵ�"
                        SQL = SQL & vbCrLf & "   AND PID = ? "                                                    'ȯ�ڹ�ȣ"
                        SQL = SQL & vbCrLf & "   AND RSLT_STAT < ? "                                                          '�������"

                        Set AdoCmd_ORACLE = New ADODB.Command
                        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
                        
                        AdoCmd_ORACLE.CommandType = adCmdText
                        AdoCmd_ORACLE.CommandText = SQL
                        
                        '-- �ý��� ��¥ �������� �Լ� : gsDBDateTime
                        
                        If Send_State = "1" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "2" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "3" Then

                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        End If
                            
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_NO", adVarChar, , 1, sCnt)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 15, lsPid)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                            
                        AdoCmd_ORACLE.Execute sqlRet
                        Set AdoCmd_ORACLE = Nothing
                End If
            '/------------------------------------
            
            '/------------------------------------ ������̺� �׷��ڵ� ���� ������Ʈ
                If Trim(State_B) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                 '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                    '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                      '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "2" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                 '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                 '�߰������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                    '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                      '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "3" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '����Է���"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                 '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                 '�߰������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = ? "                                 '����������"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = ? "                                 '���������Ͻ�"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                    '����ڵ�
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '���������
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                      '��������Ͻ�
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        End If
                        
                        SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "                                          '��ü��ȣ"
                        SQL = SQL & vbCrLf & "   AND EXMN_CD = ? "                                          '�˻��ڵ�"
                        SQL = SQL & vbCrLf & "   AND PID = ? "                                              'ȯ�ڹ�ȣ"
                        SQL = SQL & vbCrLf & "   AND RSLT_STAT < ? "                                        '�������"

                        Set AdoCmd_ORACLE = New ADODB.Command
                        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser                                         'ADOConnection
                        
                        AdoCmd_ORACLE.CommandType = adCmdText
                        AdoCmd_ORACLE.CommandText = SQL
                        
                        '-- �ý��� ��¥ �������� �Լ� : gsDBDateTime
                        
                        If Send_State = "1" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "2" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "3" Then

                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        End If
                            
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_NO", adVarChar, , 1, sCnt)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 15, lsPid)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                        
                        
                        
                        AdoCmd_ORACLE.Execute sqlRet
                        Set AdoCmd_ORACLE = Nothing
                End If
            '/------------------------------------
            
            '/------------------------------------ �������̺� STATE ������Ʈ
                If Send_State = "" Then cn_Ser.RollbackTrans: Exit Function
                
                If Send_State = "1" Then
                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ? " & vbCrLf
                ElseIf Send_State = "2" Then
                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ? " & vbCrLf
                ElseIf Send_State = "3" Then
                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ?, LAST_RPTG_DT = ? " & vbCrLf
                End If
                SQL = SQL & " WHERE SPCM_NO = ? " & vbCrLf
                SQL = SQL & "   AND RSLT_STAT <> ? " & vbCrLf
                SQL = SQL & "   AND SPCM_STAT = ? " & vbCrLf
                SQL = SQL & "   AND EXMN_CD IN (?,?,?,?) "
                
                Set AdoCmd_ORACLE = New ADODB.Command
                Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
                
                AdoCmd_ORACLE.CommandType = adCmdText
                AdoCmd_ORACLE.CommandText = SQL
                
                '-- �ý��� ��¥ �������� �Լ� : gsDBDateTime
                
                If Send_State = "1" Then
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
                    
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(GetText(.vasTemp, iRow, 2)))
                    
                ElseIf Send_State = "2" Then
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 15, "2")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , 5, gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
                    
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(GetText(.vasTemp, iRow, 2)))

                ElseIf Send_State = "3" Then
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
                    
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(GetText(.vasTemp, iRow, 2)))
                End If
                
                AdoCmd_ORACLE.Execute sqlRet
                Set AdoCmd_ORACLE = Nothing
                '##### ���ε� ���� - 7 ##############################################
     
                If sqlRet < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            
            '/------------------------------------
            End If
        Next iRow
        
        '/------------------------------------ �������̺� STATE ������Ʈ
         If Send_State = "" Then cn_Ser.RollbackTrans: Exit Function
                
         Dim MDMD_EXAMCODE   As String
         Dim MDMD_LIST
         Dim sParam_string   As String
         
         MDMD_EXAMCODE = ""
         MDMD_LIST = Split(gState_G, ",")
         For i = 0 To UBound(MDMD_LIST) - 1
             MDMD_EXAMCODE = MDMD_EXAMCODE & ",'" & MDMD_LIST(i) & "'"
         Next i
         MDMD_LIST = Split(gState_M, ",")
         For i = 0 To UBound(MDMD_LIST) - 1
             MDMD_EXAMCODE = MDMD_EXAMCODE & ",'" & MDMD_LIST(i) & "'"
         Next i
         MDMD_LIST = Split(gState_B, ",")
         For i = 0 To UBound(MDMD_LIST) - 1
             MDMD_EXAMCODE = MDMD_EXAMCODE & ",'" & MDMD_LIST(i) & "'"
         Next i
         
         MDMD_EXAMCODE = MDMD_EXAMCODE & "," & ExamCode_Spec
         MDMD_EXAMCODE = Mid(MDMD_EXAMCODE, 2)
         MDMD_LIST = Split(Replace(MDMD_EXAMCODE, "'", ""), ",")
         
         For i = 0 To UBound(MDMD_LIST)
             If sParam_string <> "" Then
                 sParam_string = sParam_string & ",?"
             Else
                 sParam_string = ",?"
             End If
         Next i
         
         sParam_string = Mid(sParam_string, 2)
         
        If Send_State = "1" Then
            SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ? " & vbCrLf
        ElseIf Send_State = "2" Then
            SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ? " & vbCrLf
        ElseIf Send_State = "3" Then
            SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ?, LAST_RPTG_DT = ? " & vbCrLf
        End If
        SQL = SQL & " WHERE SPCM_NO = ? " & vbCrLf
        SQL = SQL & "   AND RSLT_STAT <> ? " & vbCrLf
        SQL = SQL & "   AND SPCM_STAT = ? " & vbCrLf
        SQL = SQL & "   AND EXMN_CD IN (" & sParam_string & ")"
        
        Set AdoCmd_ORACLE = New ADODB.Command
        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
        
        AdoCmd_ORACLE.CommandType = adCmdText
        AdoCmd_ORACLE.CommandText = SQL
        
        '-- �ý��� ��¥ �������� �Լ� : gsDBDateTime
        
        If Send_State = "1" Then
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
            
        ElseIf Send_State = "2" Then
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 15, "2")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , 5, gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")

        ElseIf Send_State = "3" Then
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
            
        End If
             
             For i = 0 To UBound(MDMD_LIST)
                 AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(MDMD_LIST(i)))
             Next i
             
'                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
'                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
'                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
'                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(GetText(.vasTemp, iRow, 2)))
        
        AdoCmd_ORACLE.Execute sqlRet
        Set AdoCmd_ORACLE = Nothing
        '##### ���ε� ���� - 7 ##############################################

        If sqlRet < 0 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------

''        '/------------------------------------ ó�����̺� STATE ������Ʈ
''
''        '///////// ó�����̺�
''        SQL = "UPDATE SPSLMJBBI "
''        If Send_State = "1" Then
''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & .lblUser.Caption & "', "
''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''        ElseIf Send_State = "2" Then
''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & .lblUser.Caption & "', "
''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''        ElseIf Send_State = "3" Then
''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & .lblUser.Caption & "', "
''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''        End If
''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
''        res = SendQuery(gServer, SQL)
''
''        If res = -1 Then
''            SaveQuery SQL
''            cn_Ser.RollbackTrans
''            Exit Function
''
''        End If
''        '/------------------------------------

        '/------------------------------------ ó�����̺� STATE ������Ʈ
            SQL = "UPDATE SPSLMJBBI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = SYSDATE " & vbCrLf
            SQL = SQL & " WHERE SPCM_NO = ? " & vbCrLf
            SQL = SQL & "   AND RSLT_STAT < ? " & vbCrLf
            SQL = SQL & "   AND SPCM_STAT = ? " & vbCrLf
            SQL = SQL & "   AND PID = ? "
            
            Set AdoCmd_ORACLE = New ADODB.Command
            Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
            AdoCmd_ORACLE.CommandType = adCmdText
            AdoCmd_ORACLE.CommandText = SQL
            
            If Send_State = "1" Then
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
            ElseIf Send_State = "2" Then
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
            ElseIf Send_State = "3" Then
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            End If
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 20, lsPid)
            
            AdoCmd_ORACLE.Execute sqlRet
            Set AdoCmd_ORACLE = Nothing
        '/------------------------------------
        
'''''        '/------------------------------------ ó�����̺� ������Ʈ(MDMDHTORD)
'''''        If Send_State = "3" Then
'''''
'''''                           SQL = "UPDATE MDMDHTORD "
'''''            SQL = SQL & vbCrLf & "   SET PRSC_STAT = '51'"      '/50 ���񺸰�, 51 ��������
'''''            SQL = SQL & vbCrLf & "     , RPTG_DT = SYSDATE"
'''''            SQL = SQL & vbCrLf & "     , AMEN_ID = '" & .lblUser.Caption & "'"
'''''            SQL = SQL & vbCrLf & " WHERE (PRSC_SQNO, PRSC_CD) "
'''''            SQL = SQL & vbCrLf & "       IN (SELECT PRSC_SQNO, EXMN_CD "
'''''            SQL = SQL & vbCrLf & "             FROM SPSLMJBDI "
'''''            SQL = SQL & vbCrLf & "            WHERE SPCM_NO = '" & lsSpecNo & "' "
'''''            SQL = SQL & vbCrLf & "              AND EXMN_CD IN (" & Trim(MDMD_EXAMCODE) & ") "                     '�˻��ڵ�"
'''''            SQL = SQL & vbCrLf & "              AND SPCM_STAT = '2') "
'''''            SQL = SQL & vbCrLf & "AND DC_DVSN = 'O' "
'''''
'''''            Save_Raw_Data "[ó�������Ʈ]" & SQL
'''''
'''''            res = SendQuery(gServer, SQL)
'''''            If res = -1 Then
'''''                SaveQuery SQL
'''''                cn_Ser.RollbackTrans
'''''                Exit Function
'''''            End If
'''''        End If
        
                  SQL = "UPDATE MDMDHTORD "
            SQL = SQL & vbCrLf & "   SET PRSC_STAT = '51'"      '/50 ���񺸰�, 51 ��������
            SQL = SQL & vbCrLf & "     , RPTG_DT = SYSDATE"
            SQL = SQL & vbCrLf & "     , AMEN_ID = '" & .lblUser.Caption & "'"
            SQL = SQL & vbCrLf & "     , updt_DT = SYSDATE"
            SQL = SQL & vbCrLf & " WHERE PRSC_SQNO = "
            SQL = SQL & "                   (SELECT PRSC_SQNO "
            SQL = SQL & vbCrLf & "             FROM SPSLMJBDI "
            SQL = SQL & vbCrLf & "            WHERE BRCD_LABL_NO = '" & lsID & "') "
'            SQL = SQL & vbCrLf & "            WHERE SPCM_NO = '" & lsSpecNo & "' "
'            SQL = SQL & vbCrLf & "              AND EXMN_CD IN (" & Trim(MDMD_EXAMCODE) & ") "                     '�˻��ڵ�"
'            SQL = SQL & vbCrLf & "              AND SPCM_STAT = '2') "
            SQL = SQL & vbCrLf & "AND DC_DVSN = 'O' "
            
            Save_Raw_Data "[ó�������Ʈ]" & SQL
            
            res = SendQuery(gServer, SQL)
            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        
        '/------------------------------------ ó�����̺� ������Ʈ
        
        
        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data = 1
    End With
End Function


Public Function gsDBDateTime() As Date
Dim sRs     As ADODB.Recordset
Dim strSQL  As String

    Set sRs = New ADODB.Recordset
    
    strSQL = "select sysdate from dual"
    sRs.Open strSQL, cn_Ser, adOpenStatic, adLockReadOnly
    
    If Not sRs.EOF Then
        gsDBDateTime = sRs("SYSDATE")
    Else
        gsDBDateTime = Now
    End If
    sRs.Close
    Set sRs = Nothing

End Function


Function Insert_Data_R(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'������ ����Ÿ ���̽��� ����
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim lsInsertTime    As String
    Dim sCnt            As String
    

    Insert_Data_R = -1
    
    lsID = ""
    lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode))
    lsSpecNo = Trim(GetText(frmInterface.vasRID, argSpcRow, colSpecNo))
    lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, colPID))
    lsInsertTime = Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
    'lsInsertTime = Trim(Format(GetDateFull, "yyyymmddhhmmss"))
    
    
    'Local���� ȯ�ں��� ����� ��������
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
          " And barcode = '" & Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
          " And diskno = '" & Trim(GetText(frmInterface.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
          " And posno = '" & Trim(GetText(frmInterface.vasRID, argSpcRow, colPos)) & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    frmInterface.vasTemp.MaxRows = frmInterface.vasTemp.DataRowCnt + 1
    
    gHIVPosFlag = -1
    
    sCnt = ""
    'db_BeginTran gServer
    '������ ����� �����ϱ�
    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
        sCnt = ""
        
        SQL = "SELECT RSLT_NO FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '�˻��ڵ�"
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
        SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '�������"
        res = db_select_Col(gServer, SQL)
        sCnt = CLng(gReadBuf(0)) + 1
        
        SQL = "UPDATE SPSLHRRST "
        SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "                   '���(�����)
        SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "                   '���(�������)"
        SQL = SQL & vbCrLf & "       DLTA_YN = 'N', "                                                           'Delta üũ"
        SQL = SQL & vbCrLf & "       PANC_YN = 'N', "                                                           'Panic üũ"
        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = 'test', "                                                   '����Է���"
        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = 'test', "                                                   '�߰�������"
        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�"
        SQL = SQL & vbCrLf & "       LAST_RPTR_ID = 'test', "                                                   '����������"
        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
        SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "                                                        '���������
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
        SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "                                                          '�������"
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '�˻��ڵ�"
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
        SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '�������"
        res = SendQuery(gServer, SQL)
        If res < 0 Then
            SaveQuery SQL
           ' db_RollBack gServer
            Exit Function
        End If
        
    Next iRow
    
    
    
    
    SQL = "SELECT EXMN_CD FROM SPSLHRRST "
    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND AND EXMN_CD NOT LIKE '%G%' "
    SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
    SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
    
    If res = 0 Then                                                                 '///// ������̺� ����� �� �� �ִ� ��� (�׷��ڵ�����)
        SQL = "Update SPSLMJBBI"
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3'"
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    
        SQL = "Update SPSLMJBDI"
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3'"
        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
    ElseIf res = -1 Then                                                             '///// ���� �����ΰ��
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = ""

       
    'db_Commit gServer
    Insert_Data_R = 1

End Function

'Function Get_Sample_Info_SPCMNO(ByVal asRow As Long) As Integer
'
'    Dim sBarcode As String
'    Dim sSpecNo As String
'    Dim sTestCd As String
'
'    Get_Sample_Info_SPCMNO = -1
'    'ȯ������ ��������
'    sSpecNo = Trim(GetText(frmInterface.vasResult, asRow, colSpecNo))
'    sTestCd = Trim(GetText(frmInterface.vasResult, asRow, colTestCd))
'
'    If sSpecNo = "" Then
'        Exit Function
'    End If
'    '���ڵ��ȣ�� ��ü��ȣ �ҷ�����FN_LABCVTPRTBCNO(SPCM_NO) --> ���ڵ�󺧹�ȣ ����
'
'    SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(sSpecNo) & "') FROM DUAL "
'    res = db_select_Col(gServer, SQL)
'    sBarcode = Trim(gReadBuf(0))
'
'    'ȯ�ڹ�ȣ, ȯ���̸�, �ֹι�ȣ, ����, ����
'    SQL = "SELECT PID, PT_NM, SEX, AGE "
'    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
'    SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
'    SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
'    SQL = SQL & vbCrLf & "  AND RSLT_STAT < '2' "
'    res = db_select_Col(gServer, SQL)
'
'    '///////// gAllExam �ڸ��� �˻� �ڵ� �־��� �����ڵ� �� �پ� �ִ°� B312001 , 02, 03
'
'    If res = 1 Then
'        SetText frmInterface.vasResult, Trim(sSpecNo), asRow, colSpecNo     '2
'        SetText frmInterface.vasResult, Trim(sBarcode), asRow, colBarcode   '3
'        SetText frmInterface.vasResult, Trim(sTestCd), asRow, colTestCd    '4
'        SetText frmInterface.vasResult, Trim(gReadBuf(0)), asRow, colPID    '6
'        SetText frmInterface.vasResult, Trim(gReadBuf(1)), asRow, colPName  '7
'        SetText frmInterface.vasResult, Trim(gReadBuf(2)), asRow, colSex    '8
'        SetText frmInterface.vasResult, Trim(gReadBuf(3)), asRow, colAge    '9
'        Get_Sample_Info_SPCMNO = 1
'    Else
'        Get_Sample_Info_SPCMNO = -1
'    End If
'
'End Function


Function Get_Sample_Info_QC(ByVal asRow As Long) As Integer

    Dim sBarcode As String
    Dim sQCdate  As String
    
    Get_Sample_Info_QC = -1
    'ȯ������ ��������
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '���� ���ڵ� ��ȣ
    If sBarcode = "" Or IsNumeric(sBarcode) = False Then
        Exit Function
    End If
    
    sQCdate = Trim(Format(GetDateFull, "yyyymmdd"))
    
    'ȯ�ڹ�ȣ, ȯ���̸�, �ֹι�ȣ, ����, ����
    SQL = "SELECT SBSN_NO, '��������', '', "
    SQL = SQL & vbCrLf & "                 (SELECT MAX(RSLT_SQNO) + 1 FROM SPSLHQRST "
    SQL = SQL & vbCrLf & "                   WHERE EQPM_CD = '" & Mid(sBarcode, 3, 3) & "' "
    SQL = SQL & vbCrLf & "                     AND SBSN_CD = '" & Mid(sBarcode, 6, 3) & "' "
    SQL = SQL & vbCrLf & "                     AND LVL_CD  = '" & Mid(sBarcode, 9, 1) & "' "
    SQL = SQL & vbCrLf & "                     AND EXMN_DY = '" & sQCdate & "' )"
    SQL = SQL & vbCrLf & " FROM SPSLMQMST "
    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(sBarcode, 3, 3) & "' "
    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(sBarcode, 6, 3) & "' "
    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(sBarcode, 9, 1) & "' "
    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam �ڸ��� �˻� �ڵ� �־��� �����ڵ� �� �پ� �ִ°� B312001 , 02, 03
    
    If res = 1 Then
        SetText frmInterface.vasID, Trim(sBarcode), asRow, colSpecNo
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge
        
        SetText frmInterface.vasList, Trim(sBarcode), 1, colSpecNo
        SetText frmInterface.vasList, Trim(gReadBuf(0)), 1, colPID
        SetText frmInterface.vasList, Trim(gReadBuf(1)), 1, colPName
        SetText frmInterface.vasList, Trim(gReadBuf(2)), 1, colSex
        SetText frmInterface.vasList, Trim(gReadBuf(3)), 1, colAge
        
        Get_Sample_Info_QC = 1
    Else
    
        Get_Sample_Info_QC = -1
        Call SaveQuery(SQL)
    End If

End Function

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim sRet    As String
    
On Error GoTo Err

    Get_Sample_Info = -1
    'ȯ������ ��������
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '##### ���ε� ���� - 99 ##############################################
''    '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
''    SQL = "SELECT FN_LABCVTBCNO('" & Trim(sBarcode) & "') FROM DUAL "
''    res = db_select_Col(gServer, SQL)
''    sSpecNo = Trim(gReadBuf(0))
    
    SQL = "SELECT FN_LABCVTBCNO(?) FROM DUAL "
    
    Save_Raw_Data "[Get_Sample_Info]" & SQL
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("FN_LABCVTBCNO", adVarChar, , 10, Trim(sBarcode))
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    sSpecNo = AdoRs_ORACLE.Fields(0) & ""
    
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    

    '##### ���ε� ���� - 99 ##############################################
    
    '##### ���ε� ���� - 98 ##############################################
''    'ȯ�ڹ�ȣ, ȯ���̸�, �ֹι�ȣ, ����, ����
''    SQL = "SELECT PID, PT_NM, SEX, AGE "
''    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
''    SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
''    SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
''    SQL = SQL & vbCrLf & "  AND RSLT_STAT <> '3' "
''    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam �ڸ��� �˻� �ڵ� �־��� �����ڵ� �� �پ� �ִ°� B312001 , 02, 03
    
''    If res = 1 Then
''        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo     '2
''        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID    '6
''        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName  '7
''        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '8
''        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge    '9
''        Get_Sample_Info = 1
''    Else
''        Get_Sample_Info = -1
''    End If

    SQL = "SELECT PID, PT_NM, SEX, AGE "
    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
    SQL = SQL & vbCrLf & "WHERE SPCM_NO = ? "
    SQL = SQL & vbCrLf & "  AND SPCM_STAT = ? "
'    SQL = SQL & vbCrLf & "  AND RSLT_STAT <> ? "
    
    Save_Raw_Data "[Get_Sample_Info]" & SQL
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(sSpecNo))
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic

    If AdoRs_ORACLE.BOF = False Then
        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo     '2
        SetText frmInterface.vasID, Trim(AdoRs_ORACLE.Fields(0) & ""), asRow, colPID    '6
        SetText frmInterface.vasID, Trim(AdoRs_ORACLE.Fields(1) & ""), asRow, colPName  '7
        SetText frmInterface.vasID, Trim(AdoRs_ORACLE.Fields(2) & ""), asRow, colSex    '8
        SetText frmInterface.vasID, Trim(AdoRs_ORACLE.Fields(3) & ""), asRow, colAge    '9
        Get_Sample_Info = 1
    Else
        '�ڵ�����
        'wsdl url�� ���� ������� �Ʒ� �ּҷ� ����Ͻø� �˴ϴ�. http://isis.nhimc:8800/service/PoctService?wsdl
        'PoctService ���� ���񽺴� registSpcmRcpn(String sBcno, String sPoctDevModel ) ȣ���Ͻø� �˴ϴ�.
        'sBcno�� ��񿡼� �Ѿ�� ��ü��ȣ(18�� �����ϴ� ���ڸ�����)�̸�, sPoctDevModel�� �Ϲ������� ��񱸺��� ���� ����text �Դϴ�.
        
'        sRet = Online_XML_Qry(sBarcode)
        
        
        Get_Sample_Info = -1
        
        
    End If

    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    '##### ���ε� ���� - 98 ##############################################

Exit Function

Err:
    Save_Raw_Data "[Get_Sample_Info]" & Err.Description

    Get_Sample_Info = -1
    
End Function



Public Function Online_XML_Qry(ByVal sBcno As String) As String
    Dim oSOAP   As MSSOAPLib30.SoapClient30
    Dim Send    As String
    Dim sParam  As String
    Dim txtSendXML  As String
    
    On Error GoTo ErrHandle
    
    Online_XML_Qry = ""
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    oSOAP.MSSoapInit gServerPath
            
    txtSendXML = "<?xml version='1.0' encoding='UTF-8'?>"
    txtSendXML = txtSendXML & vbCrLf & "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>"
    txtSendXML = txtSendXML & vbCrLf & "<soapenv:Body>"
    txtSendXML = txtSendXML & vbCrLf & "<registSpcmRcpn xmlns='http://svc.poct.ws.nhimc/'>"
    txtSendXML = txtSendXML & vbCrLf & "<arg0 xmlns=''>" & Trim(sBcno) & "</arg0>"
    txtSendXML = txtSendXML & vbCrLf & "<arg1 xmlns=''>" & gEquip & "</arg1>"
    txtSendXML = txtSendXML & vbCrLf & "</registSpcmRcpn>"
    txtSendXML = txtSendXML & vbCrLf & "</soapenv:Body>"
    txtSendXML = txtSendXML & vbCrLf & "</soapenv:Envelope>" & vbCrLf


    Call Save_Raw_Data("[Send SOAP  => " & sBcno & " ]" & gEquip)
    
    Send = oSOAP.registSpcmRcpn(sBcno, gEquip)
    
    
    Call Save_Raw_Data("[Recv SOAP => " & sBcno & " ]" & Send)
    
    Online_XML_Qry = Send
    Set oSOAP = Nothing
    
    DoEvents
    
    Exit Function
    
ErrHandle:
    Online_XML_Qry = ""
    
    Save_Raw_Data "[Online_XML_Qry]" & Err.Description

    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If


End Function

Function Get_Sample_InfoR(ByVal asRow As Long) As Integer
   Dim sBarcode As String
    Dim sSpecNo As String

    Get_Sample_InfoR = -1
    'ȯ������ ��������
    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBarcode))   '���� ���ڵ� ��ȣ
    If sBarcode = "" Then
        Exit Function
    End If
    '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
    SQL = "SELECT FN_LABCVTBCNO(" & Trim(sBarcode) & ") FROM DUAL "
    res = db_select_Col(gServer, SQL)
    
    sSpecNo = Trim(gReadBuf(0))
    
    'ȯ�ڹ�ȣ, ȯ���̸�, �ֹι�ȣ, ����, ����
    SQL = "SELECT PID, PT_NM, SEX, AGE "
    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
    SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
    SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "  AND RSLT_STAT = '0' "
    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam �ڸ��� �˻� �ڵ� �־��� �����ڵ� �� �پ� �ִ°� B312001 , 02, 03
    
    If res = 1 Then
        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge
        
        Get_Sample_InfoR = 1
    Else
    
        Get_Sample_InfoR = -1
    End If
End Function

Function EquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String

    EquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp1)
    
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
    SQL = " Select SUCD From LRESULT " & CR & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    res = db_select_Col(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        EquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function


Function GetEquipExamCode_CA1500(argEquipCode As String, argPID As String) As String
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

    GetEquipExamCode_CA1500 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
'    argPID = "1558200030"
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    res = db_select_Col(gServer, SQL)
    GetEquipExamCode_CA1500 = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            sExamCd = Trim(gReadBuf(i))
        Else
            Exit For
        End If
    Next
    
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and SUBSTR(exmn_cd,1,1) <> 'G'" & _
          "   and RSLT_NO IS NOT NULL"
          
    res = db_select_Row(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'EquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    Erase gReadBuf
          SQL = "Select equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    SQL = SQL & " order by equipcode    "
    res = db_select_Row(gLocal, SQL)
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
    
    GetEquipExamCode_CA1500 = strExamCode
    
End Function

Function GetOrderExamCode(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String

    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    res = db_select_Col(gServer, SQL)
    GetOrderExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            sExamCd = Trim(gReadBuf(i))
        Else
            Exit For
        End If
    Next
    
    '-- �˻��ڵ� ��������
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    res = db_select_Col(gServer, SQL)
    GetOrderExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
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

On Error GoTo Err

    GetOrderExamCode_New = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        sExamCd = Trim(rs_svr.Fields(0) & "")
        rs_svr.MoveNext
    Loop
    Set rs_svr = Nothing
    
    '-- �˻��ڵ� ��������
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
    
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_New = GetOrderExamCode_New & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_New <> "" Then
        GetOrderExamCode_New = Mid(GetOrderExamCode_New, 1, Len(GetOrderExamCode_New) - 1)
    End If
    
Exit Function

Err:
    GetOrderExamCode_New = ""

End Function

Function GetOrderExamCode_MIC(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

    GetOrderExamCode_MIC = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sExamCd = argPID
    
    '-- �˻��ڵ� ��������
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"

    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_MIC = GetOrderExamCode_MIC & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop

    If GetOrderExamCode_MIC <> "" Then
        GetOrderExamCode_MIC = Mid(GetOrderExamCode_MIC, 1, Len(GetOrderExamCode_MIC) - 1)
    End If

    '-- �ӽ� �׽�Ʈ��
'    GetOrderExamCode_MIC = "'L41000'"
    
End Function


Function GetEquipExamCode_E411(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim sSpecNo     As String

    GetEquipExamCode_E411 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
    SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
    res = db_select_Col(gServer, SQL)
    sSpecNo = Trim(gReadBuf(0))
    
    '-- �˻��ڵ� ��������
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    res = db_select_Row(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    If strExamCode = "" Then
'        MsgBox "������ ȯ��"
        GetEquipExamCode_E411 = ""
        Exit Function
    End If
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'EquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    
    '-- ������ �˻��ڵ��� ä�� ã��
          SQL = "Select distinct equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    
    res = db_select_Row(gLocal, SQL)
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
    
    GetEquipExamCode_E411 = Mid(strExamCode, 2)
    
End Function

'Function GetEquipExamCode_CA1500(argEquipCode As String, argPID As String) As String
''��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
''�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
'Dim i As Integer
'Dim sExamCode As String
'Dim strExamCode As String
'Dim sSpecNo     As String
'
'    GetEquipExamCode_CA1500 = ""
'
'    If Trim(argEquipCode) = "" Then
'        Exit Function
'    End If
'
'    '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
'    SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
'    res = db_select_Col(gServer, SQL)
'    sSpecNo = Trim(gReadBuf(0))
'
'    '-- �˻��ڵ� ��������
'    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
'          " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
'          "   and RSLT_NO IS NOT NULL"
'
'    res = db_select_Row(gServer, SQL)
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
'    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
'    'EquipExamCode =
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
'    res = db_select_Row(gLocal, SQL)
'    strExamCode = ""
'    For i = 0 To UBound(gReadBuf)
'
'        If gReadBuf(i) <> "" Then
'            'gReadBuf(i) = Mid(gReadBuf(i), 1, Len(gReadBuf(i)) - 1)
'            strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
'        Else
'            Exit For
'        End If
'    Next
'
'    GetEquipExamCode_CA1500 = Mid(strExamCode, 2)
'
'End Function

Function GetEquipExamCode(argEquipCode As String, argPID As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String

    GetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    res = db_select_Col(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'EquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
          SQL = "Select equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(argEquipCode) & ")"
    
    res = db_select_Col(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & Trim(gReadBuf(i)) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    GetEquipExamCode = strExamCode
    
End Function


