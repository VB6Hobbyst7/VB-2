Attribute VB_Name = "DB_Query"
Option Explicit

Public Const colCheckBox = 1
Public Const colSpecNo = 2
Public Const colBarcode = 3
Public Const colRack = 4
Public Const colPos = 5
Public Const colPID = 6
Public Const colPName = 7
Public Const colSex = 8
Public Const colAge = 9
Public Const colOCnt = 10
Public Const colRCnt = 11
Public Const colState = 12
Public Const colHct = 13


'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
Public Const colEquipCode = 1
Public Const colExamCode = 2
Public Const colExamName = 3
Public Const colResult = 4
Public Const colSeq = 5
Public Const colFLAG = 6
Public Const colEquipResult = 7
Public Const colDelta = 8
Public Const colPanic = 9

'����ڵ�� �˻��ڵ� ã��
Public gEquipExamCode As String

'�ش�˻翡 ���� �Ҽ���
Public gExamRange As String
'����ġ �� �˻��
Public gExamName As String
Public gRFVL_DVSN As String
Public gMALE_HIGH As String
Public gMALE_LOW As String
Public gFEML_HIGH As String
Public gFEML_LOW As String
Public gDELT_DVSN As String
Public gDELT_HIGH As String
Public gDELT_LOW As String
Public gDELT_DD As String
Public gPANC_DVSN As String
Public gPANC_HIGH As String
Public gPANC_LOW As String

Public gTLA_Equip As String
Public gTLA_Sub1 As String
Public gTLA_Sub2 As String

'////Lasc
Public gEXAM_CBC        As String
Public gEXAM_Diff       As String
Public gEXAM_Reti       As String
Public gEXAM_CBC_Diff   As String

'////Comment
Public gComment_All As String
Public gComment_Code As String


Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As Long = 0)
    Dim sCnt As String
    Dim sExamDate As String
    Dim RCnt As Integer
    Dim OCnt As Integer
    
'    SQL = "SELECT COUNT(*) FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
'          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'    res = db_select_Col(gLocal, SQL)
    With frmInterface
        sExamDate = Format(.dtpToday, "yyyymmdd")
        
        SQL = "DELETE FROM PAT_RES " & vbCrLf & _
              "WHERE EXAMDATE = '" & Format(.dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
              "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "  AND BARCODE = '" & Trim(GetText(.vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(GetText(.vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
              "  and examcode= '" & Trim(GetText(.vasRes, asRow2, colExamCode)) & "'"
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        
        SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
              "POSNO, PID, PNAME, " & vbCrLf & _
              "PSEX, PAGE, " & vbCrLf & _
              "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
              "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, EQUIPRESULT, RECENO, RESFLAG) " & vbCrLf & _
              "VALUES('" & gEquip & "', '" & Trim(GetText(.vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(.vasID, asRow1, colRack)) & "', " & vbCrLf & _
              "'" & Trim(GetText(.vasID, asRow1, colPos)) & "', '" & Trim(GetText(.vasID, asRow1, colPID)) & "', '" & Trim(GetText(.vasID, asRow1, colPName)) & "', " & vbCrLf & _
              "'" & Trim(GetText(.vasID, asRow1, colSex)) & "', " & 0 & ", " & vbCrLf & _
              "'" & Trim(sExamDate) & "', '" & Trim(GetText(.vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(.vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
              "'" & Trim(GetText(.vasRes, asRow2, colSeq)) & "', '" & Trim(GetText(.vasRes, asRow2, colResult)) & "', '" & Trim(GetText(.vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
              "'" & asSend & "', '" & Trim(GetText(.vasRes, asRow2, 7)) & "', '" & Trim(asEquipResult) & "', '" & Trim(GetText(.vasID, asRow1, colSpecNo)) & "', '') "
        res = SendQuery(gLocal, SQL)

    End With
End Function
''''
''''Function Insert_Data_MIC(ByVal argSpcRow As Integer) As Integer
''''    Dim iRow            As Integer
''''    Dim i               As Integer
''''    Dim j               As Integer
''''    Dim lsID            As String
''''    Dim lsSpecNo        As String
''''    Dim lsPid           As String
''''    Dim sResult         As String
''''    Dim lsInsertTime    As String
''''    Dim sCnt            As String
''''
''''On Error GoTo Err
''''
''''    Insert_Data_MIC = -1
''''
''''    lsID = ""
''''    lsID = Trim(GetText(frmInterface.vasResult, argSpcRow, colBarcode))
''''    lsSpecNo = Trim(GetText(frmInterface.vasResult, argSpcRow, colSpecNo))
''''    lsPid = Trim(GetText(frmInterface.vasResult, argSpcRow, 5))
''''    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
''''
''''    If lsSpecNo = "" Then
''''        Exit Function
''''    End If
''''
''''    'Local���� ȯ�ں��� ����� ��������
''''    ClearSpread frmInterface.vasTemp
''''
''''    SQL = " Select isocd, equipcode, examcode, result, antsize, EQUIPRESULT, refflag, panicflag, deltaflag,exmncd " & vbCrLf & _
''''          "   From pat_res " & vbCrLf & _
''''          " Where equipno = '" & gEquip & "' " & vbCrLf & _
''''          "   And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
''''          "   And barcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 3)) & "' " & vbCrLf & _
''''          "   And examcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' " & vbCrLf & _
''''          "   And receno = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 2)) & "' " & vbCrLf & _
''''          "   And isocd = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 7)) & "' "
''''    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
''''
''''    If res = -1 Then
''''        SaveQuery SQL
''''        Exit Function
''''    End If
''''
''''    frmInterface.vasTemp.MaxRows = frmInterface.vasTemp.DataRowCnt + 1
''''
''''    gHIVPosFlag = -1
''''
''''    sCnt = ""
''''
''''    cn_Ser.BeginTrans
''''
''''    '������ ����� �����ϱ�
''''    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
''''        If iRow = 1 Then
''''            '-- �̻��� ���հ�� ��ȸ
''''            SQL = "SELECT BCTR_SQNO FROM SPSLHMBAC "
''''            'SQL = "SELECT COUNT(*) FROM SPSLHMBAC "
''''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 'SPCM_NO    ��ü��ȣ
''''            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "' "         'EXMN_CD    �˻��ڵ�
''''            SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "' "         'BCTR_CD    �����ڵ�
''''
''''            res = db_select_Col(gServer, SQL)
''''            If res > 0 Then
''''                sCnt = CLng(gReadBuf(0)) + 1
''''            Else
''''                sCnt = 1
''''            End If
'''''            If res > 0 Then
'''''                '-- �̻��� ���հ���� ������ �ش� ���հ���� �����Ѵ�
'''''                SQL = "DELETE FROM SPSLHMBAC "
'''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                             'SPCM_NO    ��ü��ȣ
'''''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "' "     'EXMN_CD    �˻��ڵ�
'''''                SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "' "     'BCTR_CD    �����ڵ�
'''''
'''''                res = SendQuery(gServer, SQL)
'''''
'''''                If res < 0 Then
'''''                    SaveQuery SQL
'''''                    cn_Ser.RollbackTrans
'''''                    Exit Function
'''''                End If
'''''            End If
''''
''''            '-- �̻��� �׻������ ��ȸ
'''''            SQL = "SELECT SPCM_NO FROM SPSLHMANT "
'''''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 'SPCM_NO    ��ü��ȣ
'''''            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "' "         'EXMN_CD    �˻��ڵ�
'''''            SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "' "         'BCTR_CD    �����ڵ�
'''''
'''''            res = db_select_Col(gServer, SQL)
'''''
'''''            If res > 0 Then
'''''                '-- �̻��� �׻�������� ������ �ش� ���տ� ���� �׻��� ����� ��� �����Ѵ�
'''''                SQL = "DELETE FROM SPSLHMANT "
'''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                             'SPCM_NO    ��ü��ȣ
'''''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "' "     'EXMN_CD    �˻��ڵ�
'''''                SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "' "     'BCTR_CD    �����ڵ�
'''''
'''''                res = SendQuery(gServer, SQL)
'''''
'''''                If res < 0 Then
'''''                    SaveQuery SQL
'''''                    cn_Ser.RollbackTrans
'''''                    Exit Function
'''''                End If
'''''            End If
''''
''''            '-- �̻��� ���հ�� ����
''''            SQL = "INSERT INTO SPSLHMBAC (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      SORT_SEQ,"
''''            SQL = SQL & vbCrLf & "        SPCM_CD,      CLTR_VOL_CD,    CLTR_PERD,      PRE_RSLT_CD,    LAST_BCTR_CD,"
''''            SQL = SQL & vbCrLf & "        RSLT_RPTR_ID, RSLT_RPTG_DT,   MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,"
''''            SQL = SQL & vbCrLf & "        LAST_RPTG_DT, RSLT_STAT,      CMNT_DVSN,      EXMN_EQPM,      RMRK,   REGI_ID,   RGST_DT,AMEN_ID,UPDT_DT) "
''''            SQL = SQL & vbCrLf & " Values ( "
''''            SQL = SQL & vbCrLf & " '" & lsID & "', "                                            'SPCM_NO        ��ü��ȣ
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "    'EXMN_CD        �˻��ڵ�
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'BCTR_CD        �����ڵ�
''''            SQL = SQL & vbCrLf & " '" & sCnt & "', "                                            'BCTR_SQNO      �����Ϸù�ȣ:��ȣ-N5
''''            SQL = SQL & vbCrLf & " '" & iRow & "', "                                            'SORT_SEQ       ���ļ���
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 10)) & "', "   'SPCM_CD        ��ü�ڵ�
''''            SQL = SQL & vbCrLf & " '', "                                                        'CLTR_VOL_CD    ��緮�ڵ�:�����ڵ�
''''            SQL = SQL & vbCrLf & " '', "                                                        'CLTR_PERD      ���Ⱓ:����-V200
''''            SQL = SQL & vbCrLf & " '', "                                                        'PRE_RSLT_CD    �������ڵ�:�����ڵ�
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'LAST_BCTR_CD   ���������ڵ�
''''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'RSLT_RPTR_ID   ���������ID:������ȣ
''''            SQL = SQL & vbCrLf & " sysdate, "                                                   'RSLT_RPTG_DT   ��������Ͻ�:��¥-DT
''''            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTR_ID   �߰�������ID:������ȣ
''''            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTG_DT   �߰������Ͻ�:��¥-DT
''''            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   ����������ID:������ȣ
''''            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   ���������Ͻ�:��¥-DT
''''            SQL = SQL & vbCrLf & " '1', "                                                       'RSLT_STAT      �������:�����ڵ� ==> ������ : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT �Է�]    ?? �˻�� �����Ը� �������Ѵٰ� ��.
''''                                                                                                '                                     ���񺸰� : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT �Է�]
''''                                                                                                '                                     �������� : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT �Է�]
''''            SQL = SQL & vbCrLf & " '', "                                                        'CMNT_DVSN      �ڸ�Ʈ����
''''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        ����ڵ�:�����ڵ�
''''            SQL = SQL & vbCrLf & " '', "                                                        'RMRK           ���
''''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'REGI_ID        �����ID:������ȣ
''''            SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        ����Ͻ�:��¥-DT
''''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'AMEN_ID        ���������
''''            SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        ��������Ͻ�
''''
''''            res = SendQuery(gServer, SQL)
''''
''''            If res < 0 Then
''''                SaveQuery SQL
''''                cn_Ser.RollbackTrans
''''                Exit Function
''''            End If
''''
''''            SQL = "SELECT RSLT_NO FROM SPSLHRRST "
''''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 '��ü��ȣ"
''''            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' "  '�˻��ڵ�"
''''            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
''''            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
''''            res = db_select_Col(gServer, SQL)
''''            If res > 0 Then
''''                sCnt = CLng(gReadBuf(0)) + 1
''''
''''                '-- ������� ���ڰ��� ��츸 ��Ÿ/�д� ������ �Ѵ�.
''''                sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
''''                If IsNumeric(sResult) Then
''''                    Dim strDecision     As Variant
''''                    Dim strBarcode      As String
''''
''''                    strBarcode = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
''''                    strDecision = GetDecision(argSpcRow, strBarcode, iRow)
''''                    strDecision = Split(strDecision, "|")
''''                Else
''''                    strDecision = "||"
''''                    strDecision = Split(strDecision, "|")
''''                End If
''''
''''                SQL = "UPDATE SPSLHRRST "   '-- ������̺�
''''                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "      '���(�����)
''''                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "      '���(�������)"
''''                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & strDecision(0) & "', "                                    'H/L üũ"
''''                SQL = SQL & vbCrLf & "       DLTA_YN = '" & strDecision(1) & "', "                                      'Delta üũ"
''''                SQL = SQL & vbCrLf & "       PANC_YN = '" & strDecision(2) & "', "                                      'Panic üũ"
''''                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "', "                                     '����Է���"
''''                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
''''                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
''''                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "                                          '���������
''''                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
''''                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
''''                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '�������"
''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 '��ü��ȣ"
''''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' "  '�˻��ڵ�"
''''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
''''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
''''                res = SendQuery(gServer, SQL)
''''
''''                If res < 0 Then
''''                    SaveQuery SQL
''''                    cn_Ser.RollbackTrans
''''                    Exit Function
''''                End If
''''
''''            End If
''''
''''        End If
''''
''''        '-- �̻��� �׻������ ����
''''        If Trim(GetText(frmInterface.vasTemp, iRow, 2)) <> "" Then
''''            SQL = "INSERT INTO SPSLHMANT (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      ANTB_CD,"
''''            SQL = SQL & vbCrLf & "        SPCM_CD,      ANTB_RSLT,      DTRM_RSLT,      ANTB_EXMN_MTHD, RSLT_RPTR_ID,"
''''            SQL = SQL & vbCrLf & "        RSLT_RPTG_DT, MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,   LAST_RPTG_DT,"
''''            SQL = SQL & vbCrLf & "        RSLT_STAT,    EXMN_EQPM,      REGI_ID,        RGST_DT,AMEN_ID,UPDT_DT)"
''''            SQL = SQL & vbCrLf & " Values ( "
''''            SQL = SQL & vbCrLf & " '" & lsID & "', "                                            'SPCM_NO        ��ü��ȣ
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "    'EXMN_CD        �˻��ڵ�
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'BCTR_CD        �����ڵ�
''''            SQL = SQL & vbCrLf & " '" & sCnt & "', "                                            'BCTR_SQNO      �����Ϸù�ȣ:��ȣ-N5
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "', "    'ANTB_CD        �׻����ڵ�:�����ڵ�
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 10)) & "', "   'SPCM_CD        ��ü�ڵ�
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 5)) & "', "    'ANTB_RSLT      �׻������:�з�-V50
''''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "    'DTRM_RSLT      �������
''''            SQL = SQL & vbCrLf & " 'M', "                                                       'ANTB_EXMN_MTHD �׻����˻���:�����ڵ� ==> �˻��� M : MICRO �� �ǹ�
''''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'RSLT_RPTR_ID   ���������ID:������ȣ
''''            SQL = SQL & vbCrLf & " sysdate, "                                                   'RSLT_RPTG_DT   ��������Ͻ�:��¥-DT
''''            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTR_ID   �߰�������ID:������ȣ
''''            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTG_DT   �߰������Ͻ�:��¥-DT
''''            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   ����������ID:������ȣ
''''            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   ���������Ͻ�:��¥-DT
''''            SQL = SQL & vbCrLf & " '1', "                                                       'RSLT_STAT      �������:�����ڵ� ==> ������ : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT �Է�]    ?? �˻�� �����Ը� �������Ѵٰ� ��.
''''                                                                                                '                                     ���񺸰� : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT �Է�]
''''                                                                                                '                                     �������� : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT �Է�]
''''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        ����ڵ�:�����ڵ�
''''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'REGI_ID        �����ID:������ȣ
''''            SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        ����Ͻ�:��¥-DT
''''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'AMEN_ID        ���������
''''            SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        ��������Ͻ�
''''            res = SendQuery(gServer, SQL)
''''
''''            If res < 0 Then
''''                SaveQuery SQL
''''                cn_Ser.RollbackTrans
''''                Exit Function
''''            End If
''''        End If
''''
''''    Next iRow
''''
''''
''''    SQL = "SELECT EXMN_CD FROM SPSLHRRST "  '-- �������̺�
''''    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
''''    SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
''''    SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
''''    SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
''''    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
''''
''''    If res = 0 Then                                                                 '///// ������̺� ����� �� �� �ִ� ��� (�׷��ڵ�����)
''''        SQL = "Update SPSLMJBBI"    '-- ��ü���̺�
''''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
''''        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
''''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
''''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
''''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
''''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
''''        res = SendQuery(gServer, SQL)
''''
''''        If res = -1 Then
''''            SaveQuery SQL
''''            cn_Ser.RollbackTrans
''''            Exit Function
''''        End If
''''
''''        SQL = "Update SPSLMJBDI"    '-- ó�����̺�
''''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
''''        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
''''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
''''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
''''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
''''        res = SendQuery(gServer, SQL)
''''
''''        If res = -1 Then
''''            SaveQuery SQL
''''            cn_Ser.RollbackTrans
''''            Exit Function
''''        End If
''''
''''    ElseIf res = -1 Then                                                             '///// ���� �����ΰ��
''''        SaveQuery SQL
''''        cn_Ser.RollbackTrans
''''        Exit Function
''''
''''    Else
''''        SQL = "Update SPSLMJBBI"    '-- ��ü���̺�
''''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
''''        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
''''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
''''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
''''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
''''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
''''        res = SendQuery(gServer, SQL)
''''
''''        If res = -1 Then
''''            SaveQuery SQL
''''            Exit Function
''''        End If
''''
''''        SQL = "Update SPSLMJBDI"    '-- ó�����̺�
''''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
''''        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
''''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
''''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
''''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
''''        res = SendQuery(gServer, SQL)
''''
''''        If res = -1 Then
''''            SaveQuery SQL
''''            cn_Ser.RollbackTrans
''''            Exit Function
''''        End If
''''
''''    End If
''''
''''    SQL = ""
''''
''''    cn_Ser.CommitTrans
''''
''''    Insert_Data_MIC = 1
''''
''''    Exit Function
''''
''''Err:fk
''''    cn_Ser.RollbackTrans
''''
''''
''''End Function

Function Insert_Data_SE(ByVal argSpcRow As Integer, asSend_gubun As String, asDataGubun As String) As Integer
    '/�������̺� ��� �ֱ�
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
    Dim ExamCode_Remark     As String
    
    Dim Mid_RESULTCD    As String   '/�߰����� ����ڵ�
    Dim LAST_RESULTCD   As String   '/�������� ����ڵ�
    Dim RANGE_RESULTCD   As String   '/���Ⱓ �ڵ�
    
    Dim State_GM    As String       '//// �׷�/��Ƽ �ڵ�
    Dim State_cnt   As Integer      '//// �׷�/��Ƽ �ڵ� �� ����
    Dim State_G     As String       '//// �׷��ڵ�
    Dim State_M     As String       '//// ��Ƽ�ڵ�
    Dim State_B     As String       '//// ���͸��ڵ�
    
    Dim Send_State As String
    Dim EXAM_CD     As String
    Dim SPEC_CD     As String
    If asSend_gubun = 0 Then: Exit Function
    With frmInterface
        gComment_All = ""
        Insert_Data_SE = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        lsID = ""
        EXAM_CD = ""
        SPEC_CD = ""
        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
         
        If asDataGubun = "1" Then
            lsID = Trim(GetText(.vasWorkList, argSpcRow, colBarcode))
                           SQL = "SELECT EXAMCODE, RESULT, WORKNO FROM SPEC_RESULT "
            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(.vasWorkList, argSpcRow, colBarcode)) & "' "
            SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(.vasWorkList, argSpcRow, colPos)) & "' "
            res = db_select_Vas(gLocal, SQL, .vasTemp)
        Else
            lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
                           SQL = "SELECT EXAMCODE, RESULT, RESULTTIME, WORKNO FROM SEND_RESULT "
            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' "
            SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
            res = db_select_Vas(gLocal, SQL, .vasTemp)
        End If
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        
        
        Mid_RESULTCD = ""
        LAST_RESULTCD = ""
        
        If asSend_gubun = "2" Then
            Mid_RESULTCD = "NGA2"
            RANGE_RESULTCD = ""
            LAST_RESULTCD = ""
        ElseIf asSend_gubun = "3" Then
            LAST_RESULTCD = "NGBL"
            Mid_RESULTCD = ""
            RANGE_RESULTCD = "0112"
        End If
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To 1
            
            SQL = "SELECT FN_LABCVTBCNO(" & Trim(lsID) & ") FROM DUAL "
            res = db_select_Col(gServer, SQL)
            
            lsSpecNo = gReadBuf(0)
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 2))
            

            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" And sResult1 <> "POSITIVE" Then
                
                       
                If asSend_gubun = "2" Then
                    State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 1)))
                
'                    State_cnt = InStr(1, State_GM, "/")
'                    State_G = Mid(State_GM, 1, State_cnt - 1)
'                    State_GM = Mid(State_GM, State_cnt + 1)
'                    State_cnt = InStr(1, State_GM, "/")
'                    State_M = Mid(State_GM, 1, State_cnt - 1)
'                    State_B = Mid(State_GM, State_cnt + 1)
'
'                    If State_M <> "" Then
'                        EXAM_CD = State_M
'                        SQL = "SELECT SPCM_CD FROM SPSLHRRST "
'                        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
'                        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_M) & "' "
'                        res = db_select_Col(gServer, SQL)
'                    Else
                        EXAM_CD = Trim(Trim(GetText(.vasTemp, iRow, 1)))
                        SQL = "SELECT SPCM_CD FROM SPSLHRRST "
                        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(Trim(GetText(.vasTemp, iRow, 1))) & "' "
                        res = db_select_Col(gServer, SQL)
'                    End If
                    SPEC_CD = gReadBuf(0)
                    If SPEC_CD = "" Then cn_Ser.RollbackTrans: Exit Function
                    SQL = "SELECT MAX(BCTR_SQNO) FROM SPSLHMBAC "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(EXAM_CD) & "' "
                    res = db_select_Col(gServer, SQL)
                    If Val(gReadBuf(0)) > 0 Then Exit Function
                    
                    gReadBuf(0) = Val(gReadBuf(0)) + 1
                                        
                    SQL = "INSERT INTO SPSLHMBAC (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      SORT_SEQ,"
                    SQL = SQL & vbCrLf & "        SPCM_CD,      CLTR_VOL_CD,    CLTR_PERD,      PRE_RSLT_CD,    LAST_BCTR_CD,"
                    SQL = SQL & vbCrLf & "        RSLT_RPTR_ID, RSLT_RPTG_DT,   MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,"
                    SQL = SQL & vbCrLf & "        LAST_RPTG_DT, RSLT_STAT,      CMNT_DVSN,      EXMN_EQPM,      RMRK,   REGI_ID,   RGST_DT, AMEN_ID, UPDT_DT) "
                    SQL = SQL & vbCrLf & " Values ( "
                    SQL = SQL & vbCrLf & " '" & Trim(lsSpecNo) & "', "    'SPCM_NO        ��ü��ȣ
                    SQL = SQL & vbCrLf & " '" & Trim(EXAM_CD) & "', "    'EXMN_CD        �˻��ڵ�
                    SQL = SQL & vbCrLf & " '" & Mid_RESULTCD & "', "                                    'BCTR_CD        �����ڵ�             (���������� ��)
                    SQL = SQL & vbCrLf & " " & gReadBuf(0) & ", "                                       'BCTR_SQNO      �����Ϸù�ȣ:��ȣ-N5
                    SQL = SQL & vbCrLf & " 1, "                                                         'SORT_SEQ       ���ļ���
                    SQL = SQL & vbCrLf & " '" & SPEC_CD & "', "                                         'SPCM_CD        ��ü�ڵ�
                    SQL = SQL & vbCrLf & " '', "                                                        'CLTR_VOL_CD    ��緮�ڵ�:�����ڵ�
                    SQL = SQL & vbCrLf & " '', "                                                        'CLTR_PERD      ���Ⱓ:����-V200
                    SQL = SQL & vbCrLf & " '" & Mid_RESULTCD & "', "                                    'PRE_RSLT_CD    �������ڵ�:�����ڵ�(No Growth 2 Day) �ϵ��ڵ� ����
                    SQL = SQL & vbCrLf & " '', "                                                        'LAST_BCTR_CD   ���������ڵ�         (Bact�� No Growth �� ����)
                    SQL = SQL & vbCrLf & " '" & gEquipCode & "_INF', "                                  'RSLT_RPTR_ID   ���������ID:������7
                    SQL = SQL & vbCrLf & " sysdate, "                                                   'RSLT_RPTG_DT   ��������Ͻ�:��¥-DT
                    SQL = SQL & vbCrLf & " '" & gEquipCode & "_INF', "                                  'MDDL_RPTR_ID   �߰�������ID:������ȣ
                    SQL = SQL & vbCrLf & " sysdate, "                                                   'MDDL_RPTG_DT   �߰������Ͻ�:��¥-DT
                    SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   ����������ID:������ȣ
                    SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   ���������Ͻ�:��¥-DT
                    SQL = SQL & vbCrLf & " '1', "                                                       'RSLT_STAT      �������:�����ڵ� ==> ������ : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT �Է�]    ?? �˻�� �����Ը� �������Ѵٰ� ��.
                                                                                                        '                                     ���񺸰� : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT �Է�]
                                                                                                        '                                     �������� : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT �Է�]
                    SQL = SQL & vbCrLf & " '', "                                                        'CMNT_DVSN      �ڸ�Ʈ����
                    SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        ����ڵ�:�����ڵ�
                    SQL = SQL & vbCrLf & " '', "                                                        'RMRK           ���
                    SQL = SQL & vbCrLf & " '" & gEquipCode & "_INF', "                                  'REGI_ID        �����ID:������ȣ
                    SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        ����Ͻ�:��¥-DT
                    SQL = SQL & vbCrLf & " '" & gEquipCode & "_INF', "                                  'AMEN_ID        ���������
                    SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        ��������Ͻ�
                    
                    res = SendQuery(gServer, SQL)
                    
                    If res < 0 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                    
                ElseIf asSend_gubun = "3" Then
'                    State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 1)))
'
'                    State_cnt = InStr(1, State_GM, "/")
'                    State_G = Mid(State_GM, 1, State_cnt - 1)
'                    State_GM = Mid(State_GM, State_cnt + 1)
'                    State_cnt = InStr(1, State_GM, "/")
'                    State_M = Mid(State_GM, 1, State_cnt - 1)
'                    State_B = Mid(State_GM, State_cnt + 1)
                
                                   SQL = "UPDATE SPSLHMBAC "
                    SQL = SQL & vbCrLf & "   SET BCTR_CD = '" & LAST_RESULTCD & "' "
                    SQL = SQL & vbCrLf & "      ,LAST_BCTR_CD = '" & LAST_RESULTCD & "' "
                    SQL = SQL & vbCrLf & "      ,CLTR_PERD = '" & RANGE_RESULTCD & "' "
                    SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = '" & gEquipCode & "_INF' "
                    SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = sysdate "
                    SQL = SQL & vbCrLf & "      ,RSLT_STAT = '3' "
                    SQL = SQL & vbCrLf & "      ,AMEN_ID = '" & gEquipCode & "_INF' "
                    SQL = SQL & vbCrLf & "      ,UPDT_DT = sysdate "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 1)) & "' "
                    
                    res = SendQuery(gServer, SQL)
                    
                    If res < 0 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
                    
            End If
        Next iRow
       
        cn_Ser.CommitTrans
        Insert_Data_SE = 1
    End With
End Function

Function Insert_Data_SE_FIRST(ByVal argSpcRow As Integer, asSend_gubun As String, asDataGubun As String) As Integer
    '/�������̺� ��� �ֱ�
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
    Dim ExamCode_Remark     As String
    
    Dim Mid_RESULTCD    As String   '/�߰����� ����ڵ�
    Dim LAST_RESULTCD   As String   '/�������� ����ڵ�
    Dim RANGE_RESULTCD   As String   '/���Ⱓ �ڵ�
    
    Dim State_GM    As String       '//// �׷�/��Ƽ �ڵ�
    Dim State_cnt   As Integer      '//// �׷�/��Ƽ �ڵ� �� ����
    Dim State_G     As String       '//// �׷��ڵ�
    Dim State_M     As String       '//// ��Ƽ�ڵ�
    Dim State_B     As String       '//// ���͸��ڵ�
    
    Dim Send_State As String
    Dim EXAM_CD     As String
    Dim SPEC_CD     As String
    'If asSend_gubun = 0 Then: Exit Function
    With frmInterface
        gComment_All = ""
        Insert_Data_SE_FIRST = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        lsID = ""
        EXAM_CD = ""
        SPEC_CD = ""
        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
         
        lsID = Trim(GetText(.vasWorkList, argSpcRow, colBarcode))
                       SQL = "SELECT EXAMCODE, RESULT, WORKNO FROM SPEC_RESULT "
        SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(.vasWorkList, argSpcRow, colBarcode)) & "' "
        SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(.vasWorkList, argSpcRow, colPos)) & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp)

        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        
        
        Mid_RESULTCD = ""
        LAST_RESULTCD = ""
        
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To 1 '.vasTemp.DataRowCnt
            Save_Raw_Data GetDateFull & " [Insert_Data_SE_FIRST 1 ] "
            If Trim(GetText(.vasTemp, iRow, 1)) <> "BPF" Then
                Save_Raw_Data GetDateFull & " [Insert_Data_SE_FIRST 2 ] "
                SQL = "SELECT FN_LABCVTBCNO(" & Trim(lsID) & ") FROM DUAL "
                res = db_select_Col(gServer, SQL)
                
                lsSpecNo = gReadBuf(0)
                
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 1)))
            
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_GM = Mid(State_GM, State_cnt + 1)
                State_cnt = InStr(1, State_GM, "/")
                State_M = Mid(State_GM, 1, State_cnt - 1)
                State_B = Mid(State_GM, State_cnt + 1)
    
    '            If State_M <> "" Then
                '/��� ������ ��� �ö� ��� ���̺� INSERT �մϴ�.
                    EXAM_CD = State_M
                    SQL = "SELECT SPCM_CD FROM SPSLHRRST "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_M) & "' "
                    res = db_select_Col(gServer, SQL)
    '            Else
    '                EXAM_CD = Trim(Trim(GetText(.vasTemp, iRow, 1)))
    '                SQL = "SELECT SPCM_CD FROM SPSLHRRST "
    '                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
    '                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(Trim(GetText(.vasTemp, iRow, 1))) & "' "
    '                res = db_select_Col(gServer, SQL)
    '            End If
                If Trim(GetText(.vasTemp, iRow, 1)) = "" Then cn_Ser.RollbackTrans: Exit Function
                
                sCnt = Mid(Trim(GetText(.vasTemp, iRow, 1)), Len(Trim(GetText(.vasTemp, iRow, 1))), 1)
                
                If sCnt = "2" Then
                    ExamCode_Remark = "ANBO"
                ElseIf sCnt = "1" Then
                    ExamCode_Remark = "AEBO"
                End If
                
                
                SPEC_CD = gReadBuf(0)
                If SPEC_CD = "" Then cn_Ser.RollbackTrans: Exit Function
                SQL = "SELECT MAX(BCTR_SQNO) FROM SPSLHMBAC "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(EXAM_CD) & "' "
                res = db_select_Col(gServer, SQL)
                
                If Val(gReadBuf(0)) >= 2 Then cn_Ser.RollbackTrans: Exit Function
                gReadBuf(0) = Val(gReadBuf(0)) + 1
                                    
                SQL = "INSERT INTO SPSLHMBAC (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      SORT_SEQ,"
                SQL = SQL & vbCrLf & "        SPCM_CD,      CLTR_VOL_CD,    CLTR_PERD,      PRE_RSLT_CD,    LAST_BCTR_CD,"
                SQL = SQL & vbCrLf & "        RSLT_RPTR_ID, RSLT_RPTG_DT,   MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,"
                SQL = SQL & vbCrLf & "        LAST_RPTG_DT, RSLT_STAT,      CMNT_DVSN,      EXMN_EQPM,      RMRK,   REGI_ID,   RGST_DT, AMEN_ID, UPDT_DT) "
                SQL = SQL & vbCrLf & " Values ( "
                SQL = SQL & vbCrLf & " '" & Trim(lsSpecNo) & "', "                                  'SPCM_NO        ��ü��ȣ
                SQL = SQL & vbCrLf & " '" & Trim(EXAM_CD) & "', "                                   'EXMN_CD        �˻��ڵ�
                SQL = SQL & vbCrLf & " ' ', "                                                        'BCTR_CD        �����ڵ�             (���������� ��)
                SQL = SQL & vbCrLf & " " & gReadBuf(0) & ", "                                       'BCTR_SQNO      �����Ϸù�ȣ:��ȣ-N5
                SQL = SQL & vbCrLf & " 1, "                                                         'SORT_SEQ       ���ļ���
                SQL = SQL & vbCrLf & " '" & SPEC_CD & "', "                                         'SPCM_CD        ��ü�ڵ�
                SQL = SQL & vbCrLf & " '', "                                                        'CLTR_VOL_CD    ��緮�ڵ�:�����ڵ�
                SQL = SQL & vbCrLf & " '', "                                                        'CLTR_PERD      ���Ⱓ:����-V200
                SQL = SQL & vbCrLf & " '', "                                                        'PRE_RSLT_CD    �������ڵ�:�����ڵ�(No Growth 2 Day) �ϵ��ڵ� ����
                SQL = SQL & vbCrLf & " '', "                                                        'LAST_BCTR_CD   ���������ڵ�         (Bact�� No Growth �� ����)
                SQL = SQL & vbCrLf & " '', "                                                        'RSLT_RPTR_ID   ���������ID:������7
                SQL = SQL & vbCrLf & " '', "                                                        'RSLT_RPTG_DT   ��������Ͻ�:��¥-DT
                SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTR_ID   �߰�������ID:������ȣ
                SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTG_DT   �߰������Ͻ�:��¥-DT
                SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   ����������ID:������ȣ
                SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   ���������Ͻ�:��¥-DT
                SQL = SQL & vbCrLf & " '0', "                                                       'RSLT_STAT      �������:�����ڵ� ==> ������ : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT �Է�]    ?? �˻�� �����Ը� �������Ѵٰ� ��.
                                                                                                    '                                     ���񺸰� : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT �Է�]
                                                                                                    '                                     �������� : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT �Է�]
                SQL = SQL & vbCrLf & " '', "                                                        'CMNT_DVSN      �ڸ�Ʈ����
                SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        ����ڵ�:�����ڵ�
                SQL = SQL & vbCrLf & " '', "                                                        'RMRK           ���
                SQL = SQL & vbCrLf & " '" & gEquipCode & "_INF', "                                  'REGI_ID        �����ID:������ȣ
                SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        ����Ͻ�:��¥-DT
                SQL = SQL & vbCrLf & " '" & gEquipCode & "_INF', "                                  'AMEN_ID        ���������
                SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        ��������Ͻ�
                
                res = SendQuery(gServer, SQL)
                
                If res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            Else '/�Ҿư�ü ������
                Save_Raw_Data GetDateFull & " [Insert_Data_SE_FIRST 3 ] "
                SQL = "SELECT FN_LABCVTBCNO(" & Trim(lsID) & ") FROM DUAL "
                res = db_select_Col(gServer, SQL)
                
                lsSpecNo = gReadBuf(0)
                

    '            If State_M <> "" Then
                '/��� ������ ��� �ö� ��� ���̺� INSERT �մϴ�.
                    EXAM_CD = "L41001"
                    SQL = "SELECT SPCM_CD FROM SPSLHRRST "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(EXAM_CD) & "' "
                    res = db_select_Col(gServer, SQL)
    '            Else
    '                EXAM_CD = Trim(Trim(GetText(.vasTemp, iRow, 1)))
    '                SQL = "SELECT SPCM_CD FROM SPSLHRRST "
    '                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
    '                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(Trim(GetText(.vasTemp, iRow, 1))) & "' "
    '                res = db_select_Col(gServer, SQL)
    '            End If
                
                SPEC_CD = gReadBuf(0)
                
                
'                sCnt = Mid(Trim(GetText(.vasTemp, iRow, 1)), Len(Trim(GetText(.vasTemp, iRow, 1))), 1)
'
'                If sCnt = "2" Then
'                    ExamCode_Remark = "ANBO"
'                ElseIf sCnt = "1" Then
'                    ExamCode_Remark = "AEBO"
'                End If
                
                If SPEC_CD = "" Then cn_Ser.RollbackTrans: Exit Function
                SQL = "SELECT MAX(BCTR_SQNO) FROM SPSLHMBAC "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(EXAM_CD) & "' "
                res = db_select_Col(gServer, SQL)
                
                For i = 1 To 2
                    If Val(gReadBuf(0)) > 2 Then cn_Ser.RollbackTrans: Exit Function
                    gReadBuf(0) = Val(gReadBuf(0)) + 1
                                        
                    SQL = "INSERT INTO SPSLHMBAC (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      SORT_SEQ,"
                    SQL = SQL & vbCrLf & "        SPCM_CD,      CLTR_VOL_CD,    CLTR_PERD,      PRE_RSLT_CD,    LAST_BCTR_CD,"
                    SQL = SQL & vbCrLf & "        RSLT_RPTR_ID, RSLT_RPTG_DT,   MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,"
                    SQL = SQL & vbCrLf & "        LAST_RPTG_DT, RSLT_STAT,      CMNT_DVSN,      EXMN_EQPM,      RMRK,   REGI_ID,   RGST_DT, AMEN_ID, UPDT_DT) "
                    SQL = SQL & vbCrLf & " Values ( "
                    SQL = SQL & vbCrLf & " '" & Trim(lsSpecNo) & "', "                                  'SPCM_NO        ��ü��ȣ
                    SQL = SQL & vbCrLf & " '" & Trim(EXAM_CD) & "', "                                   'EXMN_CD        �˻��ڵ�
                    SQL = SQL & vbCrLf & " ' ', "                                                        'BCTR_CD        �����ڵ�             (���������� ��)
                    SQL = SQL & vbCrLf & " " & gReadBuf(0) & ", "                                       'BCTR_SQNO      �����Ϸù�ȣ:��ȣ-N5
                    SQL = SQL & vbCrLf & " 1, "                                                         'SORT_SEQ       ���ļ���
                    SQL = SQL & vbCrLf & " '" & SPEC_CD & "', "                                         'SPCM_CD        ��ü�ڵ�
                    SQL = SQL & vbCrLf & " '', "                                                        'CLTR_VOL_CD    ��緮�ڵ�:�����ڵ�
                    SQL = SQL & vbCrLf & " '', "                                                        'CLTR_PERD      ���Ⱓ:����-V200
                    SQL = SQL & vbCrLf & " '', "                                                        'PRE_RSLT_CD    �������ڵ�:�����ڵ�(No Growth 2 Day) �ϵ��ڵ� ����
                    SQL = SQL & vbCrLf & " '', "                                                        'LAST_BCTR_CD   ���������ڵ�         (Bact�� No Growth �� ����)
                    SQL = SQL & vbCrLf & " '', "                                                        'RSLT_RPTR_ID   ���������ID:������7
                    SQL = SQL & vbCrLf & " '', "                                                        'RSLT_RPTG_DT   ��������Ͻ�:��¥-DT
                    SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTR_ID   �߰�������ID:������ȣ
                    SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTG_DT   �߰������Ͻ�:��¥-DT
                    SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   ����������ID:������ȣ
                    SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   ���������Ͻ�:��¥-DT
                    SQL = SQL & vbCrLf & " '0', "                                                       'RSLT_STAT      �������:�����ڵ� ==> ������ : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT �Է�]    ?? �˻�� �����Ը� �������Ѵٰ� ��.
                                                                                                        '                                     ���񺸰� : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT �Է�]
                                                                                                        '                                     �������� : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT �Է�]
                    SQL = SQL & vbCrLf & " '', "                                                        'CMNT_DVSN      �ڸ�Ʈ����
                    SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        ����ڵ�:�����ڵ�
                    SQL = SQL & vbCrLf & " '', "                                                        'RMRK           ���
                    SQL = SQL & vbCrLf & " '" & gEquipCode & "_INF', "                                  'REGI_ID        �����ID:������ȣ
                    SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        ����Ͻ�:��¥-DT
                    SQL = SQL & vbCrLf & " '" & gEquipCode & "_INF', "                                  'AMEN_ID        ���������
                    SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        ��������Ͻ�
                    
                    res = SendQuery(gServer, SQL)
                    
                    If res < 0 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                Next i
                
            End If
            
        Next iRow
        Save_Raw_Data GetDateFull & " [Insert_Data_SE_FIRST 4 ] "
        
        'cn_Ser.RollbackTrans
        cn_Ser.CommitTrans
        Insert_Data_SE_FIRST = 1
    End With
End Function

Function Insert_Data_SE_MIDDLE(ByVal argSpcRow As Integer, asSend_gubun As String, asDataGubun As String) As Integer
    '/�������̺� ��� �ֱ�
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
    Dim ExamCode_Remark     As String
    
    Dim Mid_RESULTCD    As String   '/�߰����� ����ڵ�
    Dim LAST_RESULTCD   As String   '/�������� ����ڵ�
    Dim RANGE_RESULTCD   As String   '/���Ⱓ �ڵ�
    
    Dim State_GM    As String       '//// �׷�/��Ƽ �ڵ�
    Dim State_cnt   As Integer      '//// �׷�/��Ƽ �ڵ� �� ����
    Dim State_G     As String       '//// �׷��ڵ�
    Dim State_M     As String       '//// ��Ƽ�ڵ�
    Dim State_B     As String       '//// ���͸��ڵ�
    
    Dim Send_State As String
    Dim EXAM_CD     As String
    Dim SPEC_CD     As String
    'If asSend_gubun = 0 Then: Exit Function
    With frmInterface
        gComment_All = ""
        Insert_Data_SE_MIDDLE = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        lsID = ""
        EXAM_CD = ""
        SPEC_CD = ""
        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
         

        lsID = Trim(GetText(.vasWorkList, argSpcRow, colBarcode))
                       SQL = "SELECT EXAMCODE, RESULT, WORKNO FROM SPEC_RESULT "
        SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(.vasWorkList, argSpcRow, colBarcode)) & "' "
        SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(.vasWorkList, argSpcRow, colPos)) & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        
        
        Mid_RESULTCD = ""
        LAST_RESULTCD = ""
        
        
        Mid_RESULTCD = "NGA2"
        RANGE_RESULTCD = ""
        LAST_RESULTCD = ""
        
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
             Save_Raw_Data GetDateFull & " [Insert_Data_SE_MIDDLE 1 ] "
            SQL = "SELECT FN_LABCVTBCNO(" & Trim(lsID) & ") FROM DUAL "
            res = db_select_Col(gServer, SQL)
            
            lsSpecNo = gReadBuf(0)
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 2))
            
            If Trim(GetText(.vasTemp, iRow, 1)) <> "BPF" And Trim(GetText(.vasTemp, iRow, 1)) <> "" Then
                Save_Raw_Data GetDateFull & " [Insert_Data_SE_MIDDLE 1 ] "
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 1)))
            
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_GM = Mid(State_GM, State_cnt + 1)
                State_cnt = InStr(1, State_GM, "/")
                State_M = Mid(State_GM, 1, State_cnt - 1)
                State_B = Mid(State_GM, State_cnt + 1)

                EXAM_CD = State_M
                sCnt = Mid(Trim(GetText(.vasTemp, iRow, 1)), Len(Trim(GetText(.vasTemp, iRow, 1))), 1)
                
                If sCnt = "2" Then
                    ExamCode_Remark = "ANBO"
                ElseIf sCnt = "1" Then
                    ExamCode_Remark = "AEBO"
                End If
                
                
                If EXAM_CD = "" Then cn_Ser.RollbackTrans: Exit Function
                
                               SQL = "UPDATE SPSLHMBAC "
                SQL = SQL & vbCrLf & "   SET BCTR_CD = '" & Mid_RESULTCD & "' "
                SQL = SQL & vbCrLf & "      ,PRE_RSLT_CD = '" & Mid_RESULTCD & "' "
                SQL = SQL & vbCrLf & "      ,CLTR_PERD = '" & RANGE_RESULTCD & "' "
                
                '/��������ÿ� ���� -------------------------------------------------
                'SQL = SQL & vbCrLf & "      ,LAST_BCTR_CD = '' "
                'SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = '" & gEquipCode & "_INF' "
                'SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = sysdate "
                '/��������ÿ� ���� -------------------------------------------------
                
                SQL = SQL & vbCrLf & "      ,RSLT_RPTR_ID = '" & gEquipCode & "_INF' "
                SQL = SQL & vbCrLf & "      ,RSLT_RPTG_DT = sysdate "
                SQL = SQL & vbCrLf & "      ,CMNT_DVSN = '" & ExamCode_Remark & "' "
                SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = '" & gEquipCode & "_INF' "
                SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = sysdate "
                SQL = SQL & vbCrLf & "      ,RSLT_STAT = '2' "
                SQL = SQL & vbCrLf & "      ,AMEN_ID = '" & gEquipCode & "_INF' "
                SQL = SQL & vbCrLf & "      ,UPDT_DT = sysdate "
                'SQL = SQL & vbCrLf & "      ,BCTR_SQNO = " & sCnt & " "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(EXAM_CD) & "' "
                SQL = SQL & vbCrLf & "   AND BCTR_SQNO = " & sCnt & " "
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                
                res = SendQuery(gServer, SQL)
                
                If res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            ElseIf Trim(GetText(.vasTemp, iRow, 1)) = "BPF" And Trim(GetText(.vasTemp, iRow, 1)) <> "" Then
                For i = 1 To 2
                    Save_Raw_Data GetDateFull & " [Insert_Data_SE_MIDDLE 2 ] "
                    EXAM_CD = "L41001"
                    sCnt = i
                    
                    If sCnt = "2" Then
                        ExamCode_Remark = "ANBO"
                    ElseIf sCnt = "1" Then
                        ExamCode_Remark = "AEBO"
                    End If
                    
                    
                    If EXAM_CD = "" Then cn_Ser.RollbackTrans: Exit Function
                    
                                   SQL = "UPDATE SPSLHMBAC "
                    SQL = SQL & vbCrLf & "   SET BCTR_CD = '" & Mid_RESULTCD & "' "
                    SQL = SQL & vbCrLf & "      ,PRE_RSLT_CD = '" & Mid_RESULTCD & "' "
                    SQL = SQL & vbCrLf & "      ,CLTR_PERD = '" & RANGE_RESULTCD & "' "
                    
                    '/��������ÿ� ���� -------------------------------------------------
                    'SQL = SQL & vbCrLf & "      ,LAST_BCTR_CD = '' "
                    'SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = '" & gEquipCode & "_INF' "
                    'SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = sysdate "
                    '/��������ÿ� ���� -------------------------------------------------
                    
                    SQL = SQL & vbCrLf & "      ,RSLT_RPTR_ID = '" & gEquipCode & "_INF' "
                    SQL = SQL & vbCrLf & "      ,RSLT_RPTG_DT = sysdate "
                    SQL = SQL & vbCrLf & "      ,CMNT_DVSN = '" & ExamCode_Remark & "' "
                    SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = '" & gEquipCode & "_INF' "
                    SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = sysdate "
                    SQL = SQL & vbCrLf & "      ,RSLT_STAT = '2' "
                    SQL = SQL & vbCrLf & "      ,AMEN_ID = '" & gEquipCode & "_INF' "
                    SQL = SQL & vbCrLf & "      ,UPDT_DT = sysdate "
                    'SQL = SQL & vbCrLf & "      ,BCTR_SQNO = " & sCnt & " "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(EXAM_CD) & "' "
                    SQL = SQL & vbCrLf & "   AND BCTR_SQNO = " & sCnt & " "
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                    res = SendQuery(gServer, SQL)
                    
                    If res < 0 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                Next i
            End If
        Next iRow
        Save_Raw_Data GetDateFull & " [Insert_Data_SE_MIDDLE 4 ] "
        'cn_Ser.RollbackTrans
        cn_Ser.CommitTrans
        Insert_Data_SE_MIDDLE = 1
    End With
End Function

Function Insert_Data_SE_LAST(ByVal argSpcRow As Integer, asSend_gubun As String, asDataGubun As String) As Integer
    '/�������̺� ��� �ֱ�
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
    Dim ExamCode_Remark     As String
    
    Dim Mid_RESULTCD    As String   '/�߰����� ����ڵ�
    Dim LAST_RESULTCD   As String   '/�������� ����ڵ�
    Dim RANGE_RESULTCD   As String   '/���Ⱓ �ڵ�
    
    Dim State_GM    As String       '//// �׷�/��Ƽ �ڵ�
    Dim State_cnt   As Integer      '//// �׷�/��Ƽ �ڵ� �� ����
    Dim State_G     As String       '//// �׷��ڵ�
    Dim State_M     As String       '//// ��Ƽ�ڵ�
    Dim State_B     As String       '//// ���͸��ڵ�
    
    Dim Send_State As String
    Dim EXAM_CD     As String
    Dim SPEC_CD     As String
    'If asSend_gubun = 0 Then: Exit Function
    With frmInterface
        gComment_All = ""
        Insert_Data_SE_LAST = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        lsID = ""
        EXAM_CD = ""
        SPEC_CD = ""
        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
         

        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
                       SQL = "SELECT EXAMCODE, RESULT, WORKNO FROM SEND_RESULT "
        SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' "
        SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        
        
        Mid_RESULTCD = ""
        RANGE_RESULTCD = ""
        LAST_RESULTCD = "NGBL"
        
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            
            SQL = "SELECT FN_LABCVTBCNO(" & Trim(lsID) & ") FROM DUAL "
            res = db_select_Col(gServer, SQL)
            
            lsSpecNo = gReadBuf(0)
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 2))
            
           If Trim(GetText(.vasTemp, iRow, 1)) <> "BPF" And Trim(GetText(.vasTemp, iRow, 1)) = "No Growth for 5 Days" Then
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 1)))
            
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_GM = Mid(State_GM, State_cnt + 1)
                State_cnt = InStr(1, State_GM, "/")
                State_M = Mid(State_GM, 1, State_cnt - 1)
                State_B = Mid(State_GM, State_cnt + 1)

                EXAM_CD = State_M
                sCnt = Mid(Trim(GetText(.vasTemp, iRow, 1)), Len(Trim(GetText(.vasTemp, iRow, 1))), 1)
                
'                If sCnt = "1" Then
'                    ExamCode_Remark = "INBO"
'                ElseIf sCnt = "2" Then
'                    ExamCode_Remark = "IEBO"
'                End If
                
                If EXAM_CD = "" Then cn_Ser.RollbackTrans: Exit Function
                
                               SQL = "UPDATE SPSLHMBAC "
                SQL = SQL & vbCrLf & "   SET BCTR_CD = '" & LAST_RESULTCD & "' "
'                SQL = SQL & vbCrLf & "      ,PRE_RSLT_CD = '" & Mid_RESULTCD & "' "
                SQL = SQL & vbCrLf & "      ,CLTR_PERD = '" & RANGE_RESULTCD & "' "
                
                '/��������ÿ� ���� -------------------------------------------------
                SQL = SQL & vbCrLf & "      ,LAST_BCTR_CD = '" & LAST_RESULTCD & "' "
                SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = '" & gEquipCode & "_INF' "
                SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = sysdate "
                '/��������ÿ� ���� -------------------------------------------------
                'SQL = SQL & vbCrLf & "      ,CMNT_DVSN = '" & ExamCode_Remark & "' "
'                SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = '" & gEquipCode & "_INF' "
'                SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = sysdate "
                SQL = SQL & vbCrLf & "      ,RSLT_STAT = '3' "
                SQL = SQL & vbCrLf & "      ,AMEN_ID = '" & gEquipCode & "_INF' "
                SQL = SQL & vbCrLf & "      ,UPDT_DT = sysdate "
                SQL = SQL & vbCrLf & "      ,BCTR_SQNO = " & sCnt & " "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(EXAM_CD) & "' "
                SQL = SQL & vbCrLf & "   AND BCTR_SQNO = " & sCnt & " "
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                res = SendQuery(gServer, SQL)
                
                If res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            ElseIf Trim(GetText(.vasTemp, iRow, 1)) = "BPF" And Trim(GetText(.vasTemp, iRow, 1)) = "No Growth for 5 Days" Then
                For i = 1 To 2
                    EXAM_CD = "L41001"
                    sCnt = i
                    
'                    If sCnt = "1" Then
'                        ExamCode_Remark = "INBO"
'                    ElseIf sCnt = "2" Then
'                        ExamCode_Remark = "IEBO"
'                    End If
                    
                    
                    If EXAM_CD = "" Then cn_Ser.RollbackTrans: Exit Function
                    
                                   SQL = "UPDATE SPSLHMBAC "
                    SQL = SQL & vbCrLf & "   SET BCTR_CD = '" & LAST_RESULTCD & "' "
    '                SQL = SQL & vbCrLf & "      ,PRE_RSLT_CD = '" & Mid_RESULTCD & "' "
                    SQL = SQL & vbCrLf & "      ,CLTR_PERD = '" & RANGE_RESULTCD & "' "
                    
                    '/��������ÿ� ���� -------------------------------------------------
                    SQL = SQL & vbCrLf & "      ,LAST_BCTR_CD = '" & LAST_RESULTCD & "' "
                    SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = '" & gEquipCode & "_INF' "
                    SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = sysdate "
                    '/��������ÿ� ���� -------------------------------------------------
                    'SQL = SQL & vbCrLf & "      ,CMNT_DVSN = '" & ExamCode_Remark & "' "
    '                SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = '" & gEquipCode & "_INF' "
    '                SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = sysdate "
                    SQL = SQL & vbCrLf & "      ,RSLT_STAT = '3' "
                    SQL = SQL & vbCrLf & "      ,AMEN_ID = '" & gEquipCode & "_INF' "
                    SQL = SQL & vbCrLf & "      ,UPDT_DT = sysdate "
'                    SQL = SQL & vbCrLf & "      ,BCTR_SQNO = " & sCnt & " "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(EXAM_CD) & "' "
                    SQL = SQL & vbCrLf & "   AND BCTR_SQNO = " & sCnt & " "
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                    res = SendQuery(gServer, SQL)
                    
                    If res < 0 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                Next i
                    
            End If
        Next iRow
        'cn_Ser.RollbackTrans
        cn_Ser.CommitTrans
        Insert_Data_SE_LAST = 1
    End With
End Function

Function Insert_Data(ByVal argSpcRow As Integer, asSend_gubun As String, asDataGubun As String) As Integer
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
    Dim ExamCode_Remark     As String
    
    Dim State_GM    As String       '//// �׷�/��Ƽ �ڵ�
    Dim State_cnt   As Integer      '//// �׷�/��Ƽ �ڵ� �� ����
    Dim State_G     As String       '//// �׷��ڵ�
    Dim State_M     As String       '//// ��Ƽ�ڵ�
    Dim State_B     As String       '//// ���͸��ڵ�
    
    Dim Send_State As String
    
    If asSend_gubun = 0 Then: Exit Function
    With frmInterface
        gComment_All = ""
        Insert_Data = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        lsID = ""

        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
         
        If asDataGubun = "1" Then
            lsID = Trim(GetText(.vasWorkList, argSpcRow, colBarcode))
                           SQL = "SELECT EXAMCODE, RESULT FROM SPEC_RESULT "
            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(.vasWorkList, argSpcRow, colBarcode)) & "' "
            SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(.vasWorkList, argSpcRow, colPos)) & "' "
            res = db_select_Vas(gLocal, SQL, .vasTemp)
        Else
            lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
                           SQL = "SELECT EXAMCODE, RESULT, RESULTTIME FROM SEND_RESULT "
            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' "
            SQL = SQL & vbCrLf & "   AND POS = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
            res = db_select_Vas(gLocal, SQL, .vasTemp)
        End If
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            
            If Trim(GetText(.vasTemp, iRow, 2)) = "POSITIVE" Then
                sResult1 = "No Growth " & _
                           Trim(CCur(GetText(.vasTemp, iRow, 3)) / 12) & _
                           " Days "
                           
            Else
                sResult1 = Trim(GetText(.vasTemp, iRow, 2))
            End If
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                
                
                SQL = "SELECT FN_LABCVTBCNO(" & Trim(lsID) & ") FROM DUAL "
                res = db_select_Col(gServer, SQL)
                
                lsSpecNo = gReadBuf(0)
                
                If Trim(GetText(.vasTemp, iRow, 1)) <> "BPF" And Trim(GetText(.vasTemp, iRow, 1)) <> "" Then
                
                    SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "                                           '��ü��ȣ"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 1)) & "' "                      '�˻��ڵ�"
                    
                    res = db_select_Col(gServer, SQL)
                     
                    If gReadBuf(0) = "" Then cn_Ser.RollbackTrans: Exit Function
                    
                    If gReadBuf(0) >= "0" And sCnt = "" Then
                        sCnt = CLng(gReadBuf(0)) + 1
                    End If
                    
                                   SQL = "UPDATE SPSLHRRST "
                    SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                    SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult1 & "', "                                          '���(�������)"
                    
                    
                    If asSend_gubun = "2" Then
                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                    '�߰�������"
                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '�߰������Ͻ�"
                    ElseIf asSend_gubun = "3" Then
                        If sResult1 = "POSITIVE" Then
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                    '����������"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                        Else
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                    '����������"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                        End If
                    End If
                    
                    SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                    SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
                    If asSend_gubun = "2" Then
                        SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "                                                      '�������" (1:�Է� , 2:�߰�����, 3:��������)
                    ElseIf asSend_gubun = "3" Then
                        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "                                                      '�������" (1:�Է� , 2:�߰�����, 3:��������)
                    End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "                                       '��ü��ȣ"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 1)) & "' "                     '�˻��ڵ�"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
'                    Debug.Print SQL
                    res = SendQuery(gServer, SQL)
                    
                    If res < 0 Then
                        SaveQuery SQL
                       ' db_RollBack gServer
                       cn_Ser.RollbackTrans
                        Exit Function
                    End If
                
                
                    State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 1)))
                    State_cnt = InStr(1, State_GM, "/")
                    State_G = Mid(State_GM, 1, State_cnt - 1)
                    State_GM = Mid(State_GM, State_cnt + 1)
                    State_cnt = InStr(1, State_GM, "/")
                    State_M = Mid(State_GM, 1, State_cnt - 1)
                    State_B = Mid(State_GM, State_cnt + 1)
                ElseIf Trim(GetText(.vasTemp, iRow, 1)) = "BPF" And Trim(GetText(.vasTemp, iRow, 1)) <> "" Then
                    
                    SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "                                           '��ü��ȣ"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & gAllExam & ") "                      '�˻��ڵ�"
                    res = db_select_Col(gServer, SQL)
                     
                    If gReadBuf(0) = "" Then cn_Ser.RollbackTrans: Exit Function
                    
                    If gReadBuf(0) >= "0" Then
                        sCnt = CLng(gReadBuf(0)) + 1
                    End If
                    
                                   SQL = "UPDATE SPSLHRRST "
                    SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                    SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult1 & "', "                                          '���(�������)"
                    
                    
                    If asSend_gubun = "2" Then
                        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                    '�߰�������"
                        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '�߰������Ͻ�"
                    ElseIf asSend_gubun = "3" Then
                        If sResult1 = "POSITIVE" Then
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                    '����������"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                        Else
                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                    '����������"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                        End If
                    End If

                    SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                    SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
                    If asSend_gubun = "2" Then
                        SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "                                                      '�������" (1:�Է� , 2:�߰�����, 3:��������)
                    ElseIf asSend_gubun = "3" Then
                        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "                                                      '�������" (1:�Է� , 2:�߰�����, 3:��������)
                    End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(lsSpecNo) & "' "                                       '��ü��ȣ"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & gAllExam & ") "                                            '�˻��ڵ�"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                    SQL = SQL & vbCrLf & "   AND EXMN_CD NOT IN ('L41001') "
                    SQL = SQL & vbCrLf & "   AND EXMN_CD <> PRSC_CD "
                    res = SendQuery(gServer, SQL)
                    If res < 0 Then
                        SaveQuery SQL
                       ' db_RollBack gServer
                       cn_Ser.RollbackTrans
                        Exit Function
                    End If
                
                
                    State_GM = RsltState_Check(lsSpecNo, "L41001")
                    State_cnt = InStr(1, State_GM, "/")
                    State_G = Mid(State_GM, 1, State_cnt - 1)
                    State_GM = Mid(State_GM, State_cnt + 1)
                    State_cnt = InStr(1, State_GM, "/")
                    State_M = Mid(State_GM, 1, State_cnt - 1)
                    State_B = Mid(State_GM, State_cnt + 1)
                
                
                End If
                
                
                
                '/------------------------------------ ������̺� �׷��ڵ� ���� ������Ʈ
                If Trim(State_G) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H �� ������ : �˻����� ��������� �ִ´�
                        If asSend_gubun = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                            'SQL = SQL & vbCrLf & "       REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                            'SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult1 & "', "                                          '���(�������)"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf asSend_gubun = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                            'SQL = SQL & vbCrLf & "       REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                            'SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult1 & "', "                                          '���(�������)"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf asSend_gubun = "3" Then

                            SQL = SQL & vbCrLf & " SET   LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '����������"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
                            'SQL = SQL & vbCrLf & "       REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                            'SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult1 & "', "                                          '���(�������)"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '�˻��ڵ�"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
                '/------------------------------------
                
                '/------------------------------------ ������̺� ��Ƽ�ڵ� ���� ������Ʈ
                If Trim(State_M) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H �� ������ : �˻����� ��������� �ִ´�
                        If asSend_gubun = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                            'SQL = SQL & vbCrLf & "       REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                            'SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult1 & "', "                                          '���(�������)"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf asSend_gubun = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�
                            'SQL = SQL & vbCrLf & "       REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                            'SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult1 & "', "                                          '���(�������)"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf asSend_gubun = "3" Then

                            SQL = SQL & vbCrLf & " SET   LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '����������"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
                            'SQL = SQL & vbCrLf & "       REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                            'SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult1 & "', "                                          '���(�������)"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_M) & "' "                                        '�˻��ڵ�"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            
            '/------------------------------------ ������̺� �׷��ڵ� ���� ������Ʈ
                If Trim(State_B) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    
                        '/////////  D/P/H �� ������ : �˻����� ��������� �ִ´�
                        If asSend_gubun = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf asSend_gubun = "2" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "_INF', "                                 '�߰�������"
                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '�߰������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        ElseIf asSend_gubun = "3" Then

                            SQL = SQL & vbCrLf & " SET   LAST_RPTR_ID = '" & gEquipCode & "_INF', "                                 '����������"
                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '���������Ͻ�"
                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
                        End If
                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_B) & "' "                                        '�˻��ڵ�"
                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                    res = SendQuery(gServer, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        cn_Ser.RollbackTrans
                        Exit Function
                    End If
                End If
            '/------------------------------------
            
            '/------------------------------------ �������̺� STATE ������Ʈ
                '////////// ���� ���̺�
                SQL = "UPDATE SPSLMJBDI "
                If asSend_gubun = "1" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf asSend_gubun = "2" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                ElseIf asSend_gubun = "3" Then
                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "','" & Trim(State_B) & "', '" & Trim(GetText(.vasTemp, iRow, 1)) & "') "                    '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
                res = SendQuery(gServer, SQL)
        
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If

            '/------------------------------------
                
                
            End If
        Next iRow
        
        If asSend_gubun = "" Or asSend_gubun = "0" Then cn_Ser.RollbackTrans: Exit Function
        
        '/------------------------------------ ó�����̺� STATE ������Ʈ
        '///////// ó�����̺�
        SQL = "UPDATE SPSLMJBBI "
        If asSend_gubun = "1" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf asSend_gubun = "2" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        ElseIf asSend_gubun = "3" Then
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        End If
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '3' "
        res = SendQuery(gServer, SQL)

        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------
        If asSend_gubun = "2" Then
        '/------------------------------------ ó�����̺� ������Ʈ(MDMDHTORD)
                       SQL = "UPDATE MDMDHTORD "
        SQL = SQL & vbCrLf & "   SET PRSC_STAT = '50'"      '/50 ���񺸰�, 51 ��������
        SQL = SQL & vbCrLf & "     , RPTG_DT = SYSDATE"
        SQL = SQL & vbCrLf & "     , AMEN_ID = 'POCT'"
        'SQL = SQL & vbCrLf & "  FROM MDMDHTORD"
        SQL = SQL & vbCrLf & " WHERE (PRSC_SQNO, PRSC_CD) "
        SQL = SQL & vbCrLf & "       IN (SELECT PRSC_SQNO, EXMN_CD "
        SQL = SQL & vbCrLf & "             FROM SPSLMJBDI "
        SQL = SQL & vbCrLf & "            WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "              AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "','" & Trim(State_B) & "') "
        SQL = SQL & vbCrLf & "              AND SPCM_STAT = '2') "
        SQL = SQL & vbCrLf & "AND DC_DVSN = 'O' "
        SQL = SQL & vbCrLf & "AND PRSC_STAT > '50'"
        ElseIf asSend_gubun = "3" Then
        
                       SQL = "UPDATE MDMDHTORD "
        SQL = SQL & vbCrLf & "   SET PRSC_STAT = '51' "      '/50 ���񺸰�, 51 ��������
        SQL = SQL & vbCrLf & "     , RPTG_DT = SYSDATE"
        SQL = SQL & vbCrLf & "     , AMEN_ID = 'POCT'"
        'SQL = SQL & vbCrLf & "  FROM MDMDHTORD"
        SQL = SQL & vbCrLf & " WHERE (PRSC_SQNO, PRSC_CD) "
        SQL = SQL & vbCrLf & "       IN (SELECT PRSC_SQNO, EXMN_CD "
        SQL = SQL & vbCrLf & "             FROM SPSLMJBDI "
        SQL = SQL & vbCrLf & "            WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "              AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "','" & Trim(State_B) & "') "
        SQL = SQL & vbCrLf & "              AND SPCM_STAT = '2') "
        SQL = SQL & vbCrLf & "AND DC_DVSN = 'O' "
        SQL = SQL & vbCrLf & "AND PRSC_STAT > '51'"
        End If
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------ ó�����̺� ������Ʈ

        'db_Commit gServer
        'cn_Ser.RollbackTrans
        cn_Ser.CommitTrans
        Insert_Data = 1
    End With
End Function

Function Insert_Data_R(ByVal argSpcRow As Integer) As Integer
    Dim iRow                As Integer
    Dim i                   As Integer
    Dim j                   As Integer
    Dim lsID                As String
    Dim lsSpecNo            As String
    Dim lsPid               As String
    Dim sResult             As String
    Dim sCnt                As String
    Dim sResult1            As String
    Dim sResult2            As String
    Dim ExamCnt             As String
    Dim ExamCode_Spec       As String
    Dim ExamCode_Remark     As String
    

    With frmInterface
        Insert_Data_R = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        lsID = ""
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))

        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpExamdate.value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
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

        gHIVPosFlag = -1
        
            
        SQL = "SELECT EXMN_CD "
        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
        res = db_select_Col(gServer, SQL)
        
        j = 0
        Do While gReadBuf(j) <> ""
            If ExamCode_Remark <> "" Then
                ExamCode_Remark = ExamCode_Remark & ",'" & gReadBuf(j) & "'"
            Else
                ExamCode_Remark = "'" & gReadBuf(j) & "'"
            End If
            j = j + 1
        Loop
        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
        Next i
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
                
                Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
                
                
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '���(�������)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL üũ"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta üũ"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic üũ"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '�߰�������"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '�߰������Ͻ�"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '����������"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '�������" (1:�Է� , 2:�߰�����, 3:��������)
                
                If gComment_All <> "" Or gComment_Code <> "" Then
                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
                End If
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
    
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        
        
        '//// ������̺��� �׷��ڵ带 ������ ����� ���� �ִ°�� ó��/���� ���̺� ������Ʈ ����
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        
        ExamCnt = gReadBuf(0)
        gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// ������̺� ����� �� �� �ִ� ��� (�׷��ڵ�����) ������Ʈ
            
            '///////// ó�����̺�
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// ���� ���̺�
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '�˻��ڵ�"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
            
        ElseIf res = -1 Then                                                             '///// ���� �����ΰ��
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        Else                                                                             '///// ����� ���Է��϶��� ������Ʈ ����
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_R = 1
    End With
End Function

Function Insert_Data_PhD(ByVal argSpcRow As Integer) As Integer
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

    With frmInterface
        Insert_Data_PhD = -1

        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))

        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpToday.value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp)

        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If

        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        'db_BeginTran gServer
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            If sResult1 = "" Then
                sResult1 = Trim(GetText(.vasTemp, iRow, 3))
            ElseIf sResult1 <> "" And sResult2 = "" Then
                sResult2 = Trim(GetText(.vasTemp, iRow, 3))
                
                If IsNumeric(sResult1) = True Then
                    sResult = sResult2 & "(" & sResult1 & ")"
                ElseIf IsNumeric(sResult2) = True Then
                    sResult = sResult1 & "(" & sResult2 & ")"
                End If
                
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '�������"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
    
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult & "', "                                           '���(�����)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult & "', "                                           '���(�������)"
                SQL = SQL & vbCrLf & "       DLTA_YN = '', "                                                            'Delta üũ"
                SQL = SQL & vbCrLf & "       PANC_YN = '', "                                                            'Panic üũ"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "', "                                     '����Է���"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '�߰�������"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '�߰������Ͻ�"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '����������"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "                                                        '���������
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '�������" (1:�Է� , 2:�߰�����, 3:��������)
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                    
            End If

        Next iRow

        '//// ������̺��� �׷��ڵ带 ������ ����� ���� �ִ°�� ó��/���� ���̺� ������Ʈ ����
        SQL = "SELECT EXMN_CD FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
        SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
        res = db_select_Vas(gServer, SQL, .vasTemp1)

        If res = 0 Then                                                                 '///// ������̺� ����� �� �� �ִ� ��� (�׷��ڵ�����)
            
            '/////
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If

            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        ElseIf res = -1 Then                                                             '///// ���� �����ΰ��
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        Else
                    
            '/////
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If

            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_PhD = 1
    End With
End Function

Function DELETE_LOCAL_ONE(asBarcode As String, asExamdate As String)
    
    SQL = ""
    SQL = SQL & vbCrLf & "DELETE FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EXAMDATE = '" & asExamdate & "' "
    SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
    SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(asBarcode) & "' "
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function Insert_Data_R_PhD(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'������ ����Ÿ ���̽��� ����
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

    Insert_Data_R_PhD = -1
    With frmInterface
        lsID = ""
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))
        
        
        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp
        
        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpExamdate.value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
        
        gHIVPosFlag = -1
        

        
        'db_BeginTran gServer
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            If sResult1 = "" Then
                sResult1 = Trim(GetText(.vasTemp, iRow, 3))
            ElseIf sResult1 <> "" And sResult2 = "" Then
                sResult2 = Trim(GetText(.vasTemp, iRow, 3))
                
                If IsNumeric(sResult1) = True Then
                    sResult = sResult2 & "(" & sResult1 & ")"
                ElseIf IsNumeric(sResult2) = True Then
                    sResult = sResult1 & "(" & sResult2 & ")"
                End If
                
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "                                                          '�������"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
    
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult & "', "                                           '���(�����)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult & "', "                                           '���(�������)"
                SQL = SQL & vbCrLf & "       DLTA_YN = '', "                                                            'Delta üũ"
                SQL = SQL & vbCrLf & "       PANC_YN = '', "                                                            'Panic üũ"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '�߰�������"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '�߰������Ͻ�"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '����������"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '�������" (1:�Է� , 2:�߰�����, 3:��������)
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                    
            End If

        Next iRow
        
        
        
        
        SQL = "SELECT EXMN_CD FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
        SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NULL "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        
        If res = 0 Then                                                                 '///// ������̺� ����� �� �� �ִ� ��� (�׷��ڵ�����)
            SQL = "Update SPSLMJBBI"
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT < 2 "
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        
            SQL = "Update SPSLMJBDI"
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
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
        Else
            SQL = "Update SPSLMJBBI"
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT > 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        
            SQL = "Update SPSLMJBDI"
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT > 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        End If
        
        SQL = ""
    
           
        db_Commit gServer
        Insert_Data_R_PhD = 1
    End With
End Function


Function Insert_Data_QC(ByVal argSpcRow As Integer) As Integer
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
    Dim lsQC_Date       As String

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
              " And examdate = '" & Format(CDate(.dtpExamdate.value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
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
                    SQL = SQL & vbCrLf & "      RSLT_RPTR_ID = '" & gEquipCode & "_INF' "                                                           'Delta üũ"
                    SQL = SQL & vbCrLf & "      AMEN_ID = '" & gEquipCode & "_INF' "                                                           'Panic üũ"
                    SQL = SQL & vbCrLf & "      UPDT_DT = sysdate, "                                     '����Է���"
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
                    SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                    res = db_select_Col(gServer, SQL)
                
                    If gReadBuf(0) = "" Then
                        sCnt = "1"
                    Else
                        sCnt = CLng(gReadBuf(0)) + 1
                    End If
                    
                        If Trim(GetText(.vasTemp, iRow, 2)) <> "" Then
                            SQL = ""
                            SQL = SQL & vbCrLf & "INSERT INTO SPSLHQRST(EXMN_DY   ,EQPM_CD ,SBSN_CD ,LVL_CD  "
                            SQL = SQL & vbCrLf & "                     ,RSLT_SQNO ,EXMN_CD ,RSLT_DT ,RSLT_RPTR_ID "
                            SQL = SQL & vbCrLf & "                     ,RSLT_VALU ,SPCM_NO ,DEL_YN "
                            SQL = SQL & vbCrLf & "                     ,REGI_ID   ,RGST_DT ,AMEN_ID ,UPDT_DT) "
                            SQL = SQL & vbCrLf & "               VALUES('" & Trim(GetText(.vasTemp, iRow, 9)) & "', '" & Mid(lsID, 3, 3) & "', '" & Mid(lsID, 6, 3) & "', '" & Mid(lsID, 9, 1) & "', "
                            SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gEquipCode & "_INF', "
                            'SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gEquipCode & "_INF', "
                            SQL = SQL & vbCrLf & "                      '" & sResult1 & "', '" & lsID & "', 'N', "
                            SQL = SQL & vbCrLf & "                      '" & gEquipCode & "_INF', sysdate, '" & gEquipCode & "_INF', sysdate ) "
                            res = SendQuery(gServer, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                cn_Ser.RollbackTrans
                                Exit Function
                            End If
                            
                        End If
                        
                    End If
                    
                End If
            
        Next iRow
        
        cn_Ser.CommitTrans
        Insert_Data_QC = 1
    End With
End Function

Function Insert_Data_ABI7500(ByVal argSpcRow As Integer) As Integer
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
    Dim Remark_Result   As String
    

    With frmInterface
        Insert_Data_ABI7500 = -1
        ExamCode_Spec = ""
        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
    
        If lsSpecNo = "" Then: Insert_Data_ABI7500 = -1: Exit Function
        
        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpToday.value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
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
            
            If Trim(GetText(.vasTemp, i, 1)) = "HLA-B27" Then: Remark_Result = Trim(GetText(.vasTemp, argSpcRow, 3))

        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        'db_BeginTran gServer
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '�������"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
    
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '���(�������)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL üũ"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Delta üũ"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Panic üũ"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '�߰�������"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '�߰������Ͻ�"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '����������"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '�������" (1:�Է� , 2:�߰�����, 3:��������)
                
                If Trim(GetText(.vasTemp, iRow, 1)) = "HLA-B51" Then
                    SQL = SQL & vbCrLf & ",       EXMN_PER_OPNN = 'HLA-B27 = " & Remark_Result & "' "                                                          'Remark �Է�
                    Remark_Result = ""
                End If
                
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
    
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        
        
        '//// ������̺��� �׷��ڵ带 ������ ����� ���� �ִ°�� ó��/���� ���̺� ������Ʈ ����
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        
        ExamCnt = gReadBuf(0)
        gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// ������̺� ����� �� �� �ִ� ��� (�׷��ڵ�����) ������Ʈ
            
            '///////// ó�����̺�
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// ���� ���̺�
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '�˻��ڵ�"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
            
        ElseIf res = -1 Then                                                             '///// ���� �����ΰ��
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        Else                                                                             '///// ����� ���Է��϶��� ������Ʈ ����
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_ABI7500 = 1
    End With
End Function

Function Insert_Data_ABI7500_R(ByVal argSpcRow As Integer) As Integer
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
    Dim Remark_Result   As String

    With frmInterface
        Insert_Data_ABI7500_R = -1
        ExamCode_Spec = ""
        lsID = ""
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))
    
        If lsSpecNo = "" Then: Insert_Data_ABI7500_R = -1: Exit Function
        
        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpToday.value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
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
            
            If Trim(GetText(.vasTemp, argSpcRow, 1)) = "HLA-B27" Then: Remark_Result = Trim(GetText(.vasTemp, argSpcRow, 3))

        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        gHIVPosFlag = -1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        'db_BeginTran gServer
        cn_Ser.BeginTrans
        '������ ����� �����ϱ�
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '�������"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1
    
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '���(�����)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '���(�������)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL üũ"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Delta üũ"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Panic üũ"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "_INF', "                                 '����Է���"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '����Է��Ͻ�"
                'SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                    '�߰�������"
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                 '�߰������Ͻ�"
                'SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gEquipCode & "', "                                    '����������"
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                 '���������Ͻ�"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '����ڵ�
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "                                      '���������
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '��������Ͻ�
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '�����ȣ (��� �����ÿ� ����)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '�������" (1:�Է� , 2:�߰�����, 3:��������)
                
                If Trim(GetText(.vasTemp, iRow, 1)) = "HLA-B51" Then
                    SQL = SQL & vbCrLf & ",       EXMN_PER_OPNN = '" & Remark_Result & "' "                                                          'Remark �Է�
                    Remark_Result = ""
                End If
                
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '��ü��ȣ"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    'ȯ�ڹ�ȣ"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '�������"
                
                res = SendQuery(gServer, SQL)
                If res < 0 Then
                    SaveQuery SQL
                   ' db_RollBack gServer
                   cn_Ser.RollbackTrans
                    Exit Function
                End If
                
                SQL = "UPDATE SPSLMJBDI "
                SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
                'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
                'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '�˻��ڵ�"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
                res = SendQuery(gServer, SQL)
    
                If res = -1 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        Next iRow
        
        
        
        '//// ������̺��� �׷��ڵ带 ������ ����� ���� �ִ°�� ó��/���� ���̺� ������Ʈ ����
        SQL = "SELECT COUNT(EXMN_CD) FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
        SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
        SQL = SQL & vbCrLf & "   AND (VIEW_RSLT IS NULL OR VIEW_RSLT = '') "
        res = db_select_Vas(gServer, SQL, .vasTemp1)
        
        ExamCnt = gReadBuf(0)
        gReadBuf(0) = "0"
        
        If ExamCnt = "0" Then                                                         '///// ������̺� ����� �� �� �ִ� ��� (�׷��ڵ�����) ������Ʈ
            
            '///////// ó�����̺�
            SQL = "UPDATE SPSLMJBBI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            '////////// ���� ���̺�
            SQL = "UPDATE SPSLMJBDI "
            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
            'SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "_INF', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & ExamCode_Spec & ") "                     '�˻��ڵ�"
            SQL = SQL & vbCrLf & "   AND EXMN_CD LIKE '%G%' "
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < 2 "
            SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
            res = SendQuery(gServer, SQL)

            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
            
        ElseIf res = -1 Then                                                             '///// ���� �����ΰ��
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        Else                                                                             '///// ����� ���Է��϶��� ������Ʈ ����
        
        End If

        SQL = ""


        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data_ABI7500_R = 1
    End With
End Function

Function Insert_Data_QC_R(ByVal argSpcRow As Integer) As Integer
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
    Dim lsQC_Date       As String

    With frmInterface
        Insert_Data_QC_R = -1
        ExamCode_Spec = ""
        lsID = ""
        lsID = Trim(GetText(.vasRID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasRID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasRID, argSpcRow, colPID))
        
        lsQC_Date = Format(GetDateFull, "yyyymmdd")

        'Local���� ȯ�ں��� ����� ��������
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, RESDATE, EXAMDATE" & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpExamdate.value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasRID, argSpcRow, colPos)) & "' "
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
                    SQL = SQL & vbCrLf & "      RSLT_RPTR_ID = '" & gEquipCode & "_INF' "                                                           'Delta üũ"
                    SQL = SQL & vbCrLf & "      AMEN_ID = '" & gEquipCode & "_INF' "                                                           'Panic üũ"
                    SQL = SQL & vbCrLf & "      UPDT_DT = sysdate, "                                     '����Է���"
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
                    SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                    res = db_select_Col(gServer, SQL)
                
                    If gReadBuf(0) = "" Then
                        sCnt = "1"
                    Else
                        sCnt = CLng(gReadBuf(0)) + 1
                    End If
                    
                        If Trim(GetText(.vasTemp, iRow, 2)) <> "" Then
                            SQL = ""
                            SQL = SQL & vbCrLf & "INSERT INTO SPSLHQRST(EXMN_DY   ,EQPM_CD ,SBSN_CD ,LVL_CD  "
                            SQL = SQL & vbCrLf & "                     ,RSLT_SQNO ,EXMN_CD ,RSLT_DT ,RSLT_RPTR_ID "
                            SQL = SQL & vbCrLf & "                     ,RSLT_VALU ,SPCM_NO ,DEL_YN "
                            SQL = SQL & vbCrLf & "                     ,REGI_ID   ,RGST_DT ,AMEN_ID ,UPDT_DT) "
                            SQL = SQL & vbCrLf & "               VALUES('" & Trim(GetText(.vasTemp, iRow, 9)) & "', '" & Mid(lsID, 3, 3) & "', '" & Mid(lsID, 6, 3) & "', '" & Mid(lsID, 9, 1) & "', "
                            SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gEquipCode & "_INF', "
                            'SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gEquipCode & "_INF', "
                            SQL = SQL & vbCrLf & "                      '" & sResult1 & "', '" & lsID & "', 'N', "
                            SQL = SQL & vbCrLf & "                      '" & gEquipCode & "_INF', sysdate, '" & gEquipCode & "_INF', sysdate ) "
                            res = SendQuery(gServer, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                cn_Ser.RollbackTrans
                                Exit Function
                            End If
                            
                        End If
                        
                    End If
                    
                End If
            
        Next iRow
        
        cn_Ser.CommitTrans
        Insert_Data_QC_R = 1
    End With
End Function

Function Save_ResMemo(ByVal asRow As Long, asMessage As String)
'�޽��� �����ϱ�
    Dim sMessage As String
    
    If asMessage = "" Then
        Exit Function
    End If
    
    sMessage = ""
    
'    SQL = "SELECT MESSAGE "
'    SQL = SQL & vbCrLf & " FROM PAT_RESMEMO  "
'    SQL = SQL & vbCrLf & "WHERE EQUIPNO = '" & gEquip & "' "
'    SQL = SQL & vbCrLf & "  AND BARCODE = '" & Trim(GetText(vasID, asRow, colBarcode)) & "' "
'    SQL = SQL & vbCrLf & "  AND EXAMDATE = '" & Format(dtpToday.Text, "yyyymmdd") & "' "
'    res = db_select_Col(gLocal, SQL)
'
'    sMessage = Trim(gReadBuf(0))
    
'    If sMessage = "" Then
        SQL = "INSERT INTO PAT_RESMEMO "
        SQL = SQL & vbCrLf & "     (EXAMDATE, EQUIPNO, BARCODE, MESSAGE )"
        SQL = SQL & vbCrLf & "VALUES('" & Format(frmInterface.dtpToday, "yyyymmdd") & "', "
        SQL = SQL & vbCrLf & "      '" & gEquip & "',"
        SQL = SQL & vbCrLf & "      '" & Trim(GetText(frmInterface.vasID, asRow, colBarcode)) & "', "
        SQL = SQL & vbCrLf & "      '" & asMessage & "') "
'    Else
'        'sMessage = sMessage & vbCrLf & asMessage
'        sMessage = sMessage & ", " & asMessage

'        SQL = " Update pat_resmemo Set " & vbCrLf & _
'              " message = '" & Trim(sMessage) & "' " & vbCrLf & _
'              " Where examdate = '" & Format(Text_Today.Text, "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & Trim(GetText(vasID, asRow, colBarcode)) & "' "
'    End If
    
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function Pat_Info(asBarcode As String) As Integer
    Dim sBarcode As String
    Dim sSpecNo As String

    Pat_Info = -1
    With frmInterface
        'ȯ������ ��������
        If asBarcode = "" Or IsNumeric(asBarcode) = False Then
            Exit Function
        End If
        '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
        
        SQL = "SELECT FN_LABCVTBCNO(" & Trim(asBarcode) & ") FROM DUAL "
        res = db_select_Col(gServer, SQL)
        
        sSpecNo = Trim(gReadBuf(0))
        
        'ȯ�ڹ�ȣ, ȯ���̸�, �ֹι�ȣ, ����, ����
        SQL = "SELECT PID, PT_NM, SEX, AGE "
        SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
        SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
        SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
        'SQL = SQL & vbCrLf & "  AND RSLT_STAT < 2 "
        res = db_select_Col(gServer, SQL)
        
        '///////// gAllExam �ڸ��� �˻� �ڵ� �־��� �����ڵ� �� �پ� �ִ°� B312001 , 02, 03
        
        If res = 1 Then
            SetText .vasList, Trim(sSpecNo), 1, colSpecNo
            SetText .vasList, Trim(gReadBuf(0)), 1, colPID
            SetText .vasList, Trim(gReadBuf(1)), 1, colPName
            SetText .vasList, Trim(gReadBuf(2)), 1, colSex
            SetText .vasList, Trim(gReadBuf(3)), 1, colAge
            
            Pat_Info = 1
        Else
        
            Pat_Info = -1
            SaveQuery (SQL)
        End If
    End With
End Function

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim sBarcode As String
    Dim sSpecNo As String

    Get_Sample_Info = -1
    With frmInterface
        'ȯ������ ��������
        sBarcode = Trim(GetText(.vasWorkList, asRow, colBarcode))   '���� ���ڵ� ��ȣ
        If sBarcode = "" Or IsNumeric(sBarcode) = False Then
            Exit Function
        End If
        '���ڵ��ȣ�� ��ü��ȣ �ҷ�����
        
        SQL = "SELECT FN_LABCVTBCNO(" & Trim(sBarcode) & ") FROM DUAL "
        res = db_select_Col(gServer, SQL)
        
        sSpecNo = Trim(gReadBuf(0))
        
        'ȯ�ڹ�ȣ, ȯ���̸�, �ֹι�ȣ, ����, ����
        SQL = "SELECT A.PID, A.PT_NM, A.SEX, A.AGE, B.WORK_NO "
        SQL = SQL & vbCrLf & " FROM SPSLMJBBI A, SPSLHRRST B "
        SQL = SQL & vbCrLf & "WHERE A.SPCM_NO = B.SPCM_NO"
        SQL = SQL & vbCrLf & "  AND A.SPCM_NO = '" & sSpecNo & "' "
        SQL = SQL & vbCrLf & "  AND A.SPCM_STAT = '2' "
        'SQL = SQL & vbCrLf & "  AND RSLT_STAT < '2' "
        res = db_select_Col(gServer, SQL)
        
        '///////// gAllExam �ڸ��� �˻� �ڵ� �־��� �����ڵ� �� �پ� �ִ°� B312001 , 02, 03
        
        If res = 1 Then
            SetText .vasWorkList, Trim(gReadBuf(4)), asRow, colSpecNo
'            SetText .vasWorklist, Trim(gReadBuf(0)), asRow, colPID
'            SetText .vasWorklist, Trim(gReadBuf(1)), asRow, colPName
'            SetText .vasWorklist, Trim(gReadBuf(2)), asRow, colSex
'            SetText .vasWorklist, Trim(gReadBuf(3)), asRow, colAge
            
            Get_Sample_Info = 1
        Else
        
            Get_Sample_Info = -1
            SaveQuery (SQL)
        End If
    End With
End Function

Function Get_Sample_Info_QC(ByVal asRow As Long) As Integer
    Dim sBarcode As String
    Dim sQCdate  As String
    
    Get_Sample_Info_QC = -1
    With frmInterface
        'ȯ������ ��������
        sBarcode = Trim(GetText(.vasID, asRow, colBarcode))   '���� ���ڵ� ��ȣ
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
            SetText .vasID, Trim(sBarcode), asRow, colSpecNo
            SetText .vasID, Trim(gReadBuf(0)), asRow, colPID
            SetText .vasID, Trim(gReadBuf(1)), asRow, colPName
            SetText .vasID, Trim(gReadBuf(2)), asRow, colSex
            SetText .vasID, Trim(gReadBuf(3)), asRow, colAge
            
            SetText .vasList, Trim(sBarcode), 1, colSpecNo
            SetText .vasList, Trim(gReadBuf(0)), 1, colPID
            SetText .vasList, Trim(gReadBuf(1)), 1, colPName
            SetText .vasList, Trim(gReadBuf(2)), 1, colSex
            SetText .vasList, Trim(gReadBuf(3)), 1, colAge
            
            Get_Sample_Info_QC = 1
        Else
        
            Get_Sample_Info_QC = -1
            SaveQuery (SQL)
        End If
    End With
End Function

Function Get_Sample_InfoR(ByVal asRow As Long) As Integer
   Dim sBarcode As String
    Dim sSpecNo As String
    With frmInterface
        Get_Sample_InfoR = -1
        'ȯ������ ��������
        sBarcode = Trim(GetText(.vasRID, asRow, colBarcode))   '���� ���ڵ� ��ȣ
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
            SetText .vasRID, Trim(sSpecNo), asRow, colSpecNo
            SetText .vasRID, Trim(gReadBuf(0)), asRow, colPID
            SetText .vasRID, Trim(gReadBuf(1)), asRow, colPName
            SetText .vasRID, Trim(gReadBuf(2)), asRow, colSex
            SetText .vasRID, Trim(gReadBuf(3)), asRow, colAge
            
            Get_Sample_InfoR = 1
        Else
        
            Get_Sample_InfoR = -1
        End If
    End With
End Function

Function EquipExamCode(asEquipCode As String, asBarcode As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim SpecNo As String


    EquipExamCode = ""
    
    If Trim(asEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(asEquipCode) & "' "
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
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT FN_LABCVTBCNO('" & Trim(asBarcode) & "') FROM DUAL "
    res = db_select_Col(gServer, SQL)
    SpecNo = gReadBuf(0)
    
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(SpecNo) & "' "
    SQL = SQL & vbCrLf & "  AND A.EXMN_CD IN (" & sExamCode & ") "
    res = db_select_Col(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        gEquipExamCode = Trim(gReadBuf(0))
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT RSLT_SMNO_SIZE FROM SPSLMFBIF"
        SQL = SQL & vbCrLf & " WHERE EXMN_CD = '" & gEquipExamCode & "' "
        SQL = SQL & vbCrLf & "   AND USE_END_DY > sysdate "
        res = db_select_Col(gServer, SQL)
        gExamRange = gReadBuf(0)
    End If
    
End Function

Function EquipExamCode_QC(asEquipCode As String, asBarcode As String) As String
'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Dim i As Integer
Dim sExamCode As String
Dim SpecNo As String


    EquipExamCode_QC = ""
    
    If Trim(asEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(asEquipCode) & "' "
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
    
    
    SQL = ""
    SQL = "SELECT QC_EXMN_CD "
    SQL = SQL & vbCrLf & " FROM SPSLMQMST "
    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(asBarcode, 3, 3) & "' "
    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(asBarcode, 6, 3) & "' "
    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(asBarcode, 9, 1) & "' "
    SQL = SQL & vbCrLf & "  AND QC_EXMN_CD IN (" & sExamCode & ") "

    res = db_select_Col(gServer, SQL)

  
    If gReadBuf(0) <> "" Then
        gEquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function

Function TLA_LASC_Search(asStartDate As String, asEndDate As String)
    Dim Server_date As String
    Dim buff As String
    Dim StartDate As String
    Dim EndDate As String
    
    buff = "0.7"
    Server_date = Trim(Format(GetDateFull, "yyyy/mm/dd"))
    StartDate = DateDiff("d", Server_date, asStartDate)
    EndDate = DateDiff("d", Server_date, asEndDate)
    
    If InStr(StartDate, "-") > 0 Then: StartDate = CCur(StartDate) * -1
    If InStr(EndDate, "-") > 0 Then: EndDate = CCur(EndDate) * -1
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT '',  C.SPCM_NO, (SELECT FN_LABCVTPRTBCNO(C.SPCM_NO) FROM DUAL), C.SPCM_SQNO, substr(max(B.WORK_NO), -4),C.PID, C.PT_NM, C.SEX, C.AGE "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B, SPSLMJBBI C "
    SQL = SQL & vbCrLf & " WHERE (C.STAT_DVSN = '' OR C.STAT_DVSN IS NULL) "
    SQL = SQL & vbCrLf & "   AND B.RGST_DT BETWEEN SYSDATE - " & (CLng(StartDate) + CCur(buff))
    SQL = SQL & vbCrLf & "                                     AND SYSDATE - " & CLng(EndDate)
    SQL = SQL & vbCrLf & "   AND C.SPCM_NO = A.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND C.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = C.SPCM_STAT "
    SQL = SQL & vbCrLf & "   AND C.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND C.RSLT_STAT = A.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND C.RSLT_STAT = '0' "
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & " GROUP BY C.SPCM_NO, C.PID, C.PT_NM, C.SEX, C.AGE, C.SPCM_SQNO "
  
    res = db_select_Vas(gServer, SQL, frmInterface.vasID, frmInterface.vasID.DataRowCnt + 1)
    
End Function


Function PAT_List_Search(asStartDate As String, asEndDate As String)
    Dim Server_date As String
    
    Dim buff As String
    Dim StartDate As String
    Dim EndDate As String
    
    buff = "0.7"
    Server_date = Trim(Format(GetDateFull, "yyyy/mm/dd"))
    StartDate = DateDiff("d", Server_date, asStartDate)
    EndDate = DateDiff("d", Server_date, asEndDate)
    
    If InStr(StartDate, "-") > 0 Then: StartDate = CCur(StartDate) * -1
    If InStr(EndDate, "-") > 0 Then: EndDate = CCur(EndDate) * -1

    
With frmInterface
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT '',  B.SPCM_NO, (SELECT FN_LABCVTPRTBCNO(B.SPCM_NO) FROM DUAL),'','', C.PID ,C.PT_NM "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B, SPSLMJBBI C "
    
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND B.SPCM_NO = C.SPCM_NO"
    SQL = SQL & vbCrLf & "   AND B.RGST_DT BETWEEN SYSDATE - " & (CLng(StartDate) + CCur(buff))
    SQL = SQL & vbCrLf & "                                     AND SYSDATE - " & CLng(EndDate)
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    
    SQL = SQL & vbCrLf & "   AND C.SPCM_STAT = B.SPCM_STAT "
    SQL = SQL & vbCrLf & "   AND C.RSLT_STAT = B.RSLT_STAT "
    
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND C.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND C.RSLT_STAT = '0' "
    SQL = SQL & vbCrLf & " GROUP BY B.SPCM_NO, C.PID, C.PT_NM "
    
    res = db_select_Vas(gServer, SQL, .vasPatList)
    
    
    
    If res = 1 Then
    ElseIf res = -1 Then
        SaveQuery (SQL)
    End If

End With
End Function

Function LASC_Start_Server(ByVal argSpcRow As Integer) As Integer

'S000000009638527410     ********111001100000kim          gim          guim         000****************************************    <--- ���� �Ѱ�
'S00000000     1117559341********110000000000000000000000000000000000000000000000000000****************************************    <--- ���� �ϰ� �ִ°�
With frmInterface
    gEXAM_CBC = "N"
    gEXAM_Diff = "N"
    gEXAM_Reti = "N"
    gEXAM_CBC_Diff = "N"
    
    Call ClearSpread(.vasTemp1)
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam_CBC & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "
    res = db_select_Vas(gServer, SQL, .vasTemp1)
    
    If res > 0 Then: gEXAM_CBC = "Y"
    
    Call ClearSpread(.vasTemp1)
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam_Diff & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "
    res = db_select_Vas(gServer, SQL, .vasTemp1)
    
    If res > 0 Then: gEXAM_Diff = "Y"
    
    Call ClearSpread(.vasTemp1)
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam_Reti & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "
    res = db_select_Vas(gServer, SQL, .vasTemp1)
    
     If res > 0 Then: gEXAM_Reti = "Y"
    
    Call ClearSpread(.vasTemp1)
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam_CBC_Diff & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "
    res = db_select_Vas(gServer, SQL, .vasTemp1)
    
    If res > 0 Then: gEXAM_CBC_Diff = "Y"
    
    Call Lasc_Order_Make(Trim(GetText(.vasID, argSpcRow, colBarcode)), argSpcRow)
    
End With
End Function

Function Lasc_Order_Make(asBarcode As String, asRow As Integer)
'S000000009638527410     ********111001100000kim          gim          guim         000****************************************    <--- ���� �Ѱ�
'S00000000     1117559341********110000000000000000000000000000000000000000000000000000****************************************    <--- ���� �ϰ� �ִ°�
    Dim Order_Total         As String
    Dim Order_Head          As String
    Dim Order_Barcode       As String
    Dim Order_Order         As String
    Dim Order_OrderCBC      As String
    Dim Order_OrderDiff     As String
    Dim Order_OrderReti     As String
    Dim Order_OrderCBCDiff  As String
    Dim Order_Etc1 As String
    Dim Order_Etc2 As String
    
    '///// ���� �ʱ�ȭ
    Order_Total = ""
    Order_Head = ""
    Order_Barcode = ""
    Order_OrderCBC = "0"
    Order_OrderDiff = "0"
    Order_OrderReti = "0"
    Order_OrderCBCDiff = "0"
    Order_Etc1 = ""
    
    
    Order_Head = "S00000000"
    Order_Barcode = SetSpace(asBarcode, 15, 1)
    Order_Etc1 = "********"
    If gEXAM_CBC = "Y" Then: Order_OrderCBC = "1"
    If gEXAM_Diff = "Y" Then: Order_OrderDiff = "1"
    If gEXAM_Reti = "Y" Then: Order_OrderReti = "1"
    If gEXAM_CBC_Diff = "Y" Then: Order_OrderCBC = "1": Order_OrderDiff = "1"
    Order_Order = Order_OrderCBC & Order_OrderDiff & Order_OrderReti & "0000"
    
    Order_Etc2 = "00000000000000000000000000000000000000000000000****************************************"
    
    
    Order_Total = Order_Head & Order_Barcode & Order_Etc1 & Order_Order & Order_Etc2
    Order_Total = chrSTX & Order_Total & chrETX
    
    SetText frmInterface.vasOrder, Order_Total, frmInterface.vasOrder.DataRowCnt + 1, 1
    SetText frmInterface.vasOrder, CStr(asRow), frmInterface.vasOrder.DataRowCnt, 2

End Function

Function TLA_Start_Server(ByVal argSpcRow As Integer) As Integer
    Dim ExamCount As String
    Dim TLA_Equip As String
    Dim i As Integer
    
    '//////////////�˻�����ڵ� Count
    Dim EQ_DX1 As Integer
    Dim EQ_DX2 As Integer
    Dim EQ_DX3 As Integer
    Dim EQ_DXC As Integer
    Dim EQ_DX0 As Integer
    Dim EQ_D1C As Integer
    Dim EQ_D2C As Integer
    Dim EQ_D3C As Integer
    Dim EQ_D0C As Integer
    Dim EQ_CEN As Integer
    Dim EQ_IML As Integer
    Dim EQ_ELE As Integer
    Dim EQ_SER As Integer
    Dim EQ_COB As Integer
    Dim EQ_VST As Integer
    
    '////////////���� ����Ǻ�
    Dim EQ_NO As String
    Dim EQ_NO1 As String
    Dim EQ_NO2 As String
    Dim EQ_NO3 As String
    Dim EQ_NO_JA As String
    
    '/////////// TLA ���ü ����
    Dim TLA_MO As String
    '/////////// TLA �ڰ�ü ����
    Dim TLA_JA(0 To 6) As String
    '/////////// L8 �̸� WorkNo
    Dim A_W_No As String
    '/////////// ä�� �Ͻ�
    Dim lsRCPN_DT As String
    
With frmInterface
    TLA_Start_Server = -1
    
    EQ_DX1 = 0
    EQ_DX2 = 0
    EQ_DX3 = 0
    EQ_DXC = 0
    EQ_DX0 = 0
    EQ_D1C = 0
    EQ_D2C = 0
    EQ_D3C = 0
    EQ_D0C = 0
    EQ_CEN = 0
    EQ_IML = 0
    EQ_ELE = 0
    EQ_SER = 0
    EQ_COB = 0
    EQ_VST = 0
    A_W_No = ""
    
    Call ClearSpread(.vasTemp1)
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD  "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND B.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY A.EXMN_CD  "

    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
    
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT TO_CHAR(B.BLCL_DT,'YYYY/MM/DD'), C.SPCM_SQNO, substr(MAX(B.WORK_NO),-3)   "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A, SPSLMJBDI B, SPSLMJBBI C "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = C.SPCM_NO  "
    SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & Trim(GetText(.vasID, argSpcRow, colSpecNo)) & "'"
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND C.SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT = B.RSLT_STAT "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY B.BLCL_DT, C.SPCM_SQNO "
    
    res = db_select_Col(gServer, SQL)
    
    If Trim(Mid(GetText(frmInterface.vasTemp1, 1, 1), 1, 2)) = "L8" Then
        A_W_No = "W" & Format(gReadBuf(2), "000#")
    Else
        A_W_No = "A" & Format(gReadBuf(1), "000#")
    End If
    
    lsRCPN_DT = Trim(gReadBuf(0))
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    For i = 1 To frmInterface.vasTemp1.DataRowCnt
        If TLA_Equip <> "" Then
            TLA_Equip = TLA_Equip & ",'" & Trim(GetText(.vasTemp1, i, 1)) & "'"
        Else
            TLA_Equip = "'" & Trim(GetText(.vasTemp1, i, 1)) & "'"
        End If
    Next i
    
    Call ClearSpread(.vasTemp1)
    
    If TLA_Equip = "" Then: TLA_Equip = "''"
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EQUIPCODE, B.EQUIPNUMBER "
    SQL = SQL & vbCrLf & "  FROM EQUIPEXAM A, DIVISION B "
    SQL = SQL & vbCrLf & " WHERE A.EQUIPCODE = B.EQUIPCODE "
    SQL = SQL & vbCrLf & "   AND A.EXAMCODE IN (" & TLA_Equip & ") "
'''    SQL = SQL & vbCrLf & "   AND B.EQUIPMAIN = 'Y'"
    SQL = SQL & vbCrLf & " GROUP BY A.EQUIPCODE, B.EQUIPNUMBER "
    res = db_select_Vas(gLocal, SQL, .vasTemp1)
    
    Dim cntEquipNum As Integer
    Dim blMainFlag As Boolean
    Dim cntEquipMain As Integer
    Dim blMoFlag As Boolean
    Dim blJaFlag As Boolean
    
    
    blMainFlag = False
    cntEquipNum = 0
    cntEquipMain = 0
    
    For i = 1 To res
        If IsNumeric(Trim(GetText(.vasTemp1, i, 2))) = True Then
            blMainFlag = True
            cntEquipNum = cntEquipNum + CCur(Trim(GetText(.vasTemp1, i, 2)))
            cntEquipMain = cntEquipMain + 1
            
        End If
    Next
    
    If blMainFlag = True Then '������� �˻簡 �ɸ��� ���
        If cntEquipMain = 2 And cntEquipNum = 7 Then
            EQ_NO = cntEquipNum - 1
        
        Else
            EQ_NO = cntEquipNum
        End If
        
    End If
    
    
    For i = 1 To .vasTemp1.DataRowCnt
        If InStr(1, GetText(.vasTemp1, i, 2), "-") > 0 Then
            If EQ_NO1 = "" Then
                EQ_NO1 = GetText(.vasTemp1, i, 2)
            ElseIf EQ_NO2 = "" Then
                EQ_NO2 = GetText(.vasTemp1, i, 2)
            ElseIf EQ_NO3 = "" Then
                EQ_NO3 = GetText(.vasTemp1, i, 2)
            End If
        
        Else
            If IsNumeric(Trim(GetText(.vasTemp1, i, 2))) = False Then
                If EQ_NO1 = "" Then
                    EQ_NO1 = GetText(.vasTemp1, i, 2)
                ElseIf EQ_NO2 = "" Or EQ_NO1 <> GetText(.vasTemp1, i, 2) Then
                    EQ_NO2 = GetText(.vasTemp1, i, 2)
                ElseIf EQ_NO3 = "" Then
                    EQ_NO3 = GetText(.vasTemp1, i, 2)
                End If
            End If
            
        End If
    Next i

    
    
'''    For i = 1 To .vasTemp1.DataRowCnt
'''        If IsNumeric(GetText(.vasTemp1, i, 2)) = True And InStr(1, GetText(.vasTemp1, i, 2), "-") = 0 Then
'''            If EQ_NO = "" Then
'''                EQ_NO = CCur(GetText(.vasTemp1, i, 2))
'''            Else
'''                EQ_NO = CCur(EQ_NO) + CCur(GetText(.vasTemp1, i, 2))
'''            End If
'''
'''        ElseIf IsNumeric(GetText(.vasTemp1, i, 2)) = True And InStr(1, GetText(.vasTemp1, i, 2), "-") > 0 Then
'''            If EQ_NO1 = "" Then
'''                EQ_NO1 = GetText(.vasTemp1, i, 2)
'''            ElseIf EQ_NO2 = "" Then
'''                EQ_NO2 = GetText(.vasTemp1, i, 2)
'''            ElseIf EQ_NO3 = "" Then
'''                EQ_NO3 = GetText(.vasTemp1, i, 2)
'''            End If
'''
'''        ElseIf IsNumeric(GetText(.vasTemp1, i, 2)) = False Then
'''            If EQ_NO1 = "" Then
'''                EQ_NO1 = GetText(.vasTemp1, i, 2)
'''            ElseIf EQ_NO2 = "" Or EQ_NO1 <> GetText(.vasTemp1, i, 2) Then
'''                EQ_NO2 = GetText(.vasTemp1, i, 2)
'''            ElseIf EQ_NO3 = "" Then
'''                EQ_NO3 = GetText(.vasTemp1, i, 2)
'''            End If
'''        End If
'''    Next i
    
    If EQ_NO1 <> "" Then
        EQ_NO_JA = "'" & EQ_NO1 & "'"

        If EQ_NO2 <> "" Then
            EQ_NO_JA = EQ_NO_JA & ", '" & EQ_NO2 & "'"

            If EQ_NO3 <> "" Then
                EQ_NO_JA = EQ_NO_JA & ", '" & EQ_NO3 & "'"
            End If

        End If

    End If
    
    ClearSpread .vasTemp1
    
    If EQ_NO_JA = "" Then: EQ_NO_JA = "''"
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA, EQUIPNUMBER "
    SQL = SQL & vbCrLf & "  FROM DIVISION "
    If EQ_NO <> "" Then
        SQL = SQL & vbCrLf & "   WHERE EQUIPNUMBER = '" & EQ_NO & "' "
    Else
        SQL = SQL & vbCrLf & "   WHERE EQUIPNUMBER IN (" & EQ_NO_JA & ") "
    End If
    SQL = SQL & vbCrLf & " GROUP BY EQUIPCODE_TLA, EQUIPNUMBER "
    res = db_select_Vas(gLocal, SQL, .vasTemp1)
    
    If EQ_NO <> "" Then
        TLA_MO = Trim(GetText(.vasTemp1, 1, 1))
        
    Else
        blMoFlag = False
        blJaFlag = False
        For i = 1 To res
            If Mid(Trim(GetText(.vasTemp1, i, 2)), 1, 1) = "-" Then
                TLA_MO = Trim(GetText(.vasTemp1, i, 1))
                blMoFlag = True
                Exit For
            End If
        Next
        
        If blMoFlag = False Then
            For i = 1 To res
                If IsNumeric(Trim(GetText(.vasTemp1, i, 2))) = False Then
                    TLA_MO = Trim(GetText(.vasTemp1, i, 1))
                    blMoFlag = True
                    blJaFlag = True
                    Exit For
                End If
            Next
        End If
        
    End If
    ClearSpread .vasTemp1
    
    gReadBuf(0) = ""
    
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA , COUNT(EQUIPCODE_TLA) "
    SQL = SQL & vbCrLf & "  FROM DIVISION  "
    If EQ_NO_JA <> "" Then
        SQL = SQL & vbCrLf & "   WHERE EQUIPNUMBER IN (" & EQ_NO_JA & ") "
        SQL = SQL & vbCrLf & "     AND USEYN = 'Y' "
        SQL = SQL & vbCrLf & "     AND EQUIPCODE_TLA <> '" & TLA_MO & "' "
    Else
        SQL = SQL & vbCrLf & "   WHERE EQUIPNUMBER = '����' "
    End If
        
    SQL = SQL & vbCrLf & " GROUP BY EQUIPCODE_TLA "
    res = db_select_Vas(gLocal, SQL, .vasTemp1)
    
    Dim Code_TLA As String
    '/////// ��ȸ�� �ڵ� ���� Ȯ���ؼ�
    If GetText(.vasTemp1, 1, 2) = "" Then
        Code_TLA = "0"
    Else
        Code_TLA = GetText(.vasTemp1, 1, 2)
    End If
    
    For i = 1 To CInt(res)
        TLA_JA(i - 1) = GetText(.vasTemp1, i, 1)
    Next i
    
    If TLA_JA(0) = "" And blMoFlag = True And blJaFlag = True Then: TLA_JA(0) = TLA_MO
''    i = 0
''    Do While gReadBuf(i) <> ""
''
''        TLA_JA(i) = gReadBuf(i)
''        gReadBuf(i) = ""
''
''        i = i + 1
''    Loop
''
''    i = 0
''    Do While TLA_JA(i) <> ""
''        If TLA_JA(i) = TLA_MO And TLA_JA(i + 1) <> "" Then
''            TLA_JA(i) = TLA_JA(i + 1)
''            TLA_JA(i + 1) = ""
''        ElseIf TLA_JA(i) = TLA_MO And TLA_JA(i + 1) = "" Then
''            TLA_JA(i) = TLA_JA(i + 1)
''            TLA_JA(i + 1) = ""
''        End If
''        i = i + 1
''    Loop
    
    If TLA_MO = "" Then
        TLA_MO = TLA_JA(0)
    End If
    
    
    res = TLA_Division(argSpcRow, TLA_MO, TLA_JA(0), TLA_JA(1), TLA_JA(2), A_W_No, lsRCPN_DT)
    If res = -1 Then
        Save_Raw_Data "[TLA] " & GetDateFull & ":   ��ü ���� ����"
        Exit Function
    End If
    
    TLA_Start_Server = 1

End With
End Function

Function TLA_Division(argSpcRow As Integer, asMO As String, _
                      asJA1 As String, asJA2 As String, asJA3 As String, _
                      asA_W_No As String, asRCPN_DT As String) As Integer
                      
    TLA_Division = -1
    Dim BarCodeNo As String
    Dim Age_Conv As String
    Dim i As Integer
    
    Dim Signal                      As String
    Dim Signal_Head                 As String
    Dim Signal_Barcode              As String
    Dim Signal_SpecNo               As String
    Dim Signal_Print                As String
    Dim Signal_UseDate              As String
    Dim Signal_Pname                As String
    Dim Signal_Age                  As String
    Dim Signal_Sex                  As String
    Dim Signal_ReceDate             As String
    Dim Signal_ReceNo               As String
    Dim Signal_Info                 As String
    Dim Signal_Mo                   As String
    Dim Signal_Mo_Bunju             As String
    Dim Signal_MO_Place             As String
    Dim Signal_JA_Bansong(0 To 6)   As String
    Dim Signal_JA_Bunju(0 To 6)     As String
    Dim Signal_JA_Rank(0 To 6)      As String
    Dim Signal_JA_EQName(0 To 6)    As String
    Dim JA_CODE(1 To 3)             As String
    Dim JA_VALUE(1 To 3)            As String
    
    
    Signal_Head = ""
    Signal_Barcode = ""
    Signal_SpecNo = ""
    Signal_Print = ""
    Signal_UseDate = ""
    Signal_Pname = ""
    Signal_Age = ""
    Signal_Sex = ""
    Signal_ReceDate = ""
    Signal_ReceNo = ""
    Signal_Info = ""
    Signal_Mo = ""
    Signal_Mo_Bunju = ""
    Signal_MO_Place = ""
    
    For i = 0 To 6
       Signal_JA_Bansong(i) = ""
       Signal_JA_Bunju(i) = ""
       Signal_JA_Rank(i) = ""
       Signal_JA_EQName(i) = ""
    Next i
    
    
    '//////// �ʱ�ȭ
    Signal = ""
    Signal_Head = ""
    Signal_Barcode = ""
    Signal_SpecNo = ""
    Signal_Print = ""
    Signal_UseDate = ""
    Signal_Pname = ""
    Signal_Age = ""
    Signal_Sex = ""
    Signal_ReceDate = ""
    Signal_ReceNo = ""
    Signal_Info = ""
    Signal_Mo = ""
    Signal_Mo_Bunju = ""
    Signal_MO_Place = ""
    
    '/////// ���ü, �ڰ�ü �ʱ�ȭ
    
    
    If asMO = "" Then: Exit Function
With frmInterface
    '//////// ��� ���� �ֱ�
    BarCodeNo = Trim(GetText(.vasID, argSpcRow, colBarcode))
    
    Signal_Head = "IC"
    Signal_Barcode = SetSpace(Trim(GetText(.vasID, argSpcRow, colBarcode)), 14, 2)
    Signal_SpecNo = SetSpace(Trim(GetText(.vasID, argSpcRow, colBarcode)), 14, 2)
    
    Signal_Print = SetSpace(Format(Mid(asA_W_No, 2), "0000"), 4, 1) & "/"
    Signal_Print = Signal_Print & Trim(GetText(.vasID, argSpcRow, colPID)) & "/"
    Signal_Print = Signal_Print & "    " & "/"                                          '////// ������Ʈ ��ȸ�ؾ���
    Signal_Print = Signal_Print & Trim(GetText(.vasID, argSpcRow, colSex)) & ""
    Signal_Print = SetSpace(Signal_Print, 32, 2)
    
    Signal_UseDate = SetSpace(asRCPN_DT, 10)
    Signal_Pname = SetSpace_1(Trim(GetText(.vasID, argSpcRow, colPName)), 14, 2)
    
    
    If IsNumeric(Trim(GetText(.vasID, argSpcRow, colAge))) = True Then
        Signal_Age = CStr(Trim(CCur(Format(Date, "yyyy"))) - CCur(Trim(GetText(.vasID, argSpcRow, colAge))) - 1) & "/01"
    Else
        Age_Conv = Mid(Trim(GetText(.vasID, argSpcRow, colAge)), 1, 2)
        If IsNumeric(Age_Conv) = False Then
            Age_Conv = Mid(Trim(GetText(.vasID, argSpcRow, colAge)), 1, 1)
        End If
        
        If Age_Conv > 11 Then
            Age_Conv = 1
        Else
            Age_Conv = 2
        End If
        
        Signal_Age = CStr(Trim(CCur(Format(Date, "yyyy"))) - CCur(Age_Conv - 1)) & "/01"
    End If
    Signal_Age = SetSpace(Signal_Age, 7)
    
    Signal_Sex = Trim(GetText(.vasID, argSpcRow, colSex))
    Signal_ReceDate = Format(.dtpToday.value, "yyyy/mm/dd")
    Signal_ReceDate = Mid(Signal_ReceDate, 1, 4) & "/" & Mid(Signal_ReceDate, 6, 2) & "/" & Mid(Signal_ReceDate, 9, 2)
    Signal_ReceNo = "    "
    Signal_Info = "        " 'Mid(asA_W_No, 2)
    Signal_Mo = asMO
    
    If asJA1 <> "" And asJA2 = "" And asJA3 = "" Then
        Signal_Mo_Bunju = "1"
    ElseIf asJA1 <> "" And asJA2 <> "" And asJA3 = "" Then
        Signal_Mo_Bunju = "2"
    ElseIf asJA1 <> "" And asJA2 <> "" And asJA3 <> "" Then
        Signal_Mo_Bunju = "3"
    ElseIf asJA1 = "" And asJA2 = "" And asJA3 = "" Then
        Signal_Mo_Bunju = "0"
    End If
    
    
    Signal_MO_Place = asA_W_No
    
    Signal = Signal_Head & Signal_Barcode & Signal_SpecNo & Signal_Print & Signal_UseDate & Signal_Pname & _
             Signal_Age & Signal_Sex & Signal_ReceDate & Signal_ReceNo & Signal_Info & Signal_Mo & Signal_Mo_Bunju & Signal_MO_Place
    
    If asJA1 = "" Then
        
    ElseIf asJA1 <> "" And asJA2 = "" And asJA3 = "" Then
        ClearSpread .vasTemp1
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA, JA_VALUES"
        SQL = SQL & vbCrLf & "  FROM Division "
        SQL = SQL & vbCrLf & " WHERE EQUIPCODE_TLA = '" & asJA1 & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp1)
        
        For i = 1 To res
            If Trim(GetText(.vasTemp1, i, 2)) <> "" And Trim(GetText(.vasTemp1, i, 2)) <> "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = Format(Trim(GetText(.vasTemp1, i, 2)), "0000")
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            End If
        Next i
        
        Signal = Signal & "    " & JA_VALUE(1) & "     " & JA_CODE(1)
        ClearSpread .vasTemp1
        
    ElseIf asJA1 <> "" And asJA2 <> "" And asJA3 = "" Then
        ClearSpread .vasTemp1
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA, JA_VALUES"
        SQL = SQL & vbCrLf & "  FROM Division "
        SQL = SQL & vbCrLf & " WHERE EQUIPCODE_TLA IN ('" & asJA1 & "', '" & asJA2 & "') "
        res = db_select_Vas(gLocal, SQL, .vasTemp1)
        
        For i = 1 To res
            If Trim(GetText(.vasTemp1, i, 2)) <> "" And Trim(GetText(.vasTemp1, i, 2)) <> "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = Format(Trim(GetText(.vasTemp1, i, 2)), "0000")
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            End If
        Next i
        
        Signal = Signal & "    " & JA_VALUE(1) & "     " & JA_CODE(1)
        Signal = Signal & "      " & JA_VALUE(2) & "     " & JA_CODE(2)
        ClearSpread .vasTemp1
        
                
    ElseIf asJA1 <> "" And asJA2 <> "" And asJA3 <> "" Then
        ClearSpread .vasTemp1
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT EQUIPCODE_TLA, JA_VALUES"
        SQL = SQL & vbCrLf & "  FROM Division "
        SQL = SQL & vbCrLf & " WHERE EQUIPCODE_TLA IN ('" & asJA1 & "', '" & asJA2 & "', '" & asJA3 & "') "
        res = db_select_Vas(gLocal, SQL, .vasTemp1)
        
        For i = 1 To res
            If Trim(GetText(.vasTemp1, i, 2)) <> "" And Trim(GetText(.vasTemp1, i, 2)) <> "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = Format(Trim(GetText(.vasTemp1, i, 2)), "0000")
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            ElseIf Trim(GetText(.vasTemp1, i, 2)) = "0" Then
                JA_CODE(i) = Trim(GetText(.vasTemp1, i, 1))
                JA_VALUE(i) = "    "
            End If
        Next i
        
        Signal = Signal & "    " & JA_VALUE(1) & "     " & JA_CODE(1)
        Signal = Signal & "      " & JA_VALUE(2) & "     " & JA_CODE(2)
        Signal = Signal & "      " & JA_VALUE(3) & "     " & JA_CODE(3)
        ClearSpread .vasTemp1

    End If
End With
     
    Dim FilNum
    Dim sFileName
    FilNum = FreeFile
    
    
    If Dir("c:\his\LIS", vbDirectory) <> "LIS" Then
        MkDir ("c:\his" & "\LIS")
    End If
    
    sFileName = BarCodeNo
    
    If Dir("c:\his\LIS\" & sFileName & ".txt", vbDirectory) <> sFileName & ".txt" Then
        Open "c:\his\LIS\" & sFileName & ".txt" For Append As FilNum
        Print #FilNum, Signal
        Close FilNum
    End If
    
'    Open "c:\his\LIS\" & sFileName & ".txt" For Append As FilNum
'    Print #FilNum, Signal
'    Close FilNum
    
    SQL = ""
    
    TLA_Division = 1
    
    SQL = "UPDATE SPSLMJBBI "
    SQL = SQL & vbCrLf & "   SET STAT_DVSN = 'T' "
    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colSpecNo)) & "' "
    SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colPID)) & "' "
    SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "   AND (STAT_DVSN IS NULL OR STAT_DVSN = '') "
    res = SendQuery(gServer, SQL)
    
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� �����ڸ� �̿��� ������ ������ ��ġ�� ���ڿ��� ����
'   �μ� :
'       1.pText      : �����ڷ� ������ ���ڿ�
'       2.pPosiion   : ��ġ
'       3.pDelimiter : ������
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition �μ��� 1�� ��� For�� Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '�ش� �÷�
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

'-- �ش� ȯ�� �˻��� H/L, Delta, Panic �����ϱ�
Function GetDecision(ByVal argSpcRow As Integer, ByVal strBarNo As String, ByVal strExamCode As String, ByVal strResult As String) As String
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
    
    '-- �ش� ȯ���� ����ġ,��Ÿ,�д� ã�ƿ���
    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW                 "
    SQL = SQL & vbCrLf & " FROM SPSLMFBIF                                                                                                                      "
    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= SYSDATE                                                                                                         "
    SQL = SQL & vbCrLf & "   AND USE_END_DY >= SYSDATE                                                                                                         "
    SQL = SQL & vbCrLf & "   and EXMN_CD = '" & Trim(strExamCode) & "' "
    Set rs_DPRef = cn_Ser.Execute(SQL)
    Do Until rs_DPRef.EOF
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
        SQL = SQL & vbCrLf & "           AND A.SPCM_NO = FN_LABCVTBCNO('" & strBarNo & "')                                       "
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
                    strBefoRslt = Mid(strBefoRslt, 2)
                    If IsNumeric(Mid(strBefoRslt, 1, 1)) = True Then
                        If InStr(1, strBefoRslt, ")") > 0 Then: strBefoRslt = Mid(strBefoRslt, 1, InStr(1, strBefoRslt, ")") - 1)
                        Exit Do
                    End If
                Loop
            End If
            
            '-- ������ ������� ��
            '-- ������� ��ġ�� ��쿡�� ���Ѵ�.
            If IsNumeric(strDestRslt) Then
                If strSex = "M" Then
                    If IsNumeric(rs_DPRef.Fields("MALE_HIGH")) Then
                        If CDbl(strDestRslt) > CDbl(rs_DPRef.Fields("MALE_HIGH")) Then
                            strHLVal = "H"
                        Else
                            strHLVal = " "
                        End If
                    Else
                        strHLVal = ""
                    End If
                    
                    If IsNumeric(rs_DPRef.Fields("MALE_LOW")) Then
                        If CDbl(strDestRslt) < CDbl(rs_DPRef.Fields("MALE_LOW")) Then
                            strHLVal = "L"
                        Else
                            strHLVal = " "
                        End If
                    Else
                        strHLVal = ""
                    End If
                
                Else
                    If IsNumeric(rs_DPRef.Fields("FEML_HIGH")) Then
                        If CDbl(strDestRslt) > CDbl(rs_DPRef.Fields("FEML_HIGH")) Then
                            strHLVal = "H"
                        Else
                            strHLVal = " "
                        End If
                    Else
                        strHLVal = ""
                    End If
                    If IsNumeric(rs_DPRef.Fields("FEML_LOW")) Then
                        If CDbl(strDestRslt) < CDbl(rs_DPRef.Fields("FEML_LOW")) Then
                            strHLVal = "L"
                        Else
                            strHLVal = " "
                        End If
                    Else
                        strHLVal = ""
                    End If
                End If
            Else
                strHLVal = ""
            End If
            
            '-- Delta ����  (�Ʒ� ������ �´��� ���� �ʿ���...��)
            '-- ������� ��ġ�� ��쿡�� ���Ѵ�.
            If IsNumeric(strDestRslt) And IsNumeric(strBefoRslt) = True Then
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
                            strDelta = strDelta / CInt(rs_Delta.Fields("DELTA_TERM_DT"))        '�Ⱓ�� ��ȭ����
                    Case 4:     '4 �Ⱓ�� ��ȭ�� = ��ȭ�� / �Ⱓ
                            strDelta = ""
                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '��ȭ��
                            strDelta = CDbl(strDelta) / CInt(rs_Delta.Fields("DELTA_TERM_DT"))  '�Ⱓ�� ��ȭ��
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
            
            
            
            '-- Panic ����
            '-- ������� ��ġ�� ��쿡�� ���Ѵ�.
            If IsNumeric(strDestRslt) Then
                Select Case Trim(rs_DPRef.Fields("PANC_DVSN"))
                    Case 0:     '0 ������
                            strPanic = ""
                    Case 1:     '1 ���Ѹ�
                            If IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
                                If CDbl(strDestRslt) > rs_DPRef.Fields("PANC_HIGH") Then
                                    strPanic = "P"
                                Else
                                    strPanic = " "
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case 2:     '2 ���Ѹ�
                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) Then
                                If CDbl(strDestRslt) < rs_DPRef.Fields("PANC_LOW") Then
                                    strPanic = "P"
                                Else
                                    strPanic = " "
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case 3:     '3 ��� ���
                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) And IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
                                If (CDbl(strDestRslt) < rs_DPRef.Fields("PANC_LOW") Or CDbl(strDestRslt) > rs_DPRef.Fields("PANC_HIGH")) Then
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
            rs_Delta.MoveNext
        Loop
        
        rs_DPRef.MoveNext
    Loop
    
    Set rs_DPRef = Nothing
        
    GetDecision = strHLVal & "/" & strDelta & "/" & strPanic
    
End Function

Function Make_Remark(asExamCode As String, asSex As String, asResult As String)
'///////////// �ڸ�Ʈ ���� (�˻��)
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
    SQL = SQL & vbCrLf & " WHERE EXAMCODE = '" & asExamCode & "' "
    SQL = SQL & vbCrLf & ""
    res = db_select_Col(gLocal, SQL)

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

    If Comment_Gubun > 0 Then
        Select Case Comment_Gubun '////////// 0:�������, 1: �˻���ü����, 2:�ش�˻�����
            
            Case "1" '/// ��ü����  // ���� Function  ����
                

            Case "2" '/// �ش�˻�����

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
                    res = db_select_Col(gLocal, SQL)
                    
                    gComment_Code = gReadBuf(0)
                    
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
                    
                    res = db_select_Col(gLocal, SQL)
                    
                ElseIf Comment_MFGubun = "2" Then
                    
                    SQL = ""
                    SQL = SQL & vbCrLf & "SELECT CNTS "
                    SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
                    SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_CodeH & "' "
                    SQL = SQL & vbCrLf & ""
                    res = db_select_Col(gLocal, SQL)
                    
                    gComment_Code = gReadBuf(0)
                    
                End If
            
        End Select

    End If


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
Function RsltState_Check(asSpecNo As String, asExamCode As String) As String '/// ��� ���� : (�׷��ڵ�/��Ƽ�ڵ�) : ���°� �߰����� �����϶�
    Dim PRSC_CD_G       As String
    Dim EXMN_CD         As String
    Dim PRSC_CD_M       As String
    Dim PRSC_CD_B       As String
    
    RsltState_Check = ""
    PRSC_CD_G = " "
    PRSC_CD_M = " "
    PRSC_CD_B = " "
    Save_Raw_Data GetDateFull & " [RsltState_Check] "
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    'SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = '" & asExamCode & "' "
    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE ('%G%') "
'    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
    res = db_select_Col(gServer, SQL)
    If gReadBuf(0) <> "" Then PRSC_CD_G = gReadBuf(0): gReadBuf(0) = ""
    Save_Raw_Data GetDateFull & " [RsltState_Check 1 ] "
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    'SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('M') "
'    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
    res = db_select_Col(gServer, SQL)
       
    If gReadBuf(0) <> "" Then PRSC_CD_M = gReadBuf(0): gReadBuf(0) = ""
    Save_Raw_Data GetDateFull & " [RsltState_Check 2 ] "

    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
    'SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('B') "
'    SQL = SQL & vbCrLf & "   AND R1.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
    res = db_select_Col(gServer, SQL)
       
    If gReadBuf(0) <> "" Then PRSC_CD_B = gReadBuf(0): gReadBuf(0) = ""
    Save_Raw_Data GetDateFull & " [RsltState_Check 3 ] "
    
    
    RsltState_Check = PRSC_CD_G & "/" & PRSC_CD_M & "/" & PRSC_CD_B
    
End Function

