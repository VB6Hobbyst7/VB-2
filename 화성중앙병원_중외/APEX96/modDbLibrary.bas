Attribute VB_Name = "modDbLibrary"
Option Explicit


Private Function f_subSet_RefVal(ByVal strORCD As String, ByVal strSubCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
Dim rs_svr As ADODB.Recordset

On Error GoTo ErrorTrap
    
    strRSLT = Replace(strRSLT, "<", "")
    strRSLT = Replace(strRSLT, ">", "")
    f_subSet_RefVal = " "
    
    f_subSet_RefVal = ""
          SQL = "Select REFHIGH, REFLOW  "
    SQL = SQL & "  From EQPMASTER"
    SQL = SQL & " Where EQUIPNO = '" & gEquip & "' "
    SQL = SQL & "   And EXAMCODE =  '" & strORCD & "'"
'    SQL = SQL & "   And SUBCODE =  '" & strSubCD & "'"
    
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If Res > 0 Then
        If IsNumeric(strRSLT) And IsNumeric(Trim(gReadBuf(0))) And IsNumeric(Trim(gReadBuf(1))) Then
            If Val(strRSLT) > Val(Trim(gReadBuf(0))) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRSLT) < Val(Trim(gReadBuf(1))) Then
                f_subSet_RefVal = "L"
            Else
                f_subSet_RefVal = " "
            End If
        Else
            f_subSet_RefVal = " "
        End If
    End If
    
Exit Function

ErrorTrap:
    f_subSet_RefVal = " "
'    Set AdoRs_ORACLE = Nothing
'    Call ErrMsgProc(CallForm)
     
End Function

Function SaveTransDataW(ByVal argSpcRow As Integer, Optional ByVal pResult As String) As Integer
    Dim lsID            As String
    Dim strRegDate      As String
    Dim sResult         As String
    
On Error GoTo ErrHandle

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        strRegDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))

        cn_Ser.BeginTrans
        
        '-- ������ ����� �����ϱ�
        sResult = Mid(pResult, 1, 4000)
        
        If lsID <> "" And strRegDate <> "" And sResult <> "" Then
            '-- �������
                  SQL = "Update resultofnum Set" & vbCrLf
            SQL = SQL & "   resultindate        = to_char(sysdate,'yyyymmdd')   " & vbCrLf
            SQL = SQL & " , resultintime        = to_char(sysdate,'HH24MI')     " & vbCrLf
            SQL = SQL & " , resultinid          = '" & gUserID & "'             " & vbCrLf
            SQL = SQL & " , resultflag          = '1'                           " & vbCrLf
            SQL = SQL & " , textresultval       = '" & sResult & "'             " & vbCrLf
            'SQL = SQL & " , ANALYZERCODE        = '" & gMachCD & "'             " & vbCrLf 'APEX = 41
            SQL = SQL & " Where spcmno          = '" & lsID & "'                " & vbCrLf
            SQL = SQL & "   And resultitemcode  = '" & gAssayNM.APEX96MR_CD & "' " & vbCrLf
            SQL = SQL & "   And resultflag      < '2'                           " & vbCrLf

            Call SetSQLData("�������", SQL)

            Res = SendQuery(gServer, SQL)
            
            If Res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
            If Res = 1 Then
                SQL = " UPDATE registinfos SET" & vbCr
                SQL = SQL & " RESULTSTATE = '3'" & vbCr
                SQL = SQL & " WHERE SPCMNO = '" & lsID & "'" & vbCr
                'SQL = SQL & "   AND ORDERCODE = '" & strOrdCd & "'" & vbCr
                SQL = SQL & "   AND CLAS = 4" & vbCr
                SQL = SQL & "   AND RESULTSTATE < '4'" & vbCr
                
                Call SetSQLData("��������", SQL)
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
            
            SaveTransDataW = 1
            
        End If
        
        cn_Ser.CommitTrans
        
    
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    cn_Ser.RollbackTrans
    
End Function

Function SaveTransData_JWINFO(ByVal argSpcRow As Integer, Optional ByVal pResult As String) As Integer
    Dim lsID            As String
    Dim strRegDate      As String
    Dim sResult         As String
    Dim blnBegin        As Boolean

On Error GoTo ErrHandle

    With frmInterface
        blnBegin = False
        SaveTransData_JWINFO = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        strRegDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))

        cn_Ser.BeginTrans
        blnBegin = True
        
        '-- ������ ����� �����ϱ�
        sResult = Mid(pResult, 1, 4000)
        
        If lsID <> "" And strRegDate <> "" And sResult <> "" Then
            '-- �������
                  SQL = "Update resultofnum Set" & vbCrLf
            SQL = SQL & "   resultindate        = to_char(sysdate,'yyyymmdd')   " & vbCrLf
            SQL = SQL & " , resultintime        = to_char(sysdate,'HH24MI')     " & vbCrLf
            SQL = SQL & " , resultinid          = '" & gUserID & "'             " & vbCrLf
            SQL = SQL & " , resultflag          = '1'                           " & vbCrLf
            SQL = SQL & " , textresultval       = '" & sResult & "'             " & vbCrLf
            'SQL = SQL & " , ANALYZERCODE        = '" & gMachCD & "'             " & vbCrLf 'APEX = 41
            SQL = SQL & " Where spcmno          = '" & lsID & "'                " & vbCrLf
            SQL = SQL & "   And resultitemcode  = '" & gAssayNM.APEX96MR_CD & "' " & vbCrLf
            SQL = SQL & "   And resultflag      < '2'                           " & vbCrLf

            Call SetSQLData("�������", SQL)

            Res = SendQuery(gServer, SQL)
            
            If Res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
            
            If Res = 1 Then
                SQL = " UPDATE registinfos SET" & vbCr
                SQL = SQL & " RESULTSTATE = '3'" & vbCr
                SQL = SQL & " WHERE SPCMNO = '" & lsID & "'" & vbCr
                'SQL = SQL & "   AND ORDERCODE = '" & strOrdCd & "'" & vbCr
                SQL = SQL & "   AND CLAS = 4" & vbCr
                SQL = SQL & "   AND RESULTSTATE < '4'" & vbCr
                
                Call SetSQLData("��������", SQL)
                
                Res = SendQuery(gServer, SQL)
                
                If Res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
            
            SaveTransData_JWINFO = 1
            
        End If
        
        cn_Ser.CommitTrans
        
    
    End With

Exit Function

ErrHandle:
    SaveTransData_JWINFO = -1
    If blnBegin Then
        cn_Ser.RollbackTrans
    End If
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    
    GetSampleInfoW = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'          SQL = " SELECT DISTINCT '' AS ��������"
'    SQL = SQL & ", '' AS ��Ʈ��ȣ"
'    SQL = SQL & ", '' AS ������ȣ"
'    SQL = SQL & ", '' AS �Կ�"
'    SQL = SQL & ", '' AS �̸�"
'    SQL = SQL & ", '' AS ����"
'    SQL = SQL & ", '' AS ����" & vbCrLf
'    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'    SQL = SQL & " WHERE QC_BAR_NO = '" & sBarcode & "'" & vbCrLf
    
    
                                   
     
          'SQL = "SELECT P.PbsPatNam, O.OSPCHTNUM, R.ResLabCod, E.LabShtNam, R.ResOcmNum, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
          SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, R.ResOcmNum " & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & sBarcode & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    Res = GetDBSelectColumn(gServer, SQL)
        
    If Res = 1 Then
        SetText frmInterface.vasID, "1", asRow, colCheckBox
        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
        'SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colHOSPDATE       '������
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colCHARTNO        'íƮ��ȣ
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID            '��Ϲ�ȣ(����� �ʿ�)
        'SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colINOUT          '��/��
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPNAME          'ȯ�ڸ�
        'SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, colPSEX           '����
        'SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, colPAGE           '����
        
        GetSampleInfoW = 1
   
    Else
        GetSampleInfoW = -1
    End If

    frmInterface.vasID.RowHeight(-1) = 12

End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_BIT(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    
    GetSampleInfoW_BIT = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
     
          'SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, E.LabShtNam, R.ResOcmNum, R.ResLabCod, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
          SQL = "SELECT DISTINCT P.PbsPatNam, O.OSPCHTNUM, R.ResOcmNum, R.ResLabCod AS EXAMCODE , R.ResOdrSeq, R.ResSeq, R.ResSubSeq " & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & sBarcode & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    'SetRawData "[SQL1]" & SQL
    
    Set RS = cn_Ser.Execute(SQL)

    With frmInterface
        Do Until RS.EOF
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("EXAMCODE")) & "',"
            
            SetText .vasID, "1", asRow, colCheckBox
            SetText .vasID, sBarcode, asRow, colBARCODE
            SetText .vasID, Trim(RS.Fields("OSPCHTNUM")), asRow, colCHARTNO         'íƮ��ȣ(������� ����� �ʿ�)
            SetText .vasID, Trim(RS.Fields("ResOcmNum")), asRow, colPID             '��Ϲ�ȣ(���     ����� �ʿ�)
            SetText .vasID, Trim(RS.Fields("PbsPatNam")), asRow, colPNAME           'ȯ�ڸ�
            
            
            'SetText .vasID, "12345", asRow, colCHARTNO         'íƮ��ȣ
            'SetText .vasID, "67890", asRow, colPID            '��Ϲ�ȣ(����� �ʿ�)
            'SetText .vasID, "ȫ�渱", asRow, colPNAME           'ȯ�ڸ�
            
            '-- ȭ�鿡 ǥ��
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = asRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    '-- �������� SEQ
                    gArrEquip(intCol - colState, 7) = Trim(RS.Fields("ResOdrSeq")) & "|" & Trim(RS.Fields("ResSeq")) & "|" & Trim(RS.Fields("ResSubSeq"))   '�������� ��ȣ's
                    'gArrEquip(intCol - colState, 7) = "987" & "|" & "654" & "|" & "321"
                    Exit For
                End If
            Next
    
            RS.MoveNext
        Loop
    
        GetSampleInfoW_BIT = 1
    
    End With
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
        
    frmInterface.vasID.RowHeight(-1) = 12

End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_NTL(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim strGubun            As String
    Dim intCol              As Integer
    Dim GetOrderExamCode    As String
    Dim RS1                 As ADODB.Recordset
    Dim strRegDate          As String
    Dim lngRegNo            As Long
    
    
    GetSampleInfoW_NTL = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    strRegDate = "20" & Format(Mid(sBarcode, 1, 6), "##-##-##")
    lngRegNo = Val(Mid(sBarcode, 7))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
'    If InStr(sBarcode, "-") <= 0 Then
'        Exit Function
'    End If
    
    
    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient    'GetPatientResultList02
    'Set RS = cn_Ser.Execute("Exec Interface_GetPatientResult02 '" & gWKCD & "','" & Format(mGetP(sBarcode, 1, "-"), "####-##-##") & "','" & Val(mGetP(sBarcode, 2, "-")) & "'")
    'Set RS = cn_Ser.Execute("Exec Interface_GetPatientResult02 '" & gWKCD & "','" & Format(Mid(sBarcode, 1, 6), "yyyy-mm-dd") & "','" & Val(Mid(sBarcode, 7)) & "'")
    Set RS = cn_Ser.Execute("Exec Interface_GetPatientResult02 '" & gWKCD & "','" & strRegDate & "','" & lngRegNo & "'")
          
    With frmInterface
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                SetText .vasID, "1", asRow, colCheckBox
                SetText .vasID, Trim(RS.Fields("LabRegDate")), asRow, colHOSPDATE
                SetText .vasID, Trim(RS.Fields("PatientChartNo")), asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("LabRegNo")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("PatientName")), asRow, colPNAME
                SetText .vasID, Trim(RS.Fields("CompanyCode")), asRow, colDISKNO
                SetText .vasID, Trim(RS.Fields("PatientBirthDay")), asRow, colPOSNO
                SetText .vasID, Trim(RS.Fields("PatientSex")), asRow, colPSEX
                SetText .vasID, Trim(RS.Fields("PatientAge")), asRow, colPAGE
                Select Case Trim(RS.Fields("OrderCode")) & ""
                    Case "63100":   strGubun = "INHALANT"
                    Case "63200":   strGubun = "FOOD"
                    Case "63300":   strGubun = "ATOPY"
                End Select
                
                

                      SQL = " SELECT OrderCode, TestCode, TestSubCode " & vbCrLf
                SQL = SQL & "   FROM LC11_NTL..LabRegResult " & vbCrLf
                SQL = SQL & "  WHERE LABREGDATE = '" & strRegDate & "'" & vbCrLf
                SQL = SQL & "    AND LABREGNO   = " & lngRegNo & vbCrLf
                SQL = SQL & "    AND ORDERCODE  = '" & Trim(RS.Fields("OrderCode")) & "'"
                
                'cn_Ser.CursorLocation = adUseClient
                Set RS1 = cn_Ser.Execute(SQL, , 1)
                If Not RS1.EOF = True And Not RS1.BOF = True Then
                    Do Until RS1.EOF
                        '-- ȭ�鿡 ǥ��
                        For intCol = colState + 1 To .vasID.MaxCols
                            If Trim(RS1.Fields("TestSubCode")) = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                .vasID.Row = asRow
                                .vasID.Col = intCol
                                .vasID.BackColor = vbYellow
                                '-- �������� SEQ
                                gArrEquip(intCol - colState, 9) = Trim(RS1.Fields("OrderCode")) & "|" & Trim(RS1.Fields("TestCode")) & "|" & Trim(RS1.Fields("TestSubCode"))   '�������� ��ȣ's
                                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS1.Fields("TestSubCode")) & "',"
                                Exit For
                            End If
                        Next
                        
                        RS1.MoveNext
                    Loop
                End If
                RS1.Close
                RS.MoveNext
            Loop
        
            GetSampleInfoW_NTL = 1
        
        End If
    End With
        
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
    
    frmInterface.vasID.RowHeight(-1) = 12

End Function


'-- �˻��� ���� ��������
Function GetSampleInfoW_PHILL(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim strGubun            As String
    Dim intCol              As Integer
    Dim GetOrderExamCode    As String
    Dim RS1                 As ADODB.Recordset
    Dim strRegDate          As String
    Dim lngRegNo            As Long
    
    
    GetSampleInfoW_PHILL = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
'    strRegDate = "20" & Format(Mid(sBarcode, 1, 6), "##-##-##")
    strRegDate = Mid(sBarcode, 1, 8)
    lngRegNo = Val(Mid(sBarcode, 9))
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    '-- Record Count ������
'    cn_Ser.CursorLocation = adUseClient    'GetPatientResultList02
'    Set RS = cn_Ser.Execute("Exec Interface_GetPatientResult02 '" & gWKCD & "','" & strRegDate & "','" & lngRegNo & "'")
          
          
          
    SQL = ""
    SQL = SQL & "SELECT DISTINCT P.request_date AS ��������, P.exam_no AS ������ȣ, P.company_code AS �Ƿ�ó, P.chart_no AS ��Ʈ��ȣ, p.personal_id, p.person_name AS �̸�, " & vbCr
    SQL = SQL & "       P.worker_code, P.patient_kind, P.person_sex AS ����, P.person_age AS ����, " & vbCr
    SQL = SQL & "       R.exam_order, R.exam_code AS ITEM, E.exam_ename, R.pro_code AS ó���ڵ� " & vbCr
    SQL = SQL & "  FROM trust P, trures R, examitem E " & vbCr
    SQL = SQL & " WHERE P.request_date = '" & strRegDate & "'" & vbCr
    SQL = SQL & "   AND P.exam_no = '" & lngRegNo & "'"
'    SQL = SQL & "   AND R.exam_part collate latin1_general_cs_as = 'Z' " & vbCr
    SQL = SQL & "   AND R.pro_code IN ('" & gAssayNM.INHALANT_CD & "','" & gAssayNM.FOOD_CD & "','" & gAssayNM.ATOPY_CD & "') " & vbCr
    SQL = SQL & "   AND R.exam_code <> 'X999' " & vbCr
    SQL = SQL & "   AND P.request_date = R.request_date " & vbCr
    SQL = SQL & "   AND P.exam_no = R.exam_no " & vbCr
    SQL = SQL & "   AND R.exam_code = E.exam_code " & vbCr
    SQL = SQL & " ORDER BY P.request_date, P.exam_no "

    Call SetSQLData("���ڵ���ȸ", SQL)
          
    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
          
    With frmInterface
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                SetText .vasID, "1", asRow, colCheckBox
                SetText .vasID, Trim(RS.Fields("��������")), asRow, colHOSPDATE
                SetText .vasID, Trim(RS.Fields("��Ʈ��ȣ")), asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("�̸�")), asRow, colPNAME
                SetText .vasID, Trim(RS.Fields("�Ƿ�ó")), asRow, colDISKNO
'                SetText .vasID, Trim(RS.Fields("PatientBirthDay")), asRow, colPOSNO
                SetText .vasID, Trim(RS.Fields("����")), asRow, colPSEX
                SetText .vasID, Trim(RS.Fields("����")), asRow, colPAGE
                
'                Select Case Trim(RS.Fields("OrderCode")) & ""
'                    Case "63100":   strGubun = "INHALANT"
'                    Case "63200":   strGubun = "FOOD"
'                    Case "63300":   strGubun = "ATOPY"
'                End Select
                
                strGubun = ""
                
'''                Select Case Trim(RS.Fields("ó���ڵ�")) & ""        'ó���ڵ� ??
'''                    Case gAssayNM.INHALANT_CD: strGubun = "INHALANT"
'''                    Case gAssayNM.FOOD_CD:     strGubun = "FOOD"
'''                    Case gAssayNM.ATOPY_CD:    strGubun = "ATOPY"
''''                    Case Else:                 strGubun = "ó�����"
'''                End Select
'''
'''                SetText .vasID, strGubun, asRow, colINOUT
                
                '-- ȭ�鿡 ǥ��
                For intCol = colState + 1 To .vasID.MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                        .vasID.Row = asRow
                        .vasID.Col = intCol
                        .vasID.BackColor = vbYellow
                        '-- �������� SEQ
                        'gArrEquip(intCol - colState, 9) = Trim(RS1.Fields("OrderCode")) & "|" & Trim(RS1.Fields("TestCode")) & "|" & Trim(RS1.Fields("TestSubCode"))   '�������� ��ȣ's
                        GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
                        Exit For
                    End If
                Next
                
'''
'''                      SQL = " SELECT OrderCode, TestCode, TestSubCode " & vbCrLf
'''                SQL = SQL & "   FROM LC11_NTL..LabRegResult " & vbCrLf
'''                SQL = SQL & "  WHERE LABREGDATE = '" & strRegDate & "'" & vbCrLf
'''                SQL = SQL & "    AND LABREGNO   = " & lngRegNo & vbCrLf
'''                SQL = SQL & "    AND ORDERCODE  = '" & Trim(RS.Fields("OrderCode")) & "'"
'''
'''                'cn_Ser.CursorLocation = adUseClient
'''                Set RS1 = cn_Ser.Execute(SQL, , 1)
'''                If Not RS1.EOF = True And Not RS1.BOF = True Then
'''                    Do Until RS1.EOF
'''                        '-- ȭ�鿡 ǥ��
'''                        For intCol = colState + 1 To .vasID.MaxCols
'''                            If Trim(RS1.Fields("TestSubCode")) = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
'''                                .vasID.Row = asRow
'''                                .vasID.Col = intCol
'''                                .vasID.BackColor = vbYellow
'''                                '-- �������� SEQ
'''                                gArrEquip(intCol - colState, 9) = Trim(RS1.Fields("OrderCode")) & "|" & Trim(RS1.Fields("TestCode")) & "|" & Trim(RS1.Fields("TestSubCode"))   '�������� ��ȣ's
'''                                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS1.Fields("TestSubCode")) & "',"
'''                                Exit For
'''                            End If
'''                        Next
'''
'''                        RS1.MoveNext
'''                    Loop
'''                End If
'''                RS1.Close
                RS.MoveNext
            Loop
        
            GetSampleInfoW_PHILL = 1
        
        End If
    End With
        
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
    
    frmInterface.vasID.RowHeight(-1) = 12

End Function



'-- �˻��� ���� ��������
Function GetSampleInfoW_AMIS(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim strGubun            As String
    Dim intCol              As Integer
    Dim GetOrderExamCode    As String
    Dim RS1                 As ADODB.Recordset
'    Dim strRegDate          As String
'    Dim lngRegNo            As Long
        
    GetSampleInfoW_AMIS = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If Len(sBarcode) < 5 Then
        Exit Function
    End If
    
    If sBarcode = "" Then
        Exit Function
    End If
    

    SQL = ""
    SQL = SQL & "SELECT P.PATID AS ��Ʈ��ȣ, P.PATNAME AS �̸�, P.SEX AS ����, O.ACPTDATE AS ��������" & vbCr
    SQL = SQL & ", O.ACPTSEQ, O.RSVACPTSTATE, O.RESULTSTATE, O.DEPTCODE, O.ORDERDATE, O.SLIPNO AS ������ȣ, O.IOFLAG, O.ORDERCODE, O.ORDERNAME" & vbCr
    SQL = SQL & ", R.SPCMNO AS ���ڵ��ȣ, R.RESULTFLAG, R.RESULTNO, R.RESULTITEMCODE as ITEM " & vbCr
    SQL = SQL & "  FROM registinfos O, resultofnum R, PATMST P " & vbCr
    SQL = SQL & " WHERE O.acptdate = R.acptdate " & vbCr
    SQL = SQL & "   AND R.SPCMNO = '" & sBarcode & "'" & vbCr
    SQL = SQL & "   AND O.patid = R.patid " & vbCr
    SQL = SQL & "   AND O.acptseq = R.acptseq " & vbCr
    SQL = SQL & "   AND O.patid = P.patid " & vbCr
    SQL = SQL & "   AND O.CLAS = 4 " & vbCr '�ӻ󺴸�
    'SQL = SQL & "   AND O.ORDERCODE IN ('" & gAssayNM.INHALANT_CD & "','" & gAssayNM.FOOD_CD & "') " & vbCr
    SQL = SQL & "   AND O.ORDERCODE IN ('" & gAssayNM.APEX96M_CD & "') " & vbCr
    SQL = SQL & "   AND R.RESULTFLAG = 0 " & vbCr
'    SQL = SQL & "   AND R.resultitemcode in (" & gAllExam & ")" & vbCr
    Call SetSQLData("���ڵ���ȸ", SQL)
    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    With frmInterface
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                SetText .vasID, "1", asRow, colCheckBox
                SetText .vasID, Trim(RS.Fields("��������")) & "", asRow, colHOSPDATE
                SetText .vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", asRow, colBARCODE
                SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("�̸�")) & "", asRow, colPNAME
                'SetText .vasID, Trim(RS.Fields("�Ƿ�ó")), asRow, colDISKNO
                SetText .vasID, Trim(RS.Fields("����")) & "", asRow, colPSEX
                'SetText .vasID, Trim(RS.Fields("����")), asRow, colPAGE
                
                Select Case Trim(RS.Fields("ORDERCODE")) & ""        'ó���ڵ� ??
                    Case gAssayNM.INHALANT_CD: SetText .vasID, "INHALANT", .vasID.MaxRows, colINOUT
                    Case gAssayNM.FOOD_CD:     SetText .vasID, "FOOD", .vasID.MaxRows, colINOUT
                    Case gAssayNM.ATOPY_CD:    SetText .vasID, "ATOPY", .vasID.MaxRows, colINOUT
                    Case gAssayNM.APEX96M_CD:  SetText .vasID, "96M", .vasID.MaxRows, colINOUT
                    Case Else:
                                               SetBackColor .vasID, .vasID.MaxRows, .vasID.MaxRows, 1, colState, 202, 255, 112
                                               SetText .vasID, "ó�����", .vasID.MaxRows, colINOUT
                                                                        
                End Select
                strGubun = ""
                
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
                
                RS.MoveNext
            Loop
            GetSampleInfoW_AMIS = 1
        End If
    End With
        
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
    
    frmInterface.vasID.RowHeight(-1) = 12

End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_JWINFO(ByVal asRow As Long) As Integer
    Dim sBarcode            As String
    Dim strGubun            As String
    Dim intCol              As Integer
    Dim GetOrderExamCode    As String
    Dim RS1                 As ADODB.Recordset
        
On Error GoTo Err
    
    GetSampleInfoW_JWINFO = -1
    
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If Len(sBarcode) < 5 Then
        Exit Function
    End If
    
    If sBarcode = "" Then
        Exit Function
    End If

    SQL = ""
    SQL = SQL & "SELECT P.PATID AS ��Ʈ��ȣ, P.PATNAME AS �̸�, P.SEX AS ����, O.ACPTDATE AS ��������" & vbCr
    SQL = SQL & ", O.ACPTSEQ, O.RSVACPTSTATE, O.RESULTSTATE, O.DEPTCODE, O.ORDERDATE, O.SLIPNO AS ������ȣ, O.IOFLAG, O.ORDERCODE, O.ORDERNAME" & vbCr
    SQL = SQL & ", R.SPCMNO AS ���ڵ��ȣ, R.RESULTFLAG, R.RESULTNO, R.RESULTITEMCODE as ITEM " & vbCr
    SQL = SQL & "  FROM registinfos O, resultofnum R, PATMST P " & vbCr
    SQL = SQL & " WHERE O.acptdate = R.acptdate " & vbCr
    SQL = SQL & "   AND R.SPCMNO = '" & sBarcode & "'" & vbCr
    SQL = SQL & "   AND O.patid = R.patid " & vbCr
    SQL = SQL & "   AND O.acptseq = R.acptseq " & vbCr
    SQL = SQL & "   AND O.patid = P.patid " & vbCr
    SQL = SQL & "   AND O.CLAS = 4 " & vbCr '�ӻ󺴸�
    SQL = SQL & "   AND O.ORDERCODE IN ('" & gAssayNM.APEX96M_CD & "') " & vbCr
    SQL = SQL & "   AND R.RESULTFLAG = 0 " & vbCr
'    SQL = SQL & "   AND R.resultitemcode in (" & gAllExam & ")" & vbCr
    Call SetSQLData("���ڵ���ȸ", SQL)
    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    With frmInterface
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                SetText .vasID, "1", asRow, colCheckBox
                SetText .vasID, Trim(RS.Fields("��������")) & "", asRow, colHOSPDATE
                SetText .vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", asRow, colCHARTNO
                SetText .vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", asRow, colBARCODE
                SetText .vasID, Trim(RS.Fields("������ȣ")) & "", asRow, colPID
                SetText .vasID, Trim(RS.Fields("�̸�")) & "", asRow, colPNAME
                'SetText .vasID, Trim(RS.Fields("�Ƿ�ó")), asRow, colDISKNO
                SetText .vasID, Trim(RS.Fields("����")) & "", asRow, colPSEX
                'SetText .vasID, Trim(RS.Fields("����")), asRow, colPAGE
                
                Select Case Trim(RS.Fields("ORDERCODE")) & ""        'ó���ڵ� ??
                    Case gAssayNM.INHALANT_CD: SetText .vasID, "INHALANT", .vasID.MaxRows, colINOUT
                    Case gAssayNM.FOOD_CD:     SetText .vasID, "FOOD", .vasID.MaxRows, colINOUT
                    Case gAssayNM.ATOPY_CD:    SetText .vasID, "ATOPY", .vasID.MaxRows, colINOUT
                    Case gAssayNM.APEX96M_CD:  SetText .vasID, "96M", .vasID.MaxRows, colINOUT
                    Case Else:
                                               SetBackColor .vasID, .vasID.MaxRows, .vasID.MaxRows, 1, colState, 202, 255, 112
                                               SetText .vasID, "ó�����", .vasID.MaxRows, colINOUT
                                                                        
                End Select
                strGubun = ""
                
                GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("ITEM")) & "',"
                
                RS.MoveNext
            Loop
            GetSampleInfoW_JWINFO = 1
        End If
    End With
        
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
        gOrderExam = GetOrderExamCode
    End If
    
    frmInterface.vasID.RowHeight(-1) = 12

Exit Function
Err:

End Function


Public Function GetSameRowNum(ByVal strBarNo As String) As Integer
    Dim i As Integer

    GetSameRowNum = 0
    With frmInterface.vasID
        For i = 1 To .MaxRows
            .Row = i
            .Col = colBARCODE
            If Trim(.Text) = strBarNo Then
                GetSameRowNum = i
                Exit Function
            End If
        Next
    End With
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfoW_GINUSDLL(ByVal asRow As Long) As Integer
    Dim pBarNo  As String
    Dim i       As Integer
    Dim intCol  As Integer
    Dim strItem As String
    
    '-- ������
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    GetSampleInfoW_GINUSDLL = -1
    
    pBarNo = Trim(GetText(frmInterface.vasID, asRow, colBARCODE))
    
    If pBarNo = "" Then
        Exit Function
    End If
    
    '-- �˻�ITEM ��������
                 strRequest = "jobs" + vbTab + "Q" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + pBarNo + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.MCD + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- ���ڵ�� �˻��� ��ȸ(https://211.172.17.66)
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    With frmInterface.vasID
        If UBound(varResponse) > 0 Then
            For i = 0 To UBound(varResponse) - 1
                SetText frmInterface.vasID, "1", asRow, colCheckBox
                SetText frmInterface.vasID, Mid(mGetP(varResponse(i), 25, vbTab), 1, 8), asRow, colHOSPDATE
                SetText frmInterface.vasID, mGetP(varResponse(i), 0, vbTab), asRow, colBARCODE
                SetText frmInterface.vasID, mGetP(varResponse(i), 7, vbTab), asRow, colPID
                SetText frmInterface.vasID, mGetP(varResponse(i), 26, vbTab), asRow, colPNAME
                
                Select Case mGetP(varResponse(i), 29, vbTab)
                    Case "O": SetText frmInterface.vasID, "�ܷ�", asRow, colINOUT
                    Case "E": SetText frmInterface.vasID, "����", asRow, colINOUT
                    Case "I": SetText frmInterface.vasID, "�Կ�", asRow, colINOUT
                End Select
                
                
                For intCol = colState + 1 To .MaxCols
                    If mGetP(varResponse(i), 6, vbTab) = gArrEquip(intCol - colState, 3) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        '-- �������� SEQ
                        gArrEquip(intCol - colState, 7) = mGetP(varResponse(i), 3, vbTab) & "|" & mGetP(varResponse(i), 4, vbTab) & "|" & mGetP(varResponse(i), 5, vbTab)
                        Exit For
                    End If
                Next intCol
            Next i
        Else
            SetText frmInterface.vasID, "No Order", asRow, colState
        End If
    End With
    
    GetSampleInfoW_GINUSDLL = 1

'          SQL = " SELECT DISTINCT REQ_DT AS ��������"
'    SQL = SQL & ", LOT_NO AS ��Ʈ��ȣ"
'    SQL = SQL & ", REQ_SEQ AS ������ȣ"
'    SQL = SQL & ", '�Կ�' AS �Կ�"
'    SQL = SQL & ", 'ȫ�浿' AS �̸�"
'    SQL = SQL & ", '����' AS ����"
'    SQL = SQL & ", REQ_SEQ AS ����" & vbCrLf
'    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'    SQL = SQL & " WHERE QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
'    SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
'
'    Res = GetDBSelectColumn(gServer, SQL)
        
'    If Res = 1 Then
'        SetText frmInterface.vasID, "1", asRow, colCheckBox
'        SetText frmInterface.vasID, sBarcode, asRow, colBARCODE
'        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colHOSPDATE
'        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colCHARTNO
'        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colPID
'        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colINOUT
'        SetText frmInterface.vasID, Trim(gReadBuf(4)), asRow, colPNAME
'        SetText frmInterface.vasID, Trim(gReadBuf(5)), asRow, colPSEX
'        SetText frmInterface.vasID, Trim(gReadBuf(6)), asRow, colPAGE
'
'        GetSampleInfoW_GINUSDLL = 1
'
'    Else
'        GetSampleInfoW_GINUSDLL = -1
'    End If
'
'    frmInterface.vasID.RowHeight(-1) = 12

End Function

'Function GetSampleInfoR(ByVal asRow As Long) As Integer
'    Dim sBarcode As String
'    Dim sSpecNo As String
'
'    GetSampleInfoR = -1
'
'    '-- ȯ������ ��������
'    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBARCODE))   '���� ���ڵ� ��ȣ
'
'    If sBarcode = "" Then
'        Exit Function
'    End If
'
'    '-- ���ڵ��ȣ�� ȯ������ �ҷ�����
'          SQL = "SELECT " & gDBCOLUMN_Parm.PID & "," & gDBCOLUMN_Parm.PNAME & "," & gDBCOLUMN_Parm.PSEX & "," & gDBCOLUMN_Parm.PAGE & vbCrLf
'    SQL = SQL & "  FROM " & gDBTBL_Parm.ORDTABLE & vbCrLf
'    SQL = SQL & " WHERE " & gDBCOLUMN_Parm.BARCODE & " = '" & sBarcode & "' " & vbCrLf
'    If gDBCOLUMN_Parm.STATUS <> "" Then
'        SQL = SQL + "   AND " & gDBCOLUMN_Parm.STATUS & " = '0' " & vbCrLf
'    End If
'    If gDBCOLUMN_Parm.RESULT <> "" Then
'        SQL = SQL + "   AND (" & gDBCOLUMN_Parm.RESULT & " = '' OR " & gDBCOLUMN_Parm.RESULT & " IS NULL)"
'    End If
'
'    Res = GetDBSelectColumn(gServer, SQL)
'
'    If Res = 1 Then
'        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo
''        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
'        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPNAME
'        '-- ������ ������� �ֹι�ȣ�� ã��
'        'strSex = IIf(Mid(Trim(gReadBuf(4)), 7, 1) = "1", "M", "F")
'        'SetText frmInterface.vasID, strSex, colSex    '7  ����
''        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '7  ����
'        '-- ���̰� ������� �ֹι�ȣ�� ã��
'        'strAge = Format(Now, "yyyy") - Mid(Trim(gReadBuf(3)), 1, 4)
'        'SetText frmInterface.vasID, strAge, asRow, colAge
''        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colSex    '8  ����
'
'        GetSampleInfoR = 1
'    Else
'
'        GetSampleInfoR = -1
'    End If
'
'End Function

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
    Dim RS As ADODB.Recordset
    Dim intCol As Integer

    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
          

     
          SQL = "SELECT P.PbsPatNam, O.OSPCHTNUM, R.ResLabCod, E.LabShtNam, R.ResOcmNum, R.ResOdrSeq, R.ResSeq, R.ResSubSeq, R.ResRltVal" & vbCrLf
    SQL = SQL & "  FROM RsbInf M, ResInf R, ospinf O, PBSINF P, LabMst E" & vbCrLf
    SQL = SQL & " WHERE M.RsbBarCod = '" & argPID & "'" & vbCrLf
    SQL = SQL & "   And M.RsbAckStt <> 'A' " & vbCrLf
    SQL = SQL & "   And O.OspChkStt <> 'F' " & vbCrLf
    SQL = SQL & "   And (R.ResRepTyp Is Null or R.ResRepTyp <> 'F') " & vbCrLf
    SQL = SQL & "   And (R.ResRltVal is Null or R.ResRltVal = '') " & vbCrLf
    SQL = SQL & "   And R.ResStatus <> '5' " & vbCrLf
    SQL = SQL & "   And M.RSBACPNUM = R.ResRsbAcp" & vbCrLf
    SQL = SQL & "   And R.ResOcmNum = O.OspOcmNum" & vbCrLf
    SQL = SQL & "   and R.ResOdrSeq = O.OspOdrSeq" & vbCrLf
    SQL = SQL & "   and R.ResSeq    = O.OspSeq" & vbCrLf
    SQL = SQL & "   and O.OSPCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   and R.ResLabcod = E.LabCod" & vbCrLf
    
    Set RS = cn_Ser.Execute(SQL)
    
    Do Until RS.EOF
        GetOrderExamCode = GetOrderExamCode & "'" & Trim(RS.Fields("EXAMCODE")) & "',"
        
        '-- ȭ�鿡 ǥ��
        With frmInterface
            For intCol = colState + 1 To .vasID.MaxCols
                If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                    .vasID.Row = .vasID.ActiveRow
                    .vasID.Col = intCol
                    .vasID.BackColor = vbYellow
                    Exit For
                End If
            Next
        End With
        
        RS.MoveNext
    Loop
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
    End If
    
    RS.Close
    
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
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    
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
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    
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
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    
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

'-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ���
Function GetGetEquipExamCode_Hitachi7080(argEquipCode As String, argPID As String, Optional intRow As Long) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim strExamCode() As String
    Dim sBarcode     As String
    
    GetGetEquipExamCode_Hitachi7080 = ""
    j = 0
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    
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
    'strExamCode = ""

    'ReDim strExamCode(UBound(gReadBuf))
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            ReDim Preserve strExamCode(j)
            strExamCode(j) = Trim(gReadBuf(i))
            j = j + 1
        Else
            Exit For
        End If
    Next

    GetGetEquipExamCode_Hitachi7080 = strExamCode
    
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


