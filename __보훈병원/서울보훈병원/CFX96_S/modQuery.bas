Attribute VB_Name = "modQuery"
Option Explicit

Public SQL              As String
Public RS               As ADODB.Recordset
Public blnSameRecord    As Boolean



'�� ���ä�ο� �˻��ڵ尡 1���̻� ���� (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_AU680(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_AU680 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                strExamCode = strExamCode & Format(strSendCH, "000")
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    GetEquipExamCode_AU680 = strExamCode

End Function

'�� ���ä�ο� �˻��ڵ尡 1���̻� ���� (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_HITACHI7180(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim intIntBase      As Integer
    Dim strItems        As String           '������ �˻��׸�
    Dim blnISE          As Boolean          'Na, K, Cl �˻翩��

    strItems = String$(88, "0")
    GetEquipExamCode_HITACHI7180 = strItems
    strExamCode = ""
    blnISE = False
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            If IsNumeric(AdoRs_Local.Fields("SENDCHANNEL").Value) Then
                intIntBase = CInt(AdoRs_Local.Fields("SENDCHANNEL").Value)
                If intIntBase <> "" Then
                    '## ����׸�: 93~100
                    If intIntBase >= 93 And intIntBase <= 100 Then
                        'GoTo Skip1
                    Else
                        '## Na, K, Cl �˻翩�� Check
                        If intIntBase = 87 Or intIntBase = 88 Or intIntBase = 89 Then
                            blnISE = True
                        Else
                            Mid$(strItems, intIntBase, 1) = "1"
                        End If
                    End If
                    
                    '## TIBC �̸� UIBC,FE ������ �ش�.
                    'If lngIntBase = 98 Then
                    '    Mid$(strItems, 22, 1) = "1"     'FE
                    '    Mid$(strItems, 23, 1) = "1"     'UIBC
                    'End If
                            
                    '## B/C  (025)�׸��� ����׸��̶� ������ ������ �ȵ�(BUN,CREA)
                    '## A/G  (026)�׸��� ����׸��̶� ������ ������ �ȵ�
                    '## GLOB (032)�׸��� ����׸��̶� ������ ������ �ȵ�
                    '## I.Bil(033)�׸��� ����׸��̶� ������ ������ �ȵ�
                    '## T.P  (002)�׸��� ��ü�� Urine�϶� �˻縦 �ϸ� �ȵ�
                    '## HbA1C(23)�׸��� Hgb(20)�� A1C(21) ������ ������ ��
                    '## LDL-C(99)�׸��� ����׸��̶� ������ ������ �ȵ�(CHOL, T.G, HDL-C)
                    mOrder.SendCnt = mOrder.SendCnt + 1
                End If
            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    '## Na, K, Cl �˻翩�� Check
    If blnISE Then
        Mid$(strItems, 87, 1) = "1"
        mOrder.SendCnt = mOrder.SendCnt + 1
    End If

    AdoRs_Local.Close

    GetEquipExamCode_HITACHI7180 = strItems
    

End Function

Function SaveTransData_EONM(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpCd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
        SaveTransData_EONM = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpCd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMSUBCODE").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- ��������
                    SQL = "" '
                    SQL = SQL & "Update TB_H141_LISTAKEBODY                     " & vbCrLf
                    SQL = SQL & "   SET H141_RSLTYN    ='Y'                     " & vbCrLf
                    SQL = SQL & " WHERE H141_TSAMPLENO = '" & strBarcode & "'   " & vbCrLf
                    SQL = SQL & "   AND H141_SUGACD    = '" & strTestCd & "'    " & vbCrLf
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                    
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H131_SPPRESULT                       " & vbCrLf
                    SQL = SQL & "   SET H131_RESULT  = '" & sResult & "'        " & vbCrLf
                    SQL = SQL & " WHERE H131_SPPTYPE = '" & gHOSP.PARTCD & "'   " & vbCrLf    'L010
                    SQL = SQL & "   AND H131_SEQNO   = '" & strTestCdSub & "'   " & vbCrLf
                        
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H130_SPPRECEIVE                              " & vbCrLf
                    SQL = SQL & "   SET H130_RSLTDAT = TO_CHAR(SYSDATE, 'YYYYMMDD')     " & vbCrLf
                    SQL = SQL & "      ,H130_RSLTTM  = TO_CHAR(SYSDATE, 'HH24:MI:SS')   " & vbCrLf
                    SQL = SQL & " WHERE H130_SPPTYPE = '" & gHOSP.PARTCD & "'           " & vbCrLf    'L010
                    SQL = SQL & "   AND H130_SEQNO   = '" & strTestCdSub & "'           " & vbCrLf
                        
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H140_LISTAKEHEAD                     " & vbCrLf
                    SQL = SQL & "   SET H140_RSLTYN    = 'Y'                    " & vbCrLf
                    SQL = SQL & " WHERE H140_TSAMPLENO = '" & strBarcode & "'   " & vbCrLf
                                        
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_EONM = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_EONM = -1
    
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal
    
End Function

'Function SaveTransData_NU(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
'    Dim RsLocal         As ADODB.Recordset
'
'    Dim strSaveSeq      As String
'    Dim strExamDate     As String
'    Dim strHospDate     As String
'    Dim strBarcode      As String
'    Dim strChartNo      As String
'    Dim strPatID        As String
'    Dim strPatNm        As String
'
'    Dim strEqpCd        As String
'    Dim strOrdCd        As String
'    Dim strTestCd       As String
'    Dim strTestCdSub    As String
'    Dim sResult         As String
'    Dim sResult1        As String
'    Dim sResult2        As String
'    Dim strJudge        As String
'
'    Dim sParam          As String
'    Dim strAllResult    As String
'    Dim strDate         As String
'    Dim sRcvData        As String
'
'On Error GoTo ErrHandle
'
'    strJudge = ""
'    sResult = ""
'    sResult1 = ""
'    sResult2 = ""
'    strAllResult = ""
'    sRcvData = ""
'
'    With frmMain
'        SaveTransData_NU = -1
'
'        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
'        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
'        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
'        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
'        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
'        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
'        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
'
'        If Trim(strBarcode) = "" Then
'            Exit Function
'        End If
'
'        If Trim(strPatNm) = "" Then
'            Exit Function
'        End If
'
'        '-- Local���� ȯ�ں��� ����� ��������
'              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
'        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
'        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
'        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
'        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
'        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
'
'        Set RsLocal = New ADODB.Recordset
'        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
'        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
'            Do Until RsLocal.EOF
'                strEqpCd = RsLocal.Fields("EQUIPCODE").Value & ""
'                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
'                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
'                strTestCdSub = RsLocal.Fields("EXAMSUBCODE").Value & ""
'                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
'                sResult2 = RsLocal.Fields("RESULT").Value & ""
'                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
'
'                '-- ���������
'                If gHOSP.SAVELIS = "Y" Then
'                    sResult = sResult2
'                Else
'                    sResult = sResult1
'                End If
'
'                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
'                    strAllResult = strAllResult & strTestCd & "" & sResult & "" & strDate & "" & "1" & ""
'                End If
'                RsLocal.MoveNext
'            Loop
'        End If
'
'        RsLocal.Close
'
'        If strAllResult <> "" Then
'            sParam = ""
'            sParam = sParam & "submit_id=TXLII00101&"
'            sParam = sParam & "business_id=li&"
'            sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"     '�����ID|����ڵ�
'            sParam = sParam & "bcno=" & strBarcode & "&"                                    '��ü��ȣ(���ڵ�)
'            sParam = sParam & "result=" & strAllResult & "&"                                '���
'            sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '����ڵ�
'            sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '����ڵ�
'            sParam = sParam & "userid=" & gHOSP.USERID & "&"                                '�����ID
'
'            sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'            Call SetSQLData("�������", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'            If InStr(1, sRcvData, "<?xml version") > 0 Then
'                SaveTransData_NU = 1
'            Else
'                SaveTransData_NU = -1
'            End If
'        End If
'
'    End With
'
'Exit Function
'
'ErrHandle:
'    SaveTransData_NU = -1
'
'    Screen.MousePointer = 1
'
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_NU" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show vbModal
'
'End Function

'Function SaveTransData_HDINFO(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
'    Dim RsLocal         As ADODB.Recordset
'
'    Dim strSaveSeq      As String
'    Dim strExamDate     As String
'    Dim strHospDate     As String
'    Dim strBarcode      As String
'    Dim strChartNo      As String
'    Dim strPatID        As String
'    Dim strPatNm        As String
'
'    Dim strEqpCd        As String
'    Dim strOrdCd        As String
'    Dim strTestCd       As String
'    Dim strTestCdSub    As String
'    Dim sResult         As String
'    Dim sResult1        As String
'    Dim sResult2        As String
'    Dim strJudge        As String
'
'    Dim sParam          As String
'    Dim strAllResult    As String
'    Dim strDate         As String
'    Dim sRcvData        As String
'
'On Error GoTo ErrHandle
'
'    strJudge = ""
'    sResult = ""
'    sResult1 = ""
'    sResult2 = ""
'    strAllResult = ""
'    sRcvData = ""
'
'    With frmMain
'        SaveTransData_HDINFO = -1
'
'        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
'        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
'        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
'        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
'        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
'        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
'        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
'
'        If Trim(strBarcode) = "" Then
'            Exit Function
'        End If
'
'        If Trim(strPatNm) = "" Then
'            Exit Function
'        End If
'
'        '-- Local���� ȯ�ں��� ����� ��������
'              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
'        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
'        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
'        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
'        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
'        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
'
'        Set RsLocal = New ADODB.Recordset
'        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
'        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
'            Do Until RsLocal.EOF
'                strEqpCd = RsLocal.Fields("EQUIPCODE").Value & ""
'                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
'                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
'                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
'                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
'                sResult2 = RsLocal.Fields("RESULT").Value & ""
'                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
'
'                '-- ���������
'                If gHOSP.SAVELIS = "Y" Then
'                    sResult = sResult2
'                Else
'                    sResult = sResult1
'                End If
'
''SERVERIP "/himed2/.live?submit_id=" + argId + "&business_id=lis&bcno=" + argBarcode + "&result=" + argResult + "&eqmtcd=" + strLIS_EQCD + "&instcd=053&userid=LISBC&paste=Y&retestyn=N&nmeddilute=0"
'' -> (����IP)/himed2/.live?
''submit_id=TXLII00101&
''business_id=lis&
''bcno=(���ڵ��ȣ)&
''result=(���:�˻��ڵ�%17���%17%17�Է½ð�%171%03�����������˻��ڵ�%17���%17%17�Է½ð�%171)&eqmtcd=(����ڵ�)&
''instcd=053&
''userid=LISBC&
''paste=Y&
''retestyn=N&
''nmeddilute=0
'
''JC�޵���
''http://10.10.10.71/himed2/.live?
''submit_id = TXLII00101&
''business_id = lis&
''bcno=8285800190&
''result=             VB8506B18 %17 N %17%17 20191030142131 %171%03 VB8506B17%17N%17%1720191030142131%171%03VB8506B16%17N%17%1720191030142131%171%03VB8506B15%17N%17%1720191030142131%171%03VB8506B14%17N%17%1720191030142131%171%03VB8506B13%17N%17%1720191030142131%171%03VB8506B12%17N%17%1720191030142131%171%03VB8506B11%17N%17%1720191030142131%171%03VB8506B10%17N%17%1720191030142131%171%03VB8506B09%17N%17%1720191030142131%171%03VB8506B08%17N%17%1720191030142131%171%03VB8506B07%17N%17%1720191030142131%171%03VB8506B06%17N%17%1720191030142131%171%03VB8506B05%17N%17%1720191030142131%171%03VB8506B04%17N%17%1720191030142131%171%03VB8506B03%17N%17%1720191030142131%171%03VB8506B02%17N%17%1720191030142131%171%03VB8506B01%17N%17%1720191030142131%171%03VB8506B19%17N%17%1720191030142131%171&
''eqmtcd=008&
''instcd=053&
''userid=LISBC&
''paste=Y&
''retestyn=N&
''nmeddilute=0
'
'                strDate = Format(Now, "yyyymmddhhmmss")
'
'                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
'                    'strAllResult = strAllResult & strTestCd & "" & sResult & "" & strDate & "" & "1" & ""
'                    strAllResult = strAllResult & strTestCd & "%17" & sResult & "%17%17" & strDate & "%17" & "1" & "%03"
'                End If
'                RsLocal.MoveNext
'            Loop
'        End If
'
'        RsLocal.Close
'
'        If strAllResult <> "" Then
'            sParam = ""
'            sParam = sParam & "submit_id=TXLII00101&"
'            sParam = sParam & "business_id=lis&"
''            sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"     '�����ID|����ڵ�
'            sParam = sParam & "bcno=" & strBarcode & "&"                                    '��ü��ȣ(���ڵ�)
'            sParam = sParam & "result=" & strAllResult & "&"                                '���
'            sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '����ڵ�
'            sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '����ڵ�
'            sParam = sParam & "userid=" & gHOSP.USERID & "&"                                '�����ID
'            sParam = sParam & "paste=Y&"
'            sParam = sParam & "retestyn=N&"
'            sParam = sParam & "nmeddilute=0"
'
'            sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'            Call SetSQLData("�������", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'            If InStr(1, sRcvData, "<?xml version") > 0 Then
'                SaveTransData_HDINFO = 1
'            Else
'                SaveTransData_HDINFO = -1
'            End If
'        End If
'
'    End With
'
'Exit Function
'
'ErrHandle:
'    SaveTransData_HDINFO = -1
'
'    Screen.MousePointer = 1
'
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_HDINFO" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show vbModal
'
'End Function

Function SaveTransData_SWMC(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    Dim strSpcCd        As String
    Dim strEqpCd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim sParam          As String
    Dim strAllResult    As String
    Dim strDate         As String
    Dim sRcvData        As String
    Dim strCmnt         As String
    
    Dim Param01         As New ADODB.Parameter
    Dim Param02         As New ADODB.Parameter
    Dim Param03         As New ADODB.Parameter
    Dim Param04         As New ADODB.Parameter
    Dim Param05         As New ADODB.Parameter
    Dim Param06         As New ADODB.Parameter
    Dim Param07         As New ADODB.Parameter
    Dim Param08         As New ADODB.Parameter
    Dim Param09         As New ADODB.Parameter
    Dim Param10         As New ADODB.Parameter
    Dim Param11         As New ADODB.Parameter
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    strAllResult = ""
    sRcvData = ""
    strCmnt = "�ǽð� ������ ����ȿ�ҿ��������(Realtime reverse transcriptase PCR)"
    
    With frmMain
        SaveTransData_SWMC = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        strSpcCd = Trim(GetText(SPD, argSpcRow, colSPECIMEN))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpCd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If

                If sResult <> "" Then
                    '-- �˻������� = PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEMGT_U02
                    Set AdoCmd = New ADODB.Command
                    Set AdoCmd.ActiveConnection = AdoCn
                    With AdoCmd
                        .CommandTimeout = 15 'MEDI.
                        .CommandText = "PR_CPL_CPL0891_INSERT"  'MEDI.PR_CPL_CPL0891_INSERT
                        .CommandType = adCmdStoredProc
                        
                        Set Param01 = .CreateParameter("I_JANGBI_NAME", adVarChar, adParamInput, 40, gHOSP.MACHCD)  'EquipCode=M07
                        .Parameters.Append Param01
                        Set Param02 = .CreateParameter("I_SAMPLE_ID", adVarChar, adParamInput, 20, strBarcode)
                        .Parameters.Append Param02
                        Set Param03 = .CreateParameter("I_HANGMOG_CODE", adVarChar, adParamInput, 20, strTestCd)
                        .Parameters.Append Param03
                        Set Param04 = .CreateParameter("I_CPL_RESULT", adVarChar, adParamInput, 50, sResult)
                        .Parameters.Append Param04
                        Set Param05 = .CreateParameter("I_CHK_FLAG", adVarChar, adParamInput, 1, "N")
                        .Parameters.Append Param05
                        Set Param06 = .CreateParameter("I_CONFIRM_YN", adVarChar, adParamInput, 1, "")
                        .Parameters.Append Param06
                        Set Param07 = .CreateParameter("I_FKCPL0201", adVarChar, adParamInput, 50, strPatID)
                        .Parameters.Append Param07
                        Set Param08 = .CreateParameter("I_SPECIMEN_CODE", adVarChar, adParamInput, 2, strSpcCd)
                        .Parameters.Append Param08
                        Set Param09 = .CreateParameter("I_EMERGENCY", adVarChar, adParamInput, 1, "N")
                        .Parameters.Append Param09
                        Set Param10 = .CreateParameter("I_JANGBI_RESULT", adVarChar, adParamInput, 50, sResult)
                        .Parameters.Append Param10
                        Set Param11 = .CreateParameter("I_JANGBI_FLAG", adVarChar, adParamInput, 2000, strCmnt)
                        .Parameters.Append Param11
                         
                        SetRawData "[�������]" & gHOSP.MACHCD & "," & strBarcode & "," & strTestCd & "," & sResult & ",N, null" & "," & strPatID & "," & strSpcCd & "N" & "," & sResult & "," & strCmnt
                        .Execute
                        Set AdoCmd = Nothing
                    End With
                End If
                
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
    
    End With

Exit Function

ErrHandle:
    SaveTransData_SWMC = -1
    
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_HDINFO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal
    
End Function


Function SaveTransData_VHS(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    Dim strSpcCd        As String
    Dim strEqpCd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim sParam          As String
    Dim strAllResult    As String
    Dim strDate         As String
    Dim sRcvData        As String
    Dim strCmnt         As String
    Dim strPCmnt        As String
    Dim strHeader       As String
    Dim strFooter       As String
    
    Dim Param01         As New ADODB.Parameter
    Dim Param02         As New ADODB.Parameter
    Dim Param03         As New ADODB.Parameter
    Dim Param04         As New ADODB.Parameter
    Dim Param05         As New ADODB.Parameter
    Dim Param06         As New ADODB.Parameter
    Dim Param07         As New ADODB.Parameter
    Dim Param08         As New ADODB.Parameter
    Dim Param09         As New ADODB.Parameter
    Dim Param10         As New ADODB.Parameter
    Dim Param11         As New ADODB.Parameter
    
    Dim strErrYN        As String
    Dim strErrMsg       As String
    
    Dim strEVal         As String
    Dim strRVal         As String
    Dim strNVal         As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    strAllResult = ""
    sRcvData = ""
    
    'strCmnt = "�ǽð� ������ ����ȿ�ҿ��������(Realtime reverse transcriptase PCR)"
        
    strHeader = ""
    strFooter = ""
    strCmnt = ""
    strCmnt = strCmnt & " �� Test method         Real-time RT-PCR" & vbCrLf
    strCmnt = strCmnt & " �� Reference range     Negative" & vbCrLf
    strCmnt = strCmnt & " �� Comment" & vbCrLf
    strCmnt = strCmnt & "    ��PCR �˻�� ��ü �� �ռ��� ���ų� �������� ��ü��,  �������������� " & vbCrLf
    strCmnt = strCmnt & "      �����ϴ� ��� �������� ���ü� �ֽ��ϴ�." & vbCrLf
    strCmnt = strCmnt & "     ����, PCR �˻�� ������ ������ �˻��ϹǷ� �����հ� ����� ������ �ȵǾ� " & vbCrLf
    strCmnt = strCmnt & "      ���缺�� ���ɼ��� �ֽ��ϴ�." & vbCrLf
    strCmnt = strCmnt & "     ��� �ؼ��� ȯ���� �ӻ� ����� ����Ͽ� �Ǵ��Ͻñ� �ٶ��ϴ�." & vbCrLf
    strCmnt = strCmnt & "    " & vbCrLf
    strCmnt = strCmnt & "    " & vbCrLf
    'strCmnt = strCmnt & "�˻���:������ M.T./������:������ M.D (tel.3943) "
        
    'strCmnt = strCmnt & "Positive "
    'strCmnt = strCmnt & ""
    'strCmnt = strCmnt & ""
    'strCmnt = strCmnt & " �� Test method         Real-time RT-PCR"
    'strCmnt = strCmnt & " �� Reference range     Negative"
    'strCmnt = strCmnt & " �� Comment"
    'strCmnt = strCmnt & "    ��PCR �˻�� ��ü �� �ռ��� ���ų� �������� ��ü��,  �������������� "
    'strCmnt = strCmnt & "      �����ϴ� ��� �������� ���ü� �ֽ��ϴ�."
    'strCmnt = strCmnt & "     ����, PCR �˻�� ������ ������ �˻��ϹǷ� �����հ� ����� ������ �ȵǾ� "
    'strCmnt = strCmnt & "      ���缺�� ���ɼ��� �ֽ��ϴ�."
    'strCmnt = strCmnt & "     ��� �ؼ��� ȯ���� �ӻ� ����� ����Ͽ� �Ǵ��Ͻñ� �ٶ��ϴ�."
    'strCmnt = strCmnt & "    "
    'strCmnt = strCmnt & "    "
    'strCmnt = strCmnt & "�˻���:������ M.T./������:������ M.D (tel.3943) "
        
    ''Positive
    ''
    ''�Ƿ��Ͻ� ��ü���� �ڷγ�19 (��⵵) �缺 ������ ����Ǿ����ϴ�.
    ''
    ''
    ''�ڷγ�19:
    '' [�������� ���� �� ������ ���� ���� �����Ģ] �� ������ ���� ������ ���ܱ��ؿ� �ǰ��Ͽ� �ش� �Ƿ��� ���� ���Ǽҿ� �Ű���� �����Ͻñ�  �ٶ��ϴ�.
    ''
    strPCmnt = ""
    strPCmnt = strPCmnt & " �� Test method           Real-time RT-PCR" & vbCrLf
    strPCmnt = strPCmnt & " �� Reference range       Negative" & vbCrLf
    strPCmnt = strPCmnt & " �� Comment" & vbCrLf
    strPCmnt = strPCmnt & "    ��PCR �˻�� ��ü �� �ռ��� ���ų� �������� ��ü �� �Ǵ� ���� ����������" & vbCrLf
    strPCmnt = strPCmnt & "      �����ϴ� ���  �������� ���ü� �ֽ��ϴ�. " & vbCrLf
    strPCmnt = strPCmnt & "     ����, PCR �˻�� ������ ������ �˻� �ϹǷ� �����հ� ����� ������ �ȵǾ� " & vbCrLf
    strPCmnt = strPCmnt & "      ���缺�� ���ɼ��� �ֽ��ϴ�." & vbCrLf
    strPCmnt = strPCmnt & "     ��� �ؼ� ��, ȯ���� �ӻ� ���� �������� �Ǵ��Ͻñ� �ٶ��ϴ�." & vbCrLf
    strPCmnt = strPCmnt & "    ���ľ�ó���� ��޻�� ������ǰ�� �̿���  �˻��Դϴ�." & vbCrLf
    strPCmnt = strPCmnt & "    ���ڷγ�19(2019-nCoV)�� ���� �����������ο��� ���������� ��1�� ����������" & vbCrLf
    strPCmnt = strPCmnt & "      �����ϴ� ��ȯ���� �ڷγ�19 �����ڰ˻�(PCR)����� Ȯ���� �Ǹ�," & vbCrLf
    strPCmnt = strPCmnt & "     ������������ �Ű� ����Դϴ�." & vbCrLf
    strPCmnt = strPCmnt & "    " & vbCrLf
    'strPCmnt = strPCmnt & "<Corona CT>"
    'strPCmnt = strPCmnt & "E Gene:"
    'strPCmnt = strPCmnt & "RdRP/S Gene :"
    'strPCmnt = strPCmnt & "N Gene:"
    'strPCmnt = strPCmnt & "    " & vbCrLf
    'strPCmnt = strPCmnt & "    " & vbCrLf
    
    ''�˻���:������ M.T./������:������ M.D (tel.3943)
    ''
    strFooter = ""
    strFooter = "�˻���:" & frmMain.txtTestNm.Text & " M.T./������:������ M.D "
            
    With frmMain
        SaveTransData_VHS = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        strSpcCd = Trim(GetText(SPD, argSpcRow, colSPECIMEN))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        SQL = SQL & " ORDER BY EXAMCODE DESC"
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpCd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strEqpCd = "E" Then
                    strEVal = ""
                    If IsNumeric(sResult) And sResult <> "N/A" Then
                        strEVal = sResult
                    End If
                ElseIf strEqpCd = "RdRp" Then
                    strRVal = ""
                    If IsNumeric(sResult) And sResult <> "N/A" Then
                        strRVal = sResult
                    End If
                ElseIf strEqpCd = "N" Then
                    strNVal = ""
                    If IsNumeric(sResult) And sResult <> "N/A" Then
                        strNVal = sResult
                    End If
                Else
                    If UCase(sResult) = "NEGATIVE" Then
                        strHeader = ""
                        strHeader = strHeader & "Negative" & vbCrLf
                        strHeader = strHeader & "" & vbCrLf
                        strHeader = strHeader & "" & vbCrLf
                        strHeader = strHeader & "�ڷγ�19:"
                        strHeader = strHeader & " [�������� ���� �� ������ ���� ���� �����Ģ] �� ������ ���� ������ ���ܱ��ؿ� �ǰ��Ͽ� �ش� �Ƿ��� ���� ���Ǽҿ� �Ű���� �����Ͻñ�  �ٶ��ϴ�." & vbCrLf
            
                        strCmnt = strHeader & strCmnt & strFooter
                   
                    ElseIf UCase(sResult) = "POSITIVE" Then
                        strHeader = ""
                        strHeader = strHeader & "Positive" & vbCrLf
                        strHeader = strHeader & "" & "�Ƿ��Ͻ� ��ü���� �ڷγ�19 (��⵵) �缺 ������ ����Ǿ����ϴ�." & vbCrLf
                        strHeader = strHeader & "" & vbCrLf
                        strHeader = strHeader & "�ڷγ�19:"
                        strHeader = strHeader & " [�������� ���� �� ������ ���� ���� �����Ģ] �� ������ ���� ������ ���ܱ��ؿ� �ǰ��Ͽ� �ش� �Ƿ��� ���� ���Ǽҿ� �Ű���� �����Ͻñ�  �ٶ��ϴ�." & vbCrLf
                        
                        strPCmnt = strPCmnt & "<Corona CT>" & vbCrLf
                        strPCmnt = strPCmnt & "E Gene:" & strEVal & vbCrLf
                        strPCmnt = strPCmnt & "RdRP/S Gene :" & strRVal & vbCrLf
                        strPCmnt = strPCmnt & "N Gene:" & strNVal & vbCrLf
                        strPCmnt = strPCmnt & "    " & vbCrLf
                        strPCmnt = strPCmnt & "    " & vbCrLf
                        
                        strCmnt = strHeader & strPCmnt & strFooter
                   
                   End If
                                   
                   If sResult <> "" Then
                       '-- �˻�������
                       Set AdoCmd = New ADODB.Command
                       Set AdoCmd.ActiveConnection = AdoCn
                       With AdoCmd
                           .CommandTimeout = 15 'MEDI.
                           .CommandText = "PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEMGT_U02"
                           .CommandType = adCmdStoredProc
                           
                           Set Param01 = .CreateParameter("IN_SPCNO", adVarChar, adParamInput, 20, strBarcode)
                           .Parameters.Append Param01
                           
                           Set Param02 = .CreateParameter("IN_EXAMCD", adVarChar, adParamInput, 20, strTestCd)
                           .Parameters.Append Param02
                           
                           If Mid(strTestCd, 1, 2) = "L1" Then '��纴��
                               Set Param03 = .CreateParameter("IN_RESULT", adVarChar, adParamInput, 4000, sResult)
                           Else
                               Set Param03 = .CreateParameter("IN_RESULT", adVarChar, adParamInput, 4000, strCmnt)
                           End If
                           
                           .Parameters.Append Param03
                           
                           Set Param04 = .CreateParameter("IN_USERID", adVarChar, adParamInput, 100, frmMain.txtTestID.Text)
                           .Parameters.Append Param04
                           
                           Set Param05 = .CreateParameter("IN_RFLAG", adVarChar, adParamInput, 50, "D")            '(����: "C", �Է�: "D", ���: "N")
                           .Parameters.Append Param05
                           
                           Set Param06 = .CreateParameter("IN_EQPCD", adVarChar, adParamInput, 100, gHOSP.MACHCD)
                           .Parameters.Append Param06
                           
                           Set Param07 = .CreateParameter("IN_IMGPATH", adVarChar, adParamInput, 100, "")
                           .Parameters.Append Param07
                           
                           Set Param08 = .CreateParameter("IN_ACPNO", adVarChar, adParamInputOutput, 10, 0)
                           .Parameters.Append Param08
                           
                           Set Param09 = .CreateParameter("IN_ERRYN", adVarChar, adParamInputOutput, 100, strErrYN)
                           .Parameters.Append Param09
                           
                           Set Param10 = .CreateParameter("IN_ERRMSG", adVarChar, adParamInputOutput, 100, strErrMsg)
                           .Parameters.Append Param10
                            
                           If Mid(strTestCd, 1, 2) = "L1" Then '��纴��
                               Call SetSQLData("�������", strBarcode & "," & strTestCd & "," & strCmnt & "," & frmMain.txtTestID.Text & ",D" & "," & gHOSP.MACHCD & ","",0," & strErrYN & "," & strErrMsg, "A")
                           Else
                               Call SetSQLData("�������", strBarcode & "," & strTestCd & "," & sResult & "," & frmMain.txtTestID.Text & ",D" & "," & gHOSP.MACHCD & ","",0," & strErrYN & "," & strErrMsg, "A")
                           End If
                           
                           .Execute
                           Set AdoCmd = Nothing
                       End With
                   End If
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
    
    
        If Mid(strTestCd, 1, 2) = "L1" Then '��纴��
            SQL = ""
            SQL = SQL & "UPDATE SLA1COLMT                     " & vbCrLf
            SQL = SQL & "   SET RMK         = '" & strCmnt & "'   " & vbCrLf
            SQL = SQL & " WHERE SPC_NO      = '" & strBarcode & "'" & vbCrLf
            SQL = SQL & "   AND PT_NO       = '" & strPatID & "'  " & vbCrLf
'            SQL = SQL & "   AND EXAM_GRPCD  = '" & strPatID & "'  " & vbCrLf
            SQL = SQL & "   AND ORD_DATE    = '" & Mid(strHospDate, 1, 10) & "'  " & vbCrLf
                                
            Call SetSQLData("�������", SQL, "A")
            AdoCn.Execute SQL
        End If
        
    End With
    
    SaveTransData_VHS = 1

Exit Function

ErrHandle:
    SaveTransData_VHS = -1
    
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_VHS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal
    
End Function


Function SaveTransData_EHWA(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strSpcmCd       As String
    Dim strEqpCd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim sParam          As String
    Dim strAllResult    As String
    Dim strDate         As String
    Dim sRcvData        As String
    Dim strCmnt         As String
    Dim strHospGbn      As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    strAllResult = ""
    sRcvData = ""
    strCmnt = ""
    
    With frmMain
        SaveTransData_EHWA = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        strSpcmCd = Trim(GetText(SPD, argSpcRow, colSPECIMEN))

        'MsgBox strSpcmCd
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
'        MsgBox SQL
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            
'            MsgBox RsLocal.RecordCount
            
            Do Until RsLocal.EOF
                strEqpCd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    sParam = "                    "
                    sParam = sParam & "<Table>"
                    sParam = sParam & "<QID><![CDATA[PKG_MSE_LM_INTERFACE.PC_MSE_INTERFACE_SAVE]]></QID>"
                    sParam = sParam & "<QTYPE><![CDATA[Package]]></QTYPE>"
                    sParam = sParam & "<USERID><![CDATA[LIA]]></USERID>"
                    sParam = sParam & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
                    sParam = sParam & "<TABLENAME><![CDATA[]]></TABLENAME>"
                    
                    gHospCode = "02"
                    
                    '���ﺴ���� ���ڵ尡 ���Ϸ� ���۵ǰ�, �񵿺����� ���ڵ尡 ����Ϸ� ���۵ȴ�. �񵿹��ڵ�� ������ 13 �̻��̴�!
                    If Len(strBarcode) = 11 And IsNumeric(strBarcode) Then
                        strHospGbn = Mid(strBarcode, 1, 2)
                        If CCur(strHospGbn) > 12 Then
                            gHospCode = "02"      '�̴�񵿺���
                        Else
                            gHospCode = "01"      '�̴뼭�ﺴ��
                        End If
                    End If
    
                    sParam = sParam & "<P0><![CDATA[" & gHospCode & "]]></P0>"
                    sParam = sParam & "<P1><![CDATA[" & gHOSP.MACHCD & "]]></P1>"
                    sParam = sParam & "<P2><![CDATA[]]></P2>"
                    sParam = sParam & "<P3><![CDATA[" & gHOSP.USERID & "]]></P3>"
                    sParam = sParam & "<P4><![CDATA[" & gHOSP.MACHNM & "]]></P4>"
                    sParam = sParam & "<P5><![CDATA[]]></P5>"
                    sParam = sParam & "<P6><![CDATA[" & strExamDate & "]]></P6>"
                    sParam = sParam & "<P7><![CDATA[" & strBarcode & "]]></P7>"
                    sParam = sParam & "<P8><![CDATA[]]></P8>"
                    sParam = sParam & "<P9><![CDATA[1]]></P9>"
                    sParam = sParam & "<P10><![CDATA[" & vbTab & strTestCd & vbTab & "]]></P10>"
                    sParam = sParam & "<P11><![CDATA[" & vbTab & "" & vbTab & "]]></P11>"
                    sParam = sParam & "<P12><![CDATA[" & vbTab & sResult & vbTab & "]]></P12>"
                    sParam = sParam & "<P13><![CDATA[" & vbTab & "" & vbTab & "]]></P13>"
                    sParam = sParam & "<P14><![CDATA[" & vbTab & "" & vbTab & "]]></P14>"
                    strCmnt = ""
'                    If UCase(sResult) = "POSITIVE" Then
'                        Select Case strEqpCd
'                        Case "BP":  strCmnt = gCmnt.BPCmnt
'                        Case "CP":  strCmnt = gCmnt.CPCmnt
'                        Case "LP":  strCmnt = gCmnt.LPCmnt
'                        Case "MP":  strCmnt = gCmnt.MPCmnt
'                        End Select
'
'                        strCmnt = Replace(strCmnt, "*Specimen : ", "*Specimen : " & strSpcmCd)
'                        sParam = sParam & "<P15><![CDATA[" & vbTab & strCmnt & vbTab & "]]></P15>"   '�Ұִ߳´�.
'                    Else
'                        Select Case strEqpCd
'                        Case "BP":  strCmnt = gCmnt.BPNCmnt
'                        Case "CP":  strCmnt = gCmnt.CPNCmnt
'                        Case "LP":  strCmnt = gCmnt.LPNCmnt
'                        Case "MP":  strCmnt = gCmnt.MPNCmnt
'                        End Select
'                        strCmnt = Replace(strCmnt, "*Specimen : ", "*Specimen : " & strSpcmCd)
'                        sParam = sParam & "<P15><![CDATA[" & vbTab & strCmnt & vbTab & "]]></P15>"   '�Ұִ߳´�.
'                    End If
                    
'                    If UCase(sResult) = gCmnt.NEG Then
'                        Select Case strEqpCd
'                            Case "TV":  strCmnt = gCmnt.TVNCmnt
'                            Case "MH":  strCmnt = gCmnt.MHNCmnt
'                            Case "UU":  strCmnt = gCmnt.UUNCmnt
'                            Case "CT":  strCmnt = gCmnt.CTNCmnt
'                            Case "MG":  strCmnt = gCmnt.MGNCmnt
'                            Case "NG":  strCmnt = gCmnt.NGNCmnt
'                            Case "UP":  strCmnt = gCmnt.UPNCmnt
'                        End Select
'                        strCmnt = Replace(strCmnt, "*Specimen : ", "*Specimen : " & strSpcmCd)
'                        sParam = sParam & "<P15><![CDATA[" & vbTab & strCmnt & vbTab & "]]></P15>"   '�Ұִ߳´�.
'                    Else
'                        Select Case strEqpCd
'                            Case "TV":  strCmnt = gCmnt.TVCmnt
'                            Case "MH":  strCmnt = gCmnt.MHCmnt
'                            Case "UU":  strCmnt = gCmnt.UUCmnt
'                            Case "CT":  strCmnt = gCmnt.CTCmnt
'                            Case "MG":  strCmnt = gCmnt.MGCmnt
'                            Case "NG":  strCmnt = gCmnt.NGCmnt
'                            Case "UP":  strCmnt = gCmnt.UPCmnt
'                        End Select
'
'                        strCmnt = Replace(strCmnt, "*Specimen : ", "*Specimen : " & strSpcmCd)
'                        sParam = sParam & "<P15><![CDATA[" & vbTab & strCmnt & vbTab & "]]></P15>"   '�Ұִ߳´�.
'                    End If
                    
                    
                    sParam = sParam & "<P16><![CDATA[]]></P16>"
                    sParam = sParam & "<P17><![CDATA[" & vbTab & "" & vbTab & "]]></P17>"
                    sParam = sParam & "<P18><![CDATA[" & vbTab & "" & vbTab & "]]></P18>"
                    sParam = sParam & "</Table>"
                End If
                
                If sParam <> "" Then
                    sParam = "<Row>" & sParam & "</Row>"
                    sParam = "<?xml version='1.0' encoding='euc-kr'?>" & sParam
        
                    Online_Result_Qry sParam
                    
                    SaveTransData_EHWA = 1
                End If
                
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        
    End With

Exit Function

ErrHandle:
    SaveTransData_EHWA = -1
    
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_EHWA" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal
    
End Function

'-- �˻縶���� ��ȸ
Public Sub GetTestList()
    Dim intRow          As Long

    intRow = 0
    gAllTestCd = ""
    Erase gArrEQP

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  SEQNO,SENDCHANNEL,RSLTCHANNEL,TESTCODE,TESTNAME,ABBRNAME " & vbCrLf
    SQL = SQL & " ,RESPRECUSE,RESPREC,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH  " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER " & vbCrLf
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCrLf
    '�˻�ó���� ���� ����ϰ�� ó���ɷ� �������� �ϱ� ���ؼ�...
    SQL = SQL & " ORDER BY SEQNO ASC, TESTCODE DESC, TESTNAME "

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        'ó���ڵ�, SUB�ڵ�� �߰� 16,17
        ReDim Preserve gArrEQP(AdoRs_Local.RecordCount, 17)

        Do Until AdoRs_Local.EOF
            intRow = intRow + 1
            'Debug.Print AdoRs_Local.Fields("SEQNO").Value & "|" & AdoRs_Local.Fields("TESTCODE").Value & ""
            gArrEQP(intRow, 1) = AdoRs_Local.Fields("SEQNO").Value & ""
            gArrEQP(intRow, 2) = AdoRs_Local.Fields("TESTCODE").Value & ""
            gArrEQP(intRow, 3) = AdoRs_Local.Fields("SENDCHANNEL").Value & ""
            gArrEQP(intRow, 4) = AdoRs_Local.Fields("RSLTCHANNEL").Value & ""
            gArrEQP(intRow, 5) = AdoRs_Local.Fields("TESTNAME").Value & ""
            gArrEQP(intRow, 6) = AdoRs_Local.Fields("ABBRNAME").Value & ""
            gArrEQP(intRow, 7) = AdoRs_Local.Fields("RESPRECUSE").Value & ""
            gArrEQP(intRow, 8) = AdoRs_Local.Fields("RESPREC").Value & ""
            gArrEQP(intRow, 9) = AdoRs_Local.Fields("REFMLOW").Value & ""
            gArrEQP(intRow, 10) = AdoRs_Local.Fields("REFMHIGH").Value & ""
            gArrEQP(intRow, 11) = AdoRs_Local.Fields("REFFLOW").Value & ""
            gArrEQP(intRow, 12) = AdoRs_Local.Fields("REFFHIGH").Value & ""
            gArrEQP(intRow, 16) = ""    'ó���ڵ�� ���
            gArrEQP(intRow, 17) = ""    'SUB�ڵ�� ���

            If Trim(AdoRs_Local.Fields("TESTCODE").Value) <> "" Then
                If intRow = 1 Or gAllTestCd = "" Then
                    gAllTestCd = "'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                Else
                    gAllTestCd = gAllTestCd & ",'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                End If
            End If

            AdoRs_Local.MoveNext
        Loop
    End If

End Sub

Public Sub GetTestCode_Name()
    Dim strTestCode     As String
    Dim strTestName     As String
    Dim strTestAbNM     As String
    Dim intRow          As Integer
    
    frmMain.spdCodeName.MaxRows = 0
    frmMain.spdCodeName.MaxCols = 3
    intRow = 0

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  TESTCODE,TESTNAME,ABBRNAME           " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER                       " & vbCrLf
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCrLf
    SQL = SQL & "   AND TESTCODE IS NOT NULL            " & vbCrLf
    SQL = SQL & " ORDER BY TESTCODE,TESTNAME,ABBRNAME   " & vbCrLf

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            intRow = intRow + 1
            
            strTestCode = AdoRs_Local.Fields("TESTCODE").Value & ""
            strTestName = AdoRs_Local.Fields("TESTNAME").Value & ""
            strTestAbNM = AdoRs_Local.Fields("ABBRNAME").Value & ""
            
            With frmMain.spdCodeName
                .MaxRows = AdoRs_Local.RecordCount
                .Row = intRow
                
                Call SetText(frmMain.spdCodeName, strTestCode, intRow, 1)
                Call SetText(frmMain.spdCodeName, strTestName, intRow, 2)
                Call SetText(frmMain.spdCodeName, strTestAbNM, intRow, 3)
                        
            End With
            AdoRs_Local.MoveNext
        Loop
    End If

End Sub


'-- �˻縶���͸� ��ȸ
Public Sub GetTestListName()
    Dim intRow          As Long

    intRow = 0
    Erase gArrEQPNm

    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO,SENDCHANNEL,RSLTCHANNEL,ABBRNAME " & vbCrLf
'    SQL = SQL & " ,RESPRECUSE,RESPREC,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        
        ReDim Preserve gArrEQPNm(AdoRs_Local.RecordCount, 12)

        Do Until AdoRs_Local.EOF
            intRow = intRow + 1
            gArrEQPNm(intRow, 1) = AdoRs_Local.Fields("SEQNO").Value & ""
            gArrEQPNm(intRow, 2) = ""
            gArrEQPNm(intRow, 3) = AdoRs_Local.Fields("SENDCHANNEL").Value & ""
            gArrEQPNm(intRow, 4) = AdoRs_Local.Fields("RSLTCHANNEL").Value & ""
            gArrEQPNm(intRow, 5) = ""
            gArrEQPNm(intRow, 6) = AdoRs_Local.Fields("ABBRNAME").Value & ""
'            gArrEQPNm(intRow, 7) = AdoRs_Local.Fields("RESPRECUSE").Value & ""
'            gArrEQPNm(intRow, 8) = AdoRs_Local.Fields("RESPREC").Value & ""
'            gArrEQPNm(intRow, 9) = AdoRs_Local.Fields("REFMLOW").Value & ""
'            gArrEQPNm(intRow, 10) = AdoRs_Local.Fields("REFMHIGH").Value & ""
'            gArrEQPNm(intRow, 11) = AdoRs_Local.Fields("REFFLOW").Value & ""
'            gArrEQPNm(intRow, 12) = AdoRs_Local.Fields("REFFHIGH").Value & ""

            AdoRs_Local.MoveNext
        Loop
    End If

End Sub


'-- �˻縶���� ��ȸ
Public Sub GetTestMaster(ByVal SPD As Object)
    Dim intRow          As Long

    SPD.MaxRows = 0
    intRow = 0

    SQL = ""
    SQL = SQL & "SELECT * " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With SPD
            .MaxRows = AdoRs_Local.RecordCount '

            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(SPD, AdoRs_Local.Fields("EQUIPCD").Value & "", intRow, colLMACHCODE)
                Call SetText(SPD, AdoRs_Local.Fields("SEQNO").Value & "", intRow, colLSEQNO)
                Call SetText(SPD, AdoRs_Local.Fields("TESTCODE").Value & "", intRow, colLTESTCD)
                Call SetText(SPD, AdoRs_Local.Fields("SENDCHANNEL").Value & "", intRow, colLOCHANNEL)
                Call SetText(SPD, AdoRs_Local.Fields("RSLTCHANNEL").Value & "", intRow, colLRCHANNEL)
                Call SetText(SPD, AdoRs_Local.Fields("TESTNAME").Value & "", intRow, colLTESTNM)
                Call SetText(SPD, AdoRs_Local.Fields("ABBRNAME").Value & "", intRow, colLABBRNM)
                Call SetText(SPD, IIf(AdoRs_Local.Fields("RESPRECUSE").Value & "" = "1", "1", "0"), intRow, colLRESSPECUSE)
                Call SetText(SPD, AdoRs_Local.Fields("RESPREC").Value & "", intRow, colLRESSPEC)
                Call SetText(SPD, AdoRs_Local.Fields("REFMLOW").Value & "", intRow, colLMLOW)
                Call SetText(SPD, AdoRs_Local.Fields("REFMHIGH").Value & "", intRow, colLMHIGH)
                Call SetText(SPD, AdoRs_Local.Fields("REFFLOW").Value & "", intRow, colLFLOW)
                Call SetText(SPD, AdoRs_Local.Fields("REFFHIGH").Value & "", intRow, colLFHIGH)
                

                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 15
        End With
    End If

End Sub


''-- �˻���������� ��ȸ
'Public Sub GetOrderMST()
'    Dim intRow          As Long
'
''    gAllOrdCd = ""
''    intRow = 0
''
''    SQL = ""
''    SQL = SQL & "SELECT ORDERCODE FROM ORDMASTER " & vbCr
''    SQL = SQL & " ORDER BY ORDERCODE "
''
''    '-- Record Count ������
''    AdoCn_Local.CursorLocation = adUseClient
''    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
''    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
''        With frmMain.spdOrdMst
''            .MaxRows = AdoRs_Local.RecordCount
''            Do Until AdoRs_Local.EOF
''                intRow = intRow + 1
''                Call SetText(frmMain.spdOrdMst, AdoRs_Local.Fields("ORDERCODE").Value & "", intRow, 1)
''
''                If Trim(AdoRs_Local.Fields("ORDERCODE").Value) <> "" Then
''                    If intRow = 1 Or gAllTestCd = "" Then
''                        gAllOrdCd = "'" & AdoRs_Local.Fields("ORDERCODE").Value & "'"
''                    Else
''                        gAllOrdCd = gAllOrdCd & ",'" & AdoRs_Local.Fields("ORDERCODE").Value & "'"
''                    End If
''                End If
''
''                AdoRs_Local.MoveNext
''            Loop
''            .RowHeight(-1) = 12
''        End With
''    End If
'End Sub
'

'-- �˻��ڵ�� �˻�� ��ȸ
Public Function GetTestNm(ByVal pItem As String, Optional pFull As Boolean) As String
    Dim intRow          As Long

    GetTestNm = ""

    If pFull = True Then
        SQL = ""
        SQL = SQL & "SELECT TESTNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE TESTCODE = '" & pItem & "'"
    Else
        SQL = ""
        SQL = SQL & "SELECT ABBRNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE TESTCODE = '" & pItem & "'"
    End If

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            GetTestNm = AdoRs_Local.Fields("ITEMNM").Value & ""
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

End Function


'-- �˻��ڵ�� �˻�� ��ȸ
Public Function GetTestNmS(ByVal pItem As String, Optional pFull As Boolean) As String
    Dim strTestNM   As String
    
    strTestNM = ""
    GetTestNmS = ""

    If pFull = True Then
        SQL = ""
        SQL = SQL & "SELECT TESTNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE TESTCODE IN (" & pItem & ")"
    Else
        SQL = ""
        SQL = SQL & "SELECT ABBRNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE TESTCODE IN (" & pItem & ")"
    End If

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strTestNM = strTestNM & AdoRs_Local.Fields("ITEMNM").Value & "/"
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    GetTestNmS = strTestNM

End Function

'
''-- �˻������ ���ä�� ��ȸ
'Public Function GetRsltChannel(ByVal pItem As String) As String
'    Dim RS2             As ADODB.Recordset
'    Dim intRow          As Long
'
'    GetRsltChannel = ""
'
'    SQL = ""
'    SQL = SQL & "SELECT RSLTCHANNEL "
'    SQL = SQL & "  FROM EQPMASTER " & vbCr
'    SQL = SQL & " WHERE ABBRNAME = '" & pItem & "'"
'
'    Set RS2 = New ADODB.Recordset
'
'    '-- Record Count ������
'    AdoCn_Local.CursorLocation = adUseClient
'    Set RS2 = AdoCn_Local.Execute(SQL, , 1)
'    If Not RS2.EOF = True And Not RS2.BOF = True Then
'        Do Until RS2.EOF
'            GetRsltChannel = RS2.Fields("RSLTCHANNEL").Value & ""
'            RS2.MoveNext
'        Loop
'    End If
'
'    RS2.Close
'
'End Function
'
''-- �˻��׸� ��ȸ
'Public Function GetTest(ByVal pTestCd As String) As String
'
'On Error GoTo RST
'    GetTest = ""
'
'    SQL = ""
'    SQL = SQL & "Select ORD_NM "
'    SQL = SQL & "  From LIS_ORD_LIST_V" & vbCr
'    SQL = SQL & " Where ord_cd = '" & pTestCd & "'" & vbCr
'
'    '-- Record Count ������
'    AdoCn.CursorLocation = adUseClient
'    Set RS = AdoCn.Execute(SQL, , 1)
'    If Not RS.EOF = True And Not RS.BOF = True Then
'        Do Until RS.EOF
'            GetTest = Trim(RS.Fields("ORD_NM")) & ""
'            RS.MoveNext
'        Loop
'    End If
'
'    RS.Close
'
'Exit Function
'
'RST:
'
'                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show 'vbModal
'
'    Screen.MousePointer = 0
'
'End Function
'
''-- ��ũ����Ʈ ��ȸ
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)

    Select Case gEMR
        Case "VHS"
                Call GetWorkList_VHS(pFrom, pTo, SPD)
        
        Case "SWMC"                    '�����Ƿ��
                Call GetWorkList_SWMC(pFrom, pTo, SPD)

'        Case "HDINFO"                       '��������
'                Call GetWorkList_HDINFO(pFrom, pTo, SPD)
        
        Case "PHILL"
'                Call GetWorkList_PHILL(pFrom, pTo, SPD)

        Case "MSINFOTEC"                    'MS������
                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)

'        Case "NU"                           '��ȭIS
'                Call GetWorkList_NU(pFrom, pTo, SPD)

'        Case "AMIS"                         '�ƹ̽�
'                Call GetWorkList_AMIS(pFrom, pTo, SPD)
'
'        Case "BIGUBCARE"
'                Call GetWorkList_BIGUBCARE(pFrom, pTo, SPD)
'
'        Case "BIT"                          '��Ʈ
'                Call GetWorkList_BIT(pFrom, pTo, SPD)
'
'        Case "BIT70"                        '��Ʈ HIB70
'                Call GetWorkList_BIT70(pFrom, pTo, SPD)
'
'        Case "EMEDI"                        '�̸޵�
'                Call GetWorkList_AMIS(pFrom, pTo, SPD)
'
'        Case "EASYS"                        '������, MCC
'                Call GetWorkList_EASYS(pFrom, pTo, SPD)

'        Case "EHWA"
'                Call GetWorkList_EHWA(pFrom, pTo, SPD)

'
''        Case "EONM"                         '�̿¿�
''                Call GetWorkList_EONM(pFrom, pTo, SPD)
'
'        Case "GINUS"                         '������
''                Call GetWorkList_GINUS(pFrom, pTo, SPD)
'
'        Case "GSEN"                         '����Ŀ�´����̼���(��íƮ)
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "HWASAN"                       'ȭ��
'                Call GetWorkList_HWASAN(pFrom, pTo, SPD)
'
'        Case "JAINCOM"                      '������
'                Call GetWorkList_JAINCOM(pFrom, pTo, SPD)
'
'        Case "JWINFO"                       '�߿�����
'                Call GetWorkList_JWINFO(pFrom, pTo, SPD)
'
'        Case "KCHART"                       '�ٴ����Ʈ
'                Call GetWorkList_KCHART(pFrom, pTo, SPD)
'
'        Case "KOMAIN"                       '�߿�����
'                Call GetWorkList_KOMAIN(pFrom, pTo, SPD)
'
'        Case "KYU"                          '�Ǿ���б����� - ��ũ����Ʈ ��ɾ���
'                'Call GetWorkList_KYU(pFrom, pTo, SPD)
'
'        Case "MEDICHART"                    '�޵�íƮ
'                Call GetWorkList_MEDICHART(pFrom, pTo, SPD)
'
'        Case "MEDIIT"                       '�޵�IT(���Ƿ����)
'                Call GetWorkList_MEDIIT(pFrom, pTo, SPD)
'
'        Case "MEDITOLISS"                   '�Ƹ�����
'                Call GetWorkList_MEDITOLISS(pFrom, pTo, SPD)
'
'        Case "MCC"                          'MCC SP����
'                Call GetWorkList_MCC(pFrom, pTo, SPD)
'
'        Case "MOD"                          'MOD �ý���
'                Call GetWorkList_MOD(pFrom, pTo, SPD)
'
'        Case "MSINFOTEC"                    'MS������
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "NEOSOFT"                      '�׿�����Ʈ
'                Call GetWorkList_NEOSOFT(pFrom, pTo, SPD)
'
'        Case "ONITGUM"                      '�¾�Ƽ ����
'                Call GetWorkList_ONITGUM(pFrom, pTo, SPD)
'
'        Case "ONITEMR"                      '�¾�Ƽ EMR
'                Call GetWorkList_ONITEMR(pFrom, pTo, SPD)
'
'        Case "PLIS"                         '���̽� ������ó
'                Call GetWorkList_PLIS(pFrom, pTo, SPD)
'
'        Case "SY"                           'SY
'                Call GetWorkList_SY(Format(pFrom, "yyyy-mm-dd"), Format(pTo, "yyyy-mm-dd"), SPD)
'
'        Case "TWIN"                         '��������
'                Call GetWorkList_TWIN(pFrom, pTo, SPD)
'
'        Case "UBCARE"                       '�ǻ��
'                Call GetWorkList_UBCARE(pFrom, pTo, SPD)

'        Case "WELL"                         '��Ŀ�ӽ�
'                Call GetWorkList_WELL(pFrom, pTo, SPD)

'        Case "ONIT"
'            Call GetWorkList_onit(pFrom, pTo, SPD)

'        Case "PLIS"
'            Call GetWorkList_PLIS(pFrom, pTo, SPD)
        Case Else


    End Select

End Sub


Public Sub GetWorkList_PHILL(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo RST

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "   P.request_date  AS HOSPDATE " & vbCrLf
    SQL = SQL & " , P.exam_no       AS PID      " & vbCrLf
    SQL = SQL & " , P.company_code  AS DEPT     " & vbCrLf
    SQL = SQL & " , P.chart_no      AS CHARTNO  " & vbCrLf
    SQL = SQL & " , p.personal_id   AS BARCODE  " & vbCrLf
    SQL = SQL & " , p.person_name   AS PNAME    " & vbCrLf
    SQL = SQL & " , P.worker_code               " & vbCrLf
    SQL = SQL & " , P.patient_kind              " & vbCrLf
    SQL = SQL & " , P.person_sex    AS SEX      " & vbCrLf
    SQL = SQL & " , P.person_age    AS AGE      " & vbCrLf
    SQL = SQL & " , R.pro_code      AS ITEM     " & vbCrLf
    SQL = SQL & "  FROM trust P, trures R       " & vbCrLf
    SQL = SQL & " WHERE P.request_date BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    SQL = SQL & "   AND R.pro_code IN (" & gAllTestCd & ") " & vbCrLf
    SQL = SQL & "   AND R.exam_code <> 'X999' " & vbCrLf
    SQL = SQL & "   AND P.request_date = R.request_date " & vbCrLf
    SQL = SQL & "   AND P.exam_no = R.exam_no " & vbCrLf
    SQL = SQL & " ORDER BY P.request_date, P.exam_no " & vbCrLf

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                strTestCds = strTestCds & "'" & Trim(RS.Fields("ITEM")) & "',"

                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("DEPT")) & "", intRow, colDEPT
                    
                    SetText SPD, GetTestNmS(Mid(strTestCds, 1, Len(strTestCds) - 1)), intRow, colSTATE + 1
                    
                    SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                    frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
'        frmWorkList.chkAll.Value = "1"
    Else
'        frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
'        frmWorkList.chkAll.Value = "0"
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

RST:

End Sub


Public Sub GetWorkList_MSINFOTEC(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

'    SQL = ""
'    SQL = SQL & "SELECT DISTINCT "
'    SQL = SQL & "   P.request_date  AS HOSPDATE " & vbCrLf
'    SQL = SQL & " , P.exam_no       AS PID      " & vbCrLf
'    SQL = SQL & " , P.company_code  AS DEPT     " & vbCrLf
'    SQL = SQL & " , P.chart_no      AS CHARTNO  " & vbCrLf
'    SQL = SQL & " , p.personal_id   AS BARCODE  " & vbCrLf
'    SQL = SQL & " , p.person_name   AS PNAME    " & vbCrLf
'    SQL = SQL & " , P.worker_code               " & vbCrLf
'    SQL = SQL & " , P.patient_kind              " & vbCrLf
'    SQL = SQL & " , P.person_sex    AS SEX      " & vbCrLf
'    SQL = SQL & " , P.person_age    AS AGE      " & vbCrLf
'    SQL = SQL & " , R.pro_code      AS ITEM     " & vbCrLf
'    SQL = SQL & "  FROM trust P, trures R       " & vbCrLf
'    SQL = SQL & " WHERE P.request_date BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
'    SQL = SQL & "   AND R.pro_code IN (" & gAllTestCd & ") " & vbCrLf
'    SQL = SQL & "   AND R.exam_code <> 'X999' " & vbCrLf
'    SQL = SQL & "   AND P.request_date = R.request_date " & vbCrLf
'    SQL = SQL & "   AND P.exam_no = R.exam_no " & vbCrLf
'    SQL = SQL & " ORDER BY P.request_date, P.exam_no " & vbCrLf

    SQL = ""
    SQL = SQL & "Select a.ORDT          AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , a.SPNO          AS BARCODE  " & vbCrLf
    SQL = SQL & "     , a.PAID          AS PID      " & vbCrLf
    SQL = SQL & "     , a.NWNO          AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , b.PANM          AS PNAME    " & vbCrLf
    SQL = SQL & "     , b.SEXS          AS SEX      " & vbCrLf
    SQL = SQL & "     , b.AGES          AS AGE      " & vbCrLf
    SQL = SQL & "     , COUNT(a.ORCD)   AS CNT      " & vbCrLf
    SQL = SQL & "  From emr.LRESULT a, emr.APATINF b        " & vbCrLf
    SQL = SQL & " Where a.ORDT between  '" & pFrom & "' and '" & pTo & "'   " & vbCrLf
    SQL = SQL & "   And a.PAID = b.PAID                                     " & vbCrLf
    SQL = SQL & "   And a.ORCD IN (" & gAllTestCd & ")                      " & vbCrLf
    SQL = SQL & "   And a.OKFL <> 'Y'                                       " & vbCrLf   '-- ���Ȯ������
    'SQL = SQL & "   And a.OKFL = 'N'                                       " & vbCrLf   '-- ���Ȯ������
    'SQL = SQL & "   AND (a.RSFL IS NULL OR a.RSFL = 'N' OR a.RSFL = '')     " & vbCrLf
    SQL = SQL & " GROUP BY a.ORDT,a.SPNO,a.PAID,a.NWNO,b.PANM,b.SEXS,b.AGES " & vbCrLf
    SQL = SQL & " Order By a.ORDT,a.PAID,b.PANM                             " & vbCrLf

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
'                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
'                    Else
'                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
'                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
'                    End If
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
'            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
'            frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
'            frmWorkList.chkAll.Value = "0"
        End If
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_SWMC(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       PART_JUBSU_DATE         AS HOSPDATE " & vbCrLf
'    SQL = SQL & "     , PART_JUBSU_TIME                     " & vbCrLf
    SQL = SQL & "     , SPECIMEN_SER            AS BARCODE  " & vbCrLf
    SQL = SQL & "     , BUNHO                   AS PID      " & vbCrLf
    SQL = SQL & "     , FKCPL0201               AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , SPECIMEN_CODE           AS SPCCD    " & vbCrLf
    SQL = SQL & "     , SUNAME                  AS PNAME    " & vbCrLf
    SQL = SQL & "     , AGE                     AS AGE      " & vbCrLf
    SQL = SQL & "     , SEX                     AS SEX      " & vbCrLf
    SQL = SQL & "     , GWA                     AS DEPT     " & vbCrLf
    SQL = SQL & "     , COUNT(HANGMOG_CODE)     AS CNT      " & vbCrLf
    'SQL = SQL & "     , INTERFACE_YN                        "
    'SQL = SQL & "     , JANGBI_OUT_CODE                     "
    'SQL = SQL & "     , JANGBI_CODE                         "
    'SQL = SQL & "     , CONFIRM_YN                          "
    'SQL = SQL & "     , CPL_RESULT                          "
    SQL = SQL & "  FROM VW_CPL_INTERFACE_GUMSA_LOAD         " & vbCrLf
    SQL = SQL & " WHERE PART_JUBSU_DATE BETWEEN  '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    SQL = SQL & "   AND NVL(CONFIRM_YN, 'N') = 'N'          " & vbCrLf
    SQL = SQL & "   AND HANGMOG_CODE IN (" & gAllTestCd & ")" & vbCrLf
    'SQL = SQL & "   AND JANGBI_CODE = '" & gHOSP.MACHCD & "'"
    SQL = SQL & " GROUP BY PART_JUBSU_DATE, SPECIMEN_SER, BUNHO, FKCPL0201, SPECIMEN_CODE, SUNAME, AGE, SEX, GWA   " & vbCrLf
    SQL = SQL & " ORDER BY PART_JUBSU_DATE, SPECIMEN_SER, BUNHO, FKCPL0201, SUNAME                  " & vbCrLf
    
    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("SPCCD")), intRow, colSPECIMEN
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("DEPT")) & "", intRow, colDEPT
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                    SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                    frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
'            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
'            frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
'            frmWorkList.chkAll.Value = "0"
        End If
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_SWMC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_VHS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""



'SELECT DISTINCT  A.SPC_NO ,
'A.PT_NO, C.PT_NAME, C.SEX || '/' || FC_SUP_AGE(C.BIRTH_DATE) SEX_AGE, B.DEPT_CD || '/' || B.WARD_NO DEPT_WARD, A.ACPNO_1, A.EXAM_CD
'fROM SLAXWORKT A, SLA1COLMT B, ARRPATBAMV C
'Where A.pt_no = C.pt_no
'AND A.SPC_NO = B.SPC_NO
'and A.ORD_DATE between '2020-06-15' and '2020-06-16'
'and A.EXAM_CD IN ('L5300I','L5300J','L9084','L9085','L1202','L1203')
'AND A.RPT_YN <> 'Y'
'AND B.EXAM_PROSS_STS IN ('C', 'D', 'M')
'ORDER BY A.EXAM_CD



'SQL = "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN';"
'SQL = "ALTER SESSION SET NLS_DATE_FORMAT = 'DD-MON-RR';"
'SQL = "ALTER SESSION SET NLS_DATE_LANGUAGE = 'AMERICAN';"

'    SQL = "ALTER SESSION SET NLS_LANGUAGE = 'KOREAN';"
    SQL = "ALTER SESSION SET NLS_DATE_FORMAT = 'RR/MM/DD'"
'    SQL = "ALTER SESSION SET NLS_DATE_LANGUAGE = 'KOREAN';"

    If Not DBExec(AdoCn, SQL) Then
'        Exit Sub
    End If


    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       1                       AS SORT " & vbCrLf
    '2020-07-30 ���� : ó�����ڰ� �ƴ� �������� �� �����;� ��
    'SQL = SQL & "     , A.ORD_DATE              AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , A.UPD_DAY              AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , A.PT_NO                 AS PID  " & vbCrLf
    SQL = SQL & "     , A.SPC_NO                AS BARCODE  " & vbCrLf
    SQL = SQL & "     , A.ACPNO_1               AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , C.PT_NAME               AS PNAME    " & vbCrLf
    SQL = SQL & "     , FC_SUP_AGE(C.BIRTH_DATE) AS AGE      " & vbCrLf
    SQL = SQL & "     , C.SEX                   AS SEX      " & vbCrLf
    SQL = SQL & "     , B.DEPT_CD               AS DEPT     " & vbCrLf
'    SQL = SQL & "     , COUNT(EXAM_CD)          AS CNT      " & vbCrLf
    SQL = SQL & "  FROM SLAXWORKT A, SLA1COLMT B, ARRPATBAMV C" & vbCrLf
    SQL = SQL & " Where A.pt_no = C.pt_no" & vbCrLf
    SQL = SQL & "   AND A.SPC_NO = B.SPC_NO" & vbCrLf
    
    SQL = SQL & "   and A.ORD_DATE BETWEEN  '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
'    SQL = SQL & "   and TO_CHAR(A.ORD_DATE,'YYYY-MM-DD') BETWEEN  '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    
    
    'SQL = SQL & "   and A.EXAM_CD IN (" & gAllTestCd & ")" & vbCrLf
    '2021-05-13 L5300T �߰�
    '2021-08-04 L5300U1 �߰�
    SQL = SQL & "   and A.EXAM_CD IN ('L5300T','L5300I','L5300J','L5300M','L9084','L9085','L5300U1')" & vbCrLf
    SQL = SQL & "   AND A.RPT_YN <> 'Y'" & vbCrLf
    SQL = SQL & "   AND B.EXAM_PROSS_STS IN ('C', 'D', 'M')" & vbCrLf
    SQL = SQL & " Union All " & vbCrLf
    SQL = SQL & "SELECT DISTINCT " & vbCrLf
    SQL = SQL & "       2                       AS SORT " & vbCrLf
    '2020-07-30 ���� : ó�����ڰ� �ƴ� �������� �� �����;� ��
    'SQL = SQL & "     , A.ORD_DATE              AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , A.UPD_DAY              AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , A.PT_NO                 AS PID  " & vbCrLf
    SQL = SQL & "     , A.SPC_NO                AS BARCODE  " & vbCrLf
    SQL = SQL & "     , A.ACPNO_1               AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , C.PT_NAME               AS PNAME    " & vbCrLf
    SQL = SQL & "     , FC_SUP_AGE(C.BIRTH_DATE) AS AGE      " & vbCrLf
    SQL = SQL & "     , C.SEX                   AS SEX      " & vbCrLf
    SQL = SQL & "     , B.DEPT_CD               AS DEPT     " & vbCrLf
'    SQL = SQL & "     , COUNT(EXAM_CD)          AS CNT      " & vbCrLf
    SQL = SQL & "  FROM SLAXWORKT A, SLA1COLMT B, ARRPATBAMV C" & vbCrLf
    SQL = SQL & " Where A.pt_no = C.pt_no" & vbCrLf
    SQL = SQL & "   AND A.SPC_NO = B.SPC_NO" & vbCrLf
    '2020-07-30 ���� : ó�����ڰ� �ƴ� �������� �� �����;� ��
    
    SQL = SQL & "   and A.ORD_DATE BETWEEN  '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    
    
    
    'SQL = SQL & "   and A.EXAM_CD IN (" & gAllTestCd & ")" & vbCrLf
    '2021-06-15 �߰� ,'L1204' �߰�
    SQL = SQL & "   and A.EXAM_CD IN ('L1202','L1203','L1204')" & vbCrLf
    SQL = SQL & "   AND A.RPT_YN <> 'Y'"
    SQL = SQL & "   AND B.EXAM_PROSS_STS IN ('C', 'D', 'M')"

'    SQL = SQL & " GROUP BY PART_JUBSU_DATE, SPECIMEN_SER, BUNHO, FKCPL0201, SPECIMEN_CODE, SUNAME, AGE, SEX, GWA   " & vbCrLf
    SQL = SQL & " ORDER BY SORT, HOSPDATE, CHARTNO"
    
    
    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("DEPT")) & "", intRow, colDEPT
                    
                    'SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                    SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                    frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
'            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
'            frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
'            frmWorkList.chkAll.Value = "0"
        End If
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_VHS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub


Public Sub GetWorkList_EHWA(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim j           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    Dim sParam      As String
    Dim sRcvData    As String
    Dim varRcvData  As Variant
    Dim varTstCode  As Variant
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""


    Dim strRegDate      As String
'    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
'    Dim sParam      As String
'    Dim sRcvData    As String
'    Dim varRcvData  As Variant
'    Dim varTstCode  As Variant
'    Dim i           As Integer
'    Dim J           As Integer
    
    Dim sRes        As String
    
'On Error GoTo DBErr
    
    
    intTestCnt = 0
    gPatOrdCd = ""
    ReDim Preserve gPatTest(0)
    
'    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
'    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
'    strPatID = Trim(GetText(SPD, asRow, colPID))
'    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
'    If strBarcode = "" Then
'        Exit Function
'    End If
        
        
    Screen.MousePointer = 11
  
    sRes = Online_XML(gXml_ORDER_SELECT, strBarcode) ' "PKG_MSE_LM_INTERFACE.PC_MSE_ORDER_SELECT"
  
    If sRes <> "" Then
        For i = 0 To giIndex
            With SPD
                .ReDraw = False
                
                'ȯ�� ����/����
                With mPatient
                    .SEX = gPatInfo_Select.SEX_TP_CD
                    .AGE = gPatInfo_Select.PT_BRDY_DT
                End With

'                SetText SPD, "1", asRow, colCHECKBOX
'                SetText SPD, gPatInfo_Select.ACPT_DTM, asRow, colHOSPDATE
'                SetText SPD, gPatInfo_Select.PT_HME_DEPT_CD, asRow, colINOUT
'                SetText SPD, strBarcode, asRow, colBARCODE
'                SetText SPD, gPatInfo_Select.PT_NO, asRow, colPID
'                SetText SPD, gPatInfo_Select.PT_NM, asRow, colPNAME
'                SetText SPD, gPatInfo_Select.SEX_TP_CD, asRow, colPSEX
'                SetText SPD, gPatInfo_Select.PT_BRDY_DT, asRow, colPAGE
                
                '��������
'                SetText SPD, CStr(intTestCnt), asRow, colOCNT

                '���������� ����
                With mOrder
                    .BarNo = strBarcode
                    .PID = gPatInfo_Select.PT_NO
                    .PNAME = gPatInfo_Select.PT_NM
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With

                '-- ȭ�鿡 ǥ��
                'If Trim(varRcvData(10) & "") <> "" Then
'                    For intCol = colSTATE + 1 To .MaxCols
'                        If gExam_Select(i).TST_CD = gArrEQP(intCol - colSTATE, 2) Then
'                            .Row = asRow
'                            .Col = intCol
'                            .BackColor = vbYellow
'                            Call SetText(SPD, "��", asRow, intCol)
'                            Exit For
'                        End If
'                    Next
                'End If
                
            End With
            DoEvents
            
        Next
    Else
    
    End If
    
    RS.Close
    
    Screen.MousePointer = 0
    


DBErr:
    intTestCnt = 0
    Screen.MousePointer = 0
    
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_NU" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
    
    
'''    sParam = ""
'''    sParam = sParam & "submit_id=TRLII00101&"                                   'submit ID
'''    sParam = sParam & "business_id=li&"                                         'business_id
'''    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '����ڵ�
'''
'''    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'''
'''    Call SetSQLData("��ũ��ȸ", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'''
'''    If InStr(1, sRcvData, "<?xml version") > 0 Then
'''        varRcvData = Split(sRcvData, "CDATA[")
'''    End If
'''
'''    If UBound(varRcvData) >= 0 Then
'''        For i = 1 To UBound(varRcvData)
'''            varRcvData(i) = Mid(varRcvData(i), 1, InStr(varRcvData(i), "]") - 1)
'''        Next
'''
'''        SPD.MaxRows = 0
'''
'''        For i = 1 To UBound(varRcvData) Step 14
'''            With SPD
'''                .ReDraw = False
'''                For J = 1 To SPD.DataRowCnt
'''                    strHospDate = GetText(SPD, J, colHOSPDATE)
'''                    strBarcode = GetText(SPD, J, colBARCODE)
'''                    If Format(Mid(varRcvData(i), 1, 8), "####-##-##") = strHospDate And varRcvData(i + 2) & "" = strBarcode Then
'''                        blnSame = True
'''                    End If
'''                Next
'''
'''                If blnSame = False Then
'''                    .MaxRows = .MaxRows + 1
'''                    intRow = .MaxRows
'''
'''                    SetText SPD, "1", intRow, colCHECKBOX
'''                    SetText SPD, Format(Mid(varRcvData(i), 1, 8), "####-##-##"), intRow, colHOSPDATE
'''                    SetText SPD, varRcvData(i + 1) & "", intRow, colINOUT
'''                    SetText SPD, varRcvData(i + 2) & "", intRow, colBARCODE
'''                    SetText SPD, varRcvData(i + 3) & "", intRow, colPID
'''                    SetText SPD, varRcvData(i + 4) & "", intRow, colPNAME
'''                    SetText SPD, mGetP(varRcvData(i + 5) & "", 1, "/"), intRow, colPSEX
'''                    SetText SPD, mGetP(varRcvData(i + 5) & "", 2, "/"), intRow, colPAGE
'''
'''                    strTestCds = varRcvData(i + 9) & ""
'''                    strTestCds = Replace(strTestCds, "��", "")
'''
'''                    If InStr(varRcvData(i + 10) & "", "LIM305") > 0 Then
'''                        .SetText 14, intRow, "Inhalant"
'''                    ElseIf InStr(varRcvData(i + 10) & "", "LIM306") > 0 Then
'''                        .SetText 14, intRow, "Food"
'''                    End If
'''
''''                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
'''
''''                    If gWORKPOS = "P" Then
'''                        'SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
'''                        'frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
''''                    Else
''''                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
''''                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
''''                    End If
'''                End If
'''
'''            End With
'''
'''            blnSame = False
'''
'''            DoEvents
'''
'''            RS.MoveNext
'''        Next
'''
'''        If gWORKPOS = "P" Then
''''            frmWorkList.chkAll.Value = "1"
'''        End If
'''    Else
'''        If gWORKPOS = "P" Then
''''            frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
''''            frmWorkList.chkAll.Value = "0"
'''        End If
'''    End If
'''
'''    RS.Close
'''
'''    SPD.RowHeight(-1) = 15
'''    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub


'Public Sub GetWorkList_NU(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
'    Dim RS          As ADODB.Recordset
'    Dim blnSame     As Boolean
'
'    Dim i           As Integer
'    Dim j           As Integer
'    Dim iCnt        As Integer
'    Dim intRow      As Integer
'    Dim strHospDate As String
'    Dim strBarcode  As String
'    Dim strTestCds  As String
'    Dim sParam      As String
'    Dim sRcvData    As String
'    Dim varRcvData  As Variant
'    Dim varTstCode  As Variant
'
'
''''    Dim sSch1, sSch2 As String
''''    Dim sParam As String
''''    Dim sRcvData, sData As String
''''    Dim varRcvData As Variant
''''    Dim varTstCode As Variant
''''    Dim i As Integer
''''    Dim strTstCD As String
''''    Dim strItems As String
''''    Dim intRow As Integer
''''    Dim strTestCds As String
''''
''''On Error GoTo ErrorTrap
'
''''    sSch1 = Format(dtpSDate.Value, "yyyymmdd")
''''    sSch2 = Format(dtpEDate.Value, "yyyymmdd")
''''
''''    ClearSpread vasList
''''    vasList.MaxRows = 0
''''
''''    'strTestCds = "LIM305��LIM306��"
''''    'strTestCds = "LIM305"
''''
''''
''''    If optState(0).Value = True Then
''''        'sParam = "submit_id=TRLII00119&"                                           'submit ID
''''        sParam = "submit_id=TRLII00101&"                                            'submit ID
''''        sParam = sParam & "business_id=lis&"                                        'business_id
''''        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '�����ID|����ڵ�
''''        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                            '����ڵ�
''''        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                            '����ڵ�
''''        sParam = sParam & "startdd=" & sSch1 & "&"                                  '�����۾�����
''''        sParam = sParam & "enddd=" & sSch2 & "&"                                    '�����۾�����
''''    Else
''''        sParam = "submit_id=TRLQI00101&"                                            'submit ID
''''        sParam = sParam & "business_id=lis&"                                        'business_id
''''        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '�����ID|����ڵ�
''''        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                            '����ڵ�
''''        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                            '����ڵ�
''''        sParam = sParam & "startdd=" & sSch1 & "&"                                  '�����۾�����
''''        sParam = sParam & "enddd=" & sSch2 & "&"                                    '�����۾�����
''''    End If
''''
''''    '==> ������ ������ȸ
''''    'SetRawData "[WL_IN]" & sParam
''''        'spcacptdt ��������
''''        'acptflag �Կ��ܷ�����
''''        'bcno ��ü��ȣ
''''        'PID ��Ϲ�ȣ
''''        'patnm ȯ�ڸ�
''''        'sexage ���̼���
''''        'erprcpflag ���ޱ���
''''        'workno �۾���ȣ
''''        'tsectnm �˻���
''''        'ifreqcdlist ����û�ڵ�
''''        'tclscdlist �˻縮��Ʈ
''''        'urinextrvol ������
''''        'retestyn ��˿���
''''        'rsltstat �������
''''    sRcvData = OpenURLWithIE2(NUAPI.APIURL & sParam, Inet1)
''''
''''    Call SetSQLData("��ũ��ȸ", NUAPI.APIURL & sParam & vbNewLine & sRcvData)
''''
''''    If InStr(1, sRcvData, "<?xml version") > 0 Then
''''        varRcvData = Split(sRcvData, "CDATA[")
''''    End If
''''
''''    If UBound(varRcvData) >= 0 Then
''''        For i = 1 To UBound(varRcvData)
''''            varRcvData(i) = Mid(varRcvData(i), 1, InStr(varRcvData(i), "]") - 1)
''''        Next
''''
''''        For i = 1 To UBound(varRcvData) Step 14
''''            With vasList
''''                .MaxRows = .MaxRows + 1
''''
''''
''''                intRow = .MaxRows
''''                .Row = intRow
''''                '.Col = 7
''''                '.BackColor = vbGreen '&HC6FEFF '&H80C0FF
''''
''''                .SetText 1, intRow, "1"
''''                .SetText 2, intRow, Format(Mid(varRcvData(i), 1, 8), "####-##-##")
''''                .SetText 3, intRow, varRcvData(i + 1) & ""
''''                .SetText 4, intRow, varRcvData(i + 2) & ""
''''                .SetText 5, intRow, varRcvData(i + 3) & ""
''''                .SetText 6, intRow, varRcvData(i + 4) & ""
''''                .SetText 7, intRow, mGetP(varRcvData(i + 5) & "", 1, "/")
''''                .SetText 8, intRow, mGetP(varRcvData(i + 5) & "", 2, "/")
''''                .SetText 9, intRow, varRcvData(i + 6) & ""
''''                .SetText 10, intRow, varRcvData(i + 7) & ""
''''                .SetText 11, intRow, varRcvData(i + 8) & ""
''''
''''                strTestCds = varRcvData(i + 9) & ""
''''                strTestCds = Replace(strTestCds, "��", "")
''''
''''                If InStr(varRcvData(i + 10) & "", "LIM305") > 0 Then
''''                    .SetText 14, intRow, "Inhalant"
''''                ElseIf InStr(varRcvData(i + 10) & "", "LIM306") > 0 Then
''''                    .SetText 14, intRow, "Food"
''''                End If
''''                .RowHeight(-1) = 12
''''            End With
''''        Next
''''    End If
''''
''''    chkAll.Value = "1"
''''
''''    'vasList.MaxRows = vasList.DataRowCnt
''''    vasList.RowHeight(-1) = 13.3
''''
''''    Exit Sub
''''
''''ErrorTrap:
''''
''''    MsgBox "��ȸ ����", vbOKOnly + vbCritical, Me.Caption
'On Error GoTo ErrHandle
'
'    Screen.MousePointer = 11
'    blnSame = False
'    strTestCds = ""
'
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00101&"                                   'submit ID
'    sParam = sParam & "business_id=li&"                                         'business_id
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '����ڵ�
'
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("��ũ��ȸ", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'    If InStr(1, sRcvData, "<?xml version") > 0 Then
'        varRcvData = Split(sRcvData, "CDATA[")
'    End If
'
'    If UBound(varRcvData) >= 0 Then
'        For i = 1 To UBound(varRcvData)
'            varRcvData(i) = Mid(varRcvData(i), 1, InStr(varRcvData(i), "]") - 1)
'        Next
'
'        SPD.MaxRows = 0
'
'        For i = 1 To UBound(varRcvData) Step 14
'            With SPD
'                .ReDraw = False
'                For j = 1 To SPD.DataRowCnt
'                    strHospDate = GetText(SPD, j, colHOSPDATE)
'                    strBarcode = GetText(SPD, j, colBARCODE)
'                    If Format(Mid(varRcvData(i), 1, 8), "####-##-##") = strHospDate And varRcvData(i + 2) & "" = strBarcode Then
'                        blnSame = True
'                    End If
'                Next
'
'                If blnSame = False Then
'                    .MaxRows = .MaxRows + 1
'                    intRow = .MaxRows
'
'                    SetText SPD, "1", intRow, colCHECKBOX
'                    SetText SPD, Format(Mid(varRcvData(i), 1, 8), "####-##-##"), intRow, colHOSPDATE
'                    SetText SPD, varRcvData(i + 1) & "", intRow, colINOUT
'                    SetText SPD, varRcvData(i + 2) & "", intRow, colBARCODE
'                    SetText SPD, varRcvData(i + 3) & "", intRow, colPID
'                    SetText SPD, varRcvData(i + 4) & "", intRow, colPNAME
'                    SetText SPD, mGetP(varRcvData(i + 5) & "", 1, "/"), intRow, colPSEX
'                    SetText SPD, mGetP(varRcvData(i + 5) & "", 2, "/"), intRow, colPAGE
'
'                    strTestCds = varRcvData(i + 9) & ""
'                    strTestCds = Replace(strTestCds, "��", "")
'
'                    If InStr(varRcvData(i + 10) & "", "LIM305") > 0 Then
'                        .SetText 14, intRow, "Inhalant"
'                    ElseIf InStr(varRcvData(i + 10) & "", "LIM306") > 0 Then
'                        .SetText 14, intRow, "Food"
'                    End If
'
''                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
'
''                    If gWORKPOS = "P" Then
'                        'SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
'                        'frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
''                    Else
''                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
''                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
''                    End If
'                End If
'
'            End With
'
'            blnSame = False
'
'            DoEvents
'
'            RS.MoveNext
'        Next
'
'        If gWORKPOS = "P" Then
''            frmWorkList.chkAll.Value = "1"
'        End If
'    Else
'        If gWORKPOS = "P" Then
''            frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
''            frmWorkList.chkAll.Value = "0"
'        End If
'    End If
'
'    RS.Close
'
'    SPD.RowHeight(-1) = 15
'    SPD.ReDraw = True
'
'    Screen.MousePointer = 0
'
'Exit Sub
'
'ErrHandle:
'    Screen.MousePointer = 1
'
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show vbModal
'
'End Sub


'Public Sub GetWorkList_HDINFO(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
'    Dim RS          As ADODB.Recordset
'    Dim blnSame     As Boolean
'
'    Dim i           As Integer
'    Dim j           As Integer
'    Dim k           As Integer
'    Dim iCnt        As Integer
'    Dim intRow      As Integer
'    Dim strHospDate As String
'    Dim strBarcode  As String
'    Dim sParam      As String
'    Dim strTestCds  As String
'    Dim sRcvData    As String
'    Dim varRcvData  As Variant
'    Dim varTstCode  As Variant
'    Dim strNames    As String
'    Dim strXmlName  As String
'    Dim strWorkNo   As String
'
'    Dim l As Integer
'
'On Error GoTo RST
'
'    Screen.MousePointer = 11
'    SPD.MaxRows = 0
'
'    blnSame = False
'    strNames = ""
'    l = 0
'
'ReSearch:
'
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00123&"                               'submit ID
'    sParam = sParam & "business_id=lis&"                                    'business_id
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                        '����ڵ�
'    sParam = sParam & "startdd=" & pFrom & "&"                              '�����۾�����
'    sParam = sParam & "enddd=" & pTo & "&"                                  '�����۾�����
'
'    If gTest = "MTB" Then
'        'sParam = sParam & "testcd=C6021C&"                                  '�˻��ڵ�
'        sParam = sParam & "testcd=" & gComm.MTBORD & "&"                                  '�˻��ڵ�
'
'    ElseIf gTest = "RP" Then
'        If l = 0 Then
''            sParam = sParam & "testcd=VB8506B&"                             '�˻��ڵ�
'            sParam = sParam & "testcd=" & gComm.RP19ORD_1 & "&"                                   '�˻��ڵ�
'        Else
''            sParam = sParam & "testcd=D6802060&"                            '�˻��ڵ�
'            sParam = sParam & "testcd=" & gComm.RP19ORD_2 & "&"                                   '�˻��ڵ�
'        End If
'
'    ElseIf gTest = "PB" Then
''        sParam = sParam & "testcd=D6801&"                                   '�˻��ڵ�
'        sParam = sParam & "testcd=" & gComm.PB6ORD & "&"                                    '�˻��ڵ�
'
'    ElseIf gTest = "RPPB" Then
'        If l = 0 Then
''            sParam = sParam & "testcd=VB8506B&"                             '�˻��ڵ�    RP19��
'            sParam = sParam & "testcd=" & gComm.RP19ORD_1 & "&"                                   '�˻��ڵ�
'        ElseIf l = 1 Then
''            sParam = sParam & "testcd=D6802060&"                            '�˻��ڵ�    RP19��
'            sParam = sParam & "testcd=" & gComm.RP19ORD_2 & "&"                                   '�˻��ڵ�
'        Else
''            sParam = sParam & "testcd=D6801&"                               '�˻��ڵ�    PB 6��
'            sParam = sParam & "testcd=" & gComm.PB6ORD & "&"                                    '�˻��ڵ�
'        End If
'    End If
'
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("��ũ��ȸ", "Param:" & gHOSP.APIURL & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'    '<worklist>
'        '<bcno><![CDATA[3041900020]]></bcno>
'        '<patnm><![CDATA[�̸��]]></patnm>
'        '<prgstno><![CDATA[600603-2******]]></prgstno>
'        '<pid><![CDATA[000137388]]></pid>
'        '<sex><![CDATA[F]]></sex>
'        '<age><![CDATA[59]]></age>
'        '<spcnm><![CDATA[Throat swab]]></spcnm>
'        '<spccd><![CDATA[023]]></spccd>
'        '<tclscd><![CDATA[VB6012A]]></tclscd>
'        '<spcstat><![CDATA[4]]></spcstat>
'        '<rsltstat><![CDATA[4]]></rsltstat>                         '�������
'        '<workno><![CDATA[20191025I20001]]></workno>
'        '<testcd><![CDATA[VB6012A]]></testcd>
'        '<execprcpuniqno><![CDATA[2009768025]]></execprcpuniqno>
'        '<spcacptdt><![CDATA[20191025092308]]></spcacptdt>
'        '<prcpdd><![CDATA[20191025]]></prcpdd>
'        '<retestyn><![CDATA[N]]></retestyn>
'        '<testlrgcd><![CDATA[I]]></testlrgcd>
'        '<orddeptcd><![CDATA[RM]]></orddeptcd>
'    '</worklist>
'
'    If InStr(1, sRcvData, "<?xml version") > 0 Then
'        varRcvData = Split(sRcvData, "<worklist>")
'    End If
'
'    strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & ".xml"
'
'    Call SetXMLData(strXmlName, sRcvData)
'
'    Call DisplayNode_InfoS(App.PATH & "\Xml\" & strXmlName, UBound(varRcvData))
'
'    Kill App.PATH & "\Xml\" & strXmlName
'
'    If UBound(varRcvData) >= 1 Then
'        For i = 0 To UBound(varRcvData) - 1 'Step 19
'            With SPD
'                .ReDraw = False
'                blnSame = False
'
'                '2019-12-11 �޸�
'                '   rsltstat �� ó�� ���� �� ��...
'                '   <rsltstat><![CDATA[-]]></rsltstat>
'                '   <rsltstat><![CDATA[4]]></rsltstat>
'
'                'If GetSampleTest_HDINFO(XmlSelectS.BCNO(i)) > 0 Then
'
'                    For j = 1 To SPD.DataRowCnt
'                        strHospDate = GetText(SPD, j, colHOSPDATE)
'                        strBarcode = GetText(SPD, j, colBARCODE)
'                        If XmlSelectS.PRCPDD(i) & "" = strHospDate And XmlSelectS.BCNO(i) = strBarcode Then
'                            blnSame = True
'                            strNames = GetText(SPD, intRow, colITEMS)
'                            strNames = strNames & "|" & GetTestNm(XmlSelectS.TESTCD(i))
'
'                            SetText SPD, strNames, intRow, colITEMS
'                            strNames = ""
'                        End If
'                    Next
'
'                    If blnSame = False Then
'                        .MaxRows = .MaxRows + 1
'                        intRow = .MaxRows
'
'                        SetText SPD, "1", intRow, colCHECKBOX
'                        SetText SPD, XmlSelectS.PRCPDD(i), intRow, colHOSPDATE
'                        SetText SPD, XmlSelectS.BCNO(i), intRow, colBARCODE
'                        SetText SPD, XmlSelectS.PID(i), intRow, colPID
'                        SetText SPD, XmlSelectS.PATNM(i), intRow, colPNAME
'                        SetText SPD, XmlSelectS.SEX(i), intRow, colPSEX
'                        SetText SPD, XmlSelectS.AGE(i), intRow, colPAGE
'                        SetText SPD, XmlSelectS.SPCNM(i), intRow, colSPECIMEN
'
'                        strWorkNo = XmlSelectS.WORKNO(i)
'                        strWorkNo = Mid(strWorkNo, 1, 8) & "-" & Mid(strWorkNo, 9, 2) & "-" & Mid(strWorkNo, 11)
'                        SetText SPD, strWorkNo, intRow, colCHARTNO
'
'                        strNames = GetText(SPD, intRow, colITEMS)
'                        strNames = GetTestNm(XmlSelectS.TESTCD(i))
'                        SetText SPD, strNames, intRow, colITEMS
'
'                        If gTest = "MTB" Then
'                            SetText SPD, "MTB", intRow, colPOSNO
'
'                        ElseIf gTest = "RP" Then
'                            SetText SPD, "RP", intRow, colPOSNO
'
'                        ElseIf gTest = "PB" Then
'                            SetText SPD, "PB", intRow, colPOSNO
'
'                        ElseIf gTest = "RPPB" Then
'                            If l = 0 Then
'                                SetText SPD, "RP", intRow, colPOSNO
'                            ElseIf l = 1 Then
'                                SetText SPD, "RP", intRow, colPOSNO
'                            Else
'                                SetText SPD, "PB", intRow, colPOSNO
'                            End If
'                        End If
'
'                    End If
'                'End If
'            End With
'        Next
'    Else
'        MsgBox "��ȸ ����ڰ� �����ϴ�.", vbOKOnly + vbCritical, "��ũ����Ʈ ��ȸ"
'    End If
'
'    If gTest = "RP" And l = 0 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    If gTest = "RPPB" And l = 0 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    If gTest = "RPPB" And l = 1 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    SPD.RowHeight(-1) = 12
'    SPD.ReDraw = True
'
'    Screen.MousePointer = 0
'
'Exit Sub
'
'RST:
'
'                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_HDINFO" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show 'vbModal
'
'    Screen.MousePointer = 0
'
'End Sub

'-- �˻��� ITEM ��������
Function GetSampleITEM(ByVal asRow As Long, ByVal SPD As vaSpread) As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strRegDate      As String
    Dim strChartNo      As String
    Dim strInOut        As String
    
    Dim lngExamNo       As Long
    Dim strItems        As String
    Dim strSpcYY        As String
    Dim strSpcNo        As String
    
    GetSampleITEM = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    strInOut = Trim(GetText(SPD, asRow, colINOUT))
    
    If strBarcode = "" Then
        Exit Function
    End If
        
    Select Case gEMR
        Case "SWMC"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT HANGMOG_CODE   AS ITEM         " & vbCrLf
            SQL = SQL & "  FROM VW_CPL_INTERFACE_GUMSA_LOAD             " & vbCrLf
            SQL = SQL & " WHERE SPECIMEN_SER = '" & strBarcode & "'     " & vbCrLf
            SQL = SQL & "   AND NVL(CONFIRM_YN, 'N') = 'N'              " & vbCrLf
            SQL = SQL & "   AND HANGMOG_CODE IN (" & gAllTestCd & ")    " & vbCrLf
            'SQL = SQL & "   AND JANGBI_CODE = '" & gHOSP.MACHCD & "'   " & vbCrLf
            SQL = SQL & " ORDER BY HANGMOG_CODE                         " & vbCrLf
        
        Case "AMIS"
            SQL = ""
            SQL = SQL & "SELECT R.RESULTITEMCODE as ITEM                    " & vbCr
            SQL = SQL & "  FROM registinfos O, resultofnum R                " & vbCr
            SQL = SQL & " WHERE O.acptdate = R.acptdate                     " & vbCr
            SQL = SQL & "   AND R.SPCMNO = '" & strBarcode & "'             " & vbCr
            SQL = SQL & "   AND O.patid = R.patid                           " & vbCr
            SQL = SQL & "   AND O.acptseq = R.acptseq                       " & vbCr
            SQL = SQL & "   AND O.CLAS = 4                                  " & vbCr '�ӻ󺴸�
            SQL = SQL & "   AND R.RESULTFLAG = 0                            " & vbCr
            SQL = SQL & "   AND R.ORDERCODE IN (" & gAllOrdCd & ")          " & vbCr
            SQL = SQL & "   AND R.RESULTITEMCODE in (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & "  ORDER BY R.RESULTITEMCODE                        " & vbCr
        
        Case "BIGUBCARE"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT i.IntLabCod + cast(IntLabseq as varchar(3)) AS ITEM "
            SQL = SQL & "  from interfacedb..IntRst i, aphdb..rstinf r " & vbCr
            SQL = SQL & " WHERE r.RstOdrStt not in ('OC') " & vbCr
            SQL = SQL & "   AND (r.rstrstval = '' or rstrstval is null)" & vbCr
            'If gHOSP.MACHNM <> "HITACHI7080" Then
                SQL = SQL & "   AND i.intodrtyp = '" & gHOSP.PARTCD & "'" & vbCr  ''HEMO'
            'End If
            SQL = SQL & "   AND i.IntOdrDte = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   AND i.IntLabNum = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND i.IntChtNum = '" & strChartNo & "'" & vbCr
'            SQL = SQL & "   AND i.IntLabCod IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND i.IntLabCod + cast(IntLabseq as varchar(3)) IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND i.intlabnum = r.rstlabnum" & vbCr
            SQL = SQL & "   AND i.intodrdte = r.rstodrdte" & vbCr
            SQL = SQL & "   AND i.intlabseq = r.rstlabseq" & vbCr
            SQL = SQL & "   AND i.intlabcod = r.rstodrcod" & vbCr
        
        Case "BIT"
            SQL = ""
            SQL = SQL & " SELECT DISTINCT R.ResLabCod AS ITEM                   " & vbCr
            SQL = SQL & "   FROM RESINF AS R                                    " & vbCr
            SQL = SQL & " WHERE LTRIM(RTRIM(R.RESOCMNUM)) = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")             " & vbCr
            SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F')     " & vbCr         '--  'I':�߰� 'F' �Ϸ�"
            SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)      " & vbCr
            SQL = SQL & " Order By R.ResLabCod                                  " & vbCr
        
        Case "BIT70"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT L.LABODRCOD as ITEM                " & vbCr
            'SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D, ME_MAN M" & vbCr
            SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D                       " & vbCr
            SQL = SQL & " WHERE L.LABCHTNUM  = '" & strChartNo & "'         " & vbCr
            SQL = SQL & "   AND L.LABODRDTE  = '" & strRegDate & "'         " & vbCr
            SQL = SQL & "   AND L.LABKEYNUM  = D.DATKEYNUM                  " & vbCr                    '-- ���̺���Ű��
            SQL = SQL & "   AND L.LABATTEND  = D.DATATTEND                  " & vbCr                    '-- ������ȣ
            'SQL = SQL & "   AND L.LABATTEND = M.MANATTEND                  " & vbCr                    '-- ������ȣ
            SQL = SQL & "   AND L.LABCHTNUM  = D.DATCHTNUM                  " & vbCr                    '-- íƮ��ȣ
            SQL = SQL & "   AND L.LABCHTNUM  = M.MANCHTNUM                  " & vbCr                    '-- íƮ��ȣ
            SQL = SQL & "   AND L.LABODRDTE  = D.DATODRDTE                  " & vbCr                    '-- ó������
            SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCr
            SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCr    '-- ��ҿ���
            SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)  " & vbCr
            SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCr                            '-- ó������ (2:����, 3:����Է�)
            SQL = SQL & " Order By L.LABODRCOD                              " & vbCr
        
        Case "EONM"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT O.H141_SUGACD AS ITEM              " & vbCr
            SQL = SQL & "  FROM TB_H141_LISTAKEBODY O, TB_A110_PATINFO P    " & vbCr
            SQL = SQL & " Where P.A110_ChartNo = O.H141_CHARTNO             " & vbCr
            SQL = SQL & "   AND O.H141_TSAMPLENO  = '" & strBarcode & "'    " & vbCr
            'SQL = SQL & "   AND O.H141_NOTYYN = 'N'                         " & vbCr
            SQL = SQL & "   AND O.H141_NOTYYN       IN ('N','T')                 " & vbCr '������:T
            SQL = SQL & "   And O.H141_SUGACD in (" & gAllTestCd & ")       " & vbCr
            SQL = SQL & " ORDER BY O.H141_SUGACD                            " & vbCr
        
         Case "EASYS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM                     " & vbCr
            SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c    " & vbCr
            SQL = SQL & " WHERE a.ACC_YMD   = '" & strRegDate & "'          " & vbCr
            SQL = SQL & "   AND a.RECEPT_NO = '" & strBarcode & "'          " & vbCr
            SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ")            " & vbCr
            SQL = SQL & "   AND a.STS_CD    = 'A'                           " & vbCr 'A:����, R:�������
            SQL = SQL & "   AND a.SUTAK_CD  = ''                            " & vbCr
            SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO                   " & vbCr
            SQL = SQL & " ORDER BY ORD_CD                                   " & vbCr
        
        Case "GINUS"
            SQL = ""
            SQL = SQL & "SELECT /*+ INDEX(rslt scrrslth_ux1) INDEX (coif scccoifm_ix1) */" & vbCr
            SQL = SQL & "       rslt.cd as ITEM                                         " & vbCr
            SQL = SQL & "  FROM scrrslth rslt                                           " & vbCr
            SQL = SQL & "     , scccoifm coif                                           " & vbCr
            SQL = SQL & "     , scccodem codm                                           " & vbCr
            SQL = SQL & "     , scrprexh prex                                           " & vbCr
            SQL = SQL & "     , mosxpslh xpsl                                           " & vbCr
            SQL = SQL & "     , pmcptbsm ptbs                                           " & vbCr
            SQL = SQL & "WHERE rslt.hos_org_no   = '" & gHOSP.HOSPCD & "'               " & vbCr
            SQL = SQL & "  AND rslt.smp_no       = '" & strBarcode & "'                 " & vbCr
            SQL = SQL & "  AND rslt.exam_stus  IN ('0','1','2')                         " & vbCr
            SQL = SQL & "  AND coif.hos_org_no   = rslt.hos_org_no                      " & vbCr
            'SQL = SQL & "  AND coif.exam_cd      = rslt.cd                              " & vbCr
            SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt" & vbCr
            SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN codm.fr_dt AND codm.to_dt" & vbCr
            SQL = SQL & "  AND coif.exam_mach_cd = '" & gHOSP.MACHCD & "'               " & vbCr
            SQL = SQL & "  AND codm.hos_org_no   = coif.hos_org_no                      " & vbCr
            SQL = SQL & "  AND codm.typ_cd       = '02'                                 " & vbCr
            SQL = SQL & "  AND codm.cd           = coif.spc_cd                          " & vbCr
            SQL = SQL & "  AND prex.hos_org_no   = rslt.hos_org_no                      " & vbCr
            SQL = SQL & "  AND prex.smp_no       = rslt.smp_no                          " & vbCr
            SQL = SQL & "  AND prex.prcp_seq     = rslt.prcp_seq                        " & vbCr
            SQL = SQL & "  AND prex.exam_seq     = rslt.exam_seq                        " & vbCr
            SQL = SQL & "  AND xpsl.hos_org_no   = prex.hos_org_no                      " & vbCr
            SQL = SQL & "  AND xpsl.smp_no       = prex.smp_no                          " & vbCr
            SQL = SQL & "  AND xpsl.acp_no       = prex.prcp_seq                        " & vbCr
            SQL = SQL & "  AND xpsl.prcp_typ_cd IN ('O','C')                            " & vbCr
            SQL = SQL & "  AND ptbs.hos_org_no   = prex.hos_org_no                      " & vbCr
            SQL = SQL & "  AND ptbs.pt_no        = prex.pt_no                           " & vbCr
        
        Case "HWASAN"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT T.TESTCD as ITEM           " & vbCr
            SQL = SQL & "  FROM TC201 O, TC301 T                    " & vbCr
            SQL = SQL & " WHERE O.SPCNO = T.SPCNO                   " & vbCr
            SQL = SQL & "   AND O.SPCNO = '" & strBarcode & "'      " & vbCr
            SQL = SQL & "   And T.TESTCD in (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & " Order By T.TESTCD                         " & vbCr
        
        Case "JAINCOM"
            SQL = ""
            SQL = SQL & "SELECT DiSTINCT b.SCP42SUGACD as ITEM                  " & vbCr
            SQL = SQL & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b        " & vbCr
            SQL = SQL & " WHERE a.SCP41PCODE    = b.SCP42PCODE                  " & vbCr
            SQL = SQL & "   AND a.SCP41JDATE    = b.SCP42JDATE                  " & vbCr
            SQL = SQL & "   AND a.SCP41SID      = b.SCP42SID                    " & vbCr
            SQL = SQL & "   AND a.SCP41SPMNO2   = b.SCP42SPMNO2                 " & vbCr
            SQL = SQL & "   AND a.SCP41SPMNO2   = '" & strBarcode & "'          " & vbCr
            SQL = SQL & "   AND b.SCP42SUGACD  IN (" & gAllTestCd & ")          " & vbCr
            SQL = SQL & "   AND (b.SCP42RESULT IS NULL OR b.SCP42RESULT = '')   " & vbCr
            SQL = SQL & " ORDER BY b.SCP42SUGACD                                " & vbCr
        
        Case "JWINFO"
            'AND ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
            SQL = ""
            SQL = SQL & "SELECT DISTINCT LABCODE AS ITEM            " & vbCr
            SQL = SQL & "   FROM SLA_Labresult                      " & vbCr
            SQL = SQL & " WHERE LABCODE IN (" & gAllTestCd & ")     " & vbCr
            SQL = SQL & "   AND RECEIPTDATE = '" & strRegDate & "'  " & vbCr
            SQL = SQL & "   AND SPECIMENNUM = '" & strBarcode & "'  " & vbCr
            'SQL = SQL & "   AND JSTATUS < '3'                      " & vbCr
            SQL = SQL & " ORDER BY LABCODE                          " & vbCr
        
        Case "KOMAIN"
            SQL = ""
        
        Case "KCHART"
'            SQL = SQL & "    AND L.�˻����� = '" & gHOSP.LABCD & "'" & vbCr
            SQL = ""
            SQL = SQL & "SELECT DISTINCT (L.ó���ڵ� + L.�����ڵ�) AS ITEM                  " & vbCr
            SQL = SQL & "  FROM             TB_����˻� L                                   " & vbCr
            SQL = SQL & "       INNER JOIN  TB_�������� J ON  (L.��������ID = J.��������ID) " & vbCr
            SQL = SQL & "       INNER JOIN  TB_�����Ϲ� A ON  (J.��������   = A.��������    " & vbCr
            SQL = SQL & "                                AND   J.íƮ��ȣ   = A.íƮ��ȣ    " & vbCr
            SQL = SQL & "                                AND   J.�����ȣ   = A.�����ȣ)   " & vbCr
            SQL = SQL & " Where L.��ü��ȣ= '" & strBarcode & "'                            " & vbCr
            SQL = SQL & "   AND L.�˻���� < 5                                              " & vbCr
            SQL = SQL & "   AND L.ó���ڵ� + L.�����ڵ� IN (" & gAllTestCd & ")             " & vbCr
            SQL = SQL & " ORDER BY L.ó���ڵ�, L.�����ڵ�                                   " & vbCr
            
        Case "KYU"
            SQL = ""
            
        Case "MCC"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM             " & vbCr
            SQL = SQL & "  FROM LIS_INTERFACE1_V                    " & vbCr
            SQL = SQL & " WHERE READING_YMD = '" & strRegDate & "'  " & vbCr
            SQL = SQL & "   AND BCODE_NO    = '" & strBarcode & "'  " & vbCr
            SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ")      " & vbCr
            SQL = SQL & " ORDER BY ORD_CD                           " & vbCr
        
        Case "MEDICHART"
            SQL = ""
            SQL = SQL & "Select DISTINCT (a.ó���ڵ� + a.�����ڵ�)      AS ITEM     " & vbCr
            SQL = SQL & "  From TB_�˻��׸� a, TB_����⺻ c                        " & vbCr
            SQL = SQL & " Where a.íƮ��ȣ = '" & strChartNo & "'                   " & vbCr
            SQL = SQL & "   And a.ó���ȣ > 0                                      " & vbCr
            SQL = SQL & "   And c.������� IN ('1','5','6','7','8','9')             " & vbCr
            SQL = SQL & "   And (a.ó���ڵ� + a.�����ڵ�) IN (" & gAllTestCd & ")   " & vbCr
            SQL = SQL & "   And (a.�˻��� IS NULL OR a.�˻��� = '')             " & vbCr
            SQL = SQL & "   And a.�����    = c.�����                              " & vbCr
            SQL = SQL & "   And a.�����    = c.�����                              " & vbCr
            SQL = SQL & "   And a.������    = c.������                              " & vbCr
            SQL = SQL & "   And a.íƮ��ȣ  = c.íƮ��ȣ                            " & vbCr
            SQL = SQL & "   And (a.�˻��� IS NULL OR a.�˻��� = '')             " & vbCr
            SQL = SQL & " Order By ITEM                                             " & vbCr

'            SQL = ""
'            SQL = SQL & "Select DISTINCT (a.ó���ڵ� + a.�����ڵ�)      AS ITEM     " & vbCr
'            SQL = SQL & "  from tb_�˻��׸� " & vbCr
'            SQL = SQL & " Where íƮ��ȣ = '" & argPID & "'" & vbCr
'            SQL = SQL & "   And �����   = '" & strYear & "'" & vbCr
'            SQL = SQL & "   And �����   = '" & strMonth & "'" & vbCr
'            SQL = SQL & "   And ������   = '" & strDay & "'" & vbCr
'            SQL = SQL & "   And ó���ȣ > 0 " & vbCr
'            SQL = SQL & "   And (�˻��� is null or �˻��� = '') " & vbCr
'            SQL = SQL & "   And ó���ڵ�+�����ڵ� in (" & gAllExam & ")"
        
        Case "MEDITOLISS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT B.EXAM_CODE  AS ITEM                           " & vbCr
            SQL = SQL & "  FROM MEDITOLISS..TOTAL A, MEDITOLISS..TOTRES B               " & vbCr
            SQL = SQL + " WHERE A.EXAM_NO       = '" & strBarcode & "'                  " & vbCr
            SQL = SQL & "   And B.EXAM_CODE     IN (" & gAllTestCd & ")                 " & vbCr
            SQL = SQL & "   AND B.EXAM_PART     = 'C'                                   " & vbCr
            SQL = SQL & "   AND B.RESULT_VALUE  = ''                                    " & vbCr
            SQL = SQL & "   AND A.REQUEST_DATE  = B.REQUEST_DATE                        " & vbCr
            SQL = SQL & "   AND A.EXAM_NO       = B.EXAM_NO                             " & vbCr
                    
        Case "MOD"
            SQL = ""
            SQL = SQL & "Select Distinct c.EXAMCODE   AS ITEM           " & vbCr
            SQL = SQL & "  From EXAMREQ a, EXAMRES c                    " & vbCr
            SQL = SQL & " Where a.PID           = c.PID                 " & vbCr
            SQL = SQL & "   And a.SEQNO         = c.SEQNO               " & vbCr
            SQL = SQL & "   And a.RECENO        = c.RECENO              " & vbCr
            SQL = SQL & "   And c.SPECIMENID    = '" & strBarcode & "'  " & vbCr
            SQL = SQL & "   And c.EXAMCODE in (" & gAllTestCd & ")      " & vbCr
            SQL = SQL & "   And (c.EXAMEND = '' Or c.EXAMEND IS NULL)   " & vbCr
            SQL = SQL & " Order By c.EXAMCODE                           " & vbCr
    
        Case "MSINFOTEC"
            SQL = ""
            SQL = SQL & "Select DISTINCT ORCD AS ITEM       " & vbCrLf
            SQL = SQL & "  From emr.LRESULT                 " & vbCrLf
            SQL = SQL & " Where PAID = '" & strPatID & "'   " & vbCrLf
            SQL = SQL & "   And SPNO =  '" & strBarcode & "'" & vbCrLf
            SQL = SQL & "   And ORCD IN (" & gAllTestCd & ")" & vbCrLf
            SQL = SQL & "   And OKFL <> 'Y'                 " & vbCrLf   '-- ���Ȯ������
            SQL = SQL & " Order By ORCD                     " & vbCrLf
        
        Case "NEOSOFT"
            If strInOut = "�Կ�" Then
                SQL = ""
                SQL = SQL & "SELECT DISTINCT a.CODE as ITEM                         " & vbCr
                SQL = SQL & "  From E_ORDER..ORDER_IN" & Format(Now, "yyyy") & " a  " & vbCr
                SQL = SQL & " Where a.CHAM_INDEX =  '" & strBarcode & "'            " & vbCr
                SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                  " & vbCr
                SQL = SQL & "   AND a.TRANS = '2'                                   " & vbCr
                SQL = SQL & " ORDER BY a.CODE                                       " & vbCr
            ElseIf strInOut = "�ܷ�" Then
                SQL = ""
                SQL = SQL & "SELECT DISTINCT a.CODE as ITEM                         " & vbCr
                SQL = SQL & "  From E_ORDER..ORDER_OUT" & Format(Now, "yyyy") & " a " & vbCr
                SQL = SQL & " Where a.CHAM_INDEX =  '" & strBarcode & "'            " & vbCr
                SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                  " & vbCr
                SQL = SQL & "   AND a.TRANS = '2'                                   " & vbCr
                SQL = SQL & " ORDER BY a.CODE                                       " & vbCr
            End If
        
        Case "ONITGUM"
            SQL = ""
            SQL = SQL & "SELECT EDPSCODE     AS ITEM              " & vbCr
            SQL = SQL & "  FROM ONIT..GUMJIN_INTERFACE            " & vbCr
            SQL = SQL & " WHERE PER_GUM_NUM = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND EDPSCODE IN (" & gAllTestCd & ")  " & vbCr
            SQL = SQL & "   AND (RESULT = ''  OR RESULT IS NULL)  " & vbCr
            
        Case "ONITEMR"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT b.MAP2SEQNO   AS ITEM      " & vbCr
            SQL = SQL & "  FROM " & gSQLDB.DB & "..WAITPRSNP a      " & vbCr
            SQL = SQL & "      ," & gSQLDB.DB & "..JUN370_RESULTTB b" & vbCr
            SQL = SQL & "      ," & gSQLDB.DB & "..PEWPRSNP c       " & vbCr
            SQL = SQL & "      ," & gSQLDB.DB & "..BAGMAP2PREF d    " & vbCr
            SQL = SQL & " WHERE a.WAITSEQNO = '" & strBarcode & "'  " & vbCr
            SQL = SQL & "   AND a.JUNDAL    = '" & gHOSP.HOSPCD & "'" & vbCr        '370
            SQL = SQL & "   AND a.WAITSEQNO = b.WAITSEQNO           " & vbCr
            SQL = SQL & "   AND a.CHARTNO   = c.CHARTNO             " & vbCr
            SQL = SQL & "   AND d.LABNO     IN (" & gHOSP.LABCD & ")" & vbCr   '4
            SQL = SQL & "   AND b.MAP2SEQNO IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND b.MAP2SEQNO = d.MAP2SEQNO           " & vbCr
            SQL = SQL & "   AND (b.RESULT = '' OR b.RESULT IS NULL) " & vbCr
        
        Case "PLIS"
            If Len(strBarcode) >= 11 Then
                strSpcYY = Mid(strBarcode, 1, 2)
                strSpcNo = Mid(strBarcode, 3, 9)
            Else
                Exit Function
            End If
            
            SQL = ""
            SQL = SQL & "SELECT DISTINCT r.testcd AS ITEM        " & vbCr
            SQL = SQL & "  FROM plis..s2lab201 m                 " & vbCr
            SQL = SQL & "     , plis..s2lab302 r                 " & vbCr
            SQL = SQL & "     , plis..s2lab001 e                 " & vbCr
            SQL = SQL & " WHERE m.spcyy = '" & strSpcYY & "'     " & vbCr
            SQL = SQL & "   and m.spcno = '" & strSpcNo & "'     " & vbCr
            SQL = SQL & "   and r.testcd IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   and (r.vfydt IS NULL OR r.vfydt='')  " & vbCr
            SQL = SQL & "   and m.workarea = r.workarea          " & vbCr
            SQL = SQL & "   and m.accdt = r.accdt                " & vbCr
            SQL = SQL & "   and m.accseq = r.accseq              " & vbCr
            SQL = SQL & "   and r.testcd = e.testcd              " & vbCr
            SQL = SQL & "  Order by r.testcd                     " & vbCr
        
        Case "TWIN"
            SQL = ""
            'SQL = SQL & "SELECT DISTINCT A.MASTERCODE AS ITEM           " & vbCr
            SQL = SQL & "SELECT DISTINCT A.SUBCODE    AS ITEM           " & vbCr
            SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A             " & vbCr
            SQL = SQL & "     , TW_HSP_OCS.TWEXAM_MASTER  B             " & vbCr
            SQL = SQL & "     , TW_HSP_OCS.TWEXAM_SPECMST C             " & vbCr
            SQL = SQL & " Where A.SPECNO =  '" & strBarcode & "'        " & vbCr
            SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'     " & vbCr '����ڵ�
            SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & "   AND C.STATUS   <= '3'                       " & vbCr '�˻����
            SQL = SQL & "   And C.SPECNO  = A.SPECNO                    " & vbCr
            SQL = SQL & "   And A.MASTERCODE = B.MASTERCODE             " & vbCr
            SQL = SQL & " ORDER BY A.ITEM                               " & vbCr
                
        Case "UBCARE"
            SQL = ""
            SQL = SQL & "Select Distinct EXAMCODE AS ITEM       " & vbCr
            SQL = SQL & "  From UB_PATRESULT                    " & vbCr
            SQL = SQL & " Where BARCODE = '" & strBarcode & "'  " & vbCr
            SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   And (RESULT = '' OR RESULT IS NULL) " & vbCr
            SQL = SQL & " Order by EXAMCODE                     " & vbCr
        
            Call SetSQLData("ITEM��ȸ", SQL)
    
            '-- Record Count ������
            AdoCn_Local.CursorLocation = adUseClient
            Set RS = AdoCn_Local.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    If strItems = "" Then
                        strItems = GetTestNm(Trim(RS.Fields("ITEM")) & "", False)
                    Else
                        strItems = strItems & "/" & GetTestNm(Trim(RS.Fields("ITEM")), False)
                    End If
                    RS.MoveNext
                Loop
            End If
            
            GetSampleITEM = strItems
            
            RS.Close
            
            Exit Function
            
    End Select
            
                
    Call SetSQLData("ITEM��ȸ", SQL)
    
    If SQL <> "" Then
        '-- Record Count ������
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                If strItems = "" Then
                    strItems = GetTestNm(Trim(RS.Fields("ITEM")) & "", False)
                Else
                    strItems = strItems & "/" & GetTestNm(Trim(RS.Fields("ITEM")), False)
                End If
                RS.MoveNext
            Loop
        End If
        
        GetSampleITEM = strItems
        
        RS.Close
    Else
        GetSampleITEM = ""
    End If
    
End Function


Public Sub LetEqpMaster(ByVal pEqpCD As String)

    SQL = ""
    SQL = SQL & "UPDATE EQPMASTER SET EQUIPCD = '" & pEqpCD & "'"

    Call DBExec(AdoCn_Local, SQL)

End Sub

'-- ����� ��ȸ
Public Sub GetResultList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim strBarcode  As String
    Dim blnSame     As Boolean
    Dim strItems    As String
    Dim intOCnt     As Integer
    Dim strSaveSeq  As String
    Dim strExamDate As String

    Screen.MousePointer = 11
    iCnt = 0

    SQL = ""
    SQL = SQL & "SELECT DISTINCT SAVESEQ,EXAMDATE,EXAMTIME,HOSPDATE,BARCODE,PNAME,SENDFLAG,SENDDATE " & vbCr
    '-- �˻���
    SQL = SQL & ",SEQNO,EXAMNAME,RESULT,PREVRESULT,REFJUDGE" & vbCr

    SQL = SQL & "  FROM PATRESULT " & vbCr
    '-- �˻�����
    SQL = SQL & " WHERE EXAMDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
'    SQL = SQL & "   AND EXAMCODE IN (" & gAllTestCd & ") " & vbCr
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,BARCODE,SEQNO"

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strItems = ""
        Do Until RS.EOF
            iCnt = iCnt + 1
            With SPD
                .ReDraw = False

                strSaveSeq = GetText(SPD, intRow, colSAVESEQ)
                strExamDate = GetText(SPD, intRow, colEXAMDATE)

                If strSaveSeq <> Trim(RS.Fields("SAVESEQ")) & "" Or strExamDate <> Trim(RS.Fields("EXAMDATE")) & "" Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("SAVESEQ")) & "", intRow, colSAVESEQ
                    SetText SPD, Trim(RS.Fields("EXAMDATE")) & "", intRow, colEXAMDATE
                    SetText SPD, Trim(RS.Fields("EXAMTIME")) & "", intRow, colEXAMTIME
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME


                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                    Case "0"
                            SetText SPD, "�����", intRow, colSTATE
                    Case "1"
                            SetText SPD, "���忡��", intRow, colSTATE
                    Case "2"
                            SetText SPD, "���ۿϷ�", intRow, colSTATE
                    End Select

                    'If gEMR <> "KOMAIN" Then
                    '    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    'End If
                End If

                For intCol = colSTATE + 1 To .MaxCols
                    .Row = 0
                    .Col = intCol
                    If Trim(RS.Fields("EXAMNAME")) & "" = Trim(.Text) Then
                        SetText SPD, Trim(RS.Fields("RESULT")) & "", intRow, intCol
                        If Trim(RS.Fields("REFJUDGE")) & "" <> "" Then
                            SetForeColor SPD, intRow, intRow, intCol, intCol, 255, 0, 0
                        End If
                        Exit For
                    End If

                Next

            End With
            DoEvents

            RS.MoveNext
        Loop
        'frmMain.chkRAll.Value = "1"
    Else
        'frmMain.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
        'frmMain.chkRAll.Value = "0"
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

'    Call frmMain.GetPatTRestResult_Search(1)

    Screen.MousePointer = 0

End Sub

''-- �˻��� ���� ��������
Function GetSampleInfo(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer

    Screen.MousePointer = 11

    GetSampleInfo = -1

    Select Case gEMR
        Case "VHS"
                Call GetSampleInfo_VHS(asRow, SPD)
        
        Case "SWMC"
                Call GetSampleInfo_SWMC(asRow, SPD)
        
'        Case "HDINFO"
'                Call GetSampleInfo_HDINFO(asRow, SPD)
        
'        Case "AMIS"
'                Call GetSampleInfo_AMIS(asRow, SPD)

        Case "BIGUBCARE"
'                Call GetSampleInfo_BIGUBCARE(asRow, SPD)
'
'        Case "BIT"
'                Call GetSampleInfo_BIT(asRow, SPD)
'
'        Case "BIT70"
'                Call GetSampleInfo_BIT70(asRow, SPD)
'
'        Case "EMEDI"
'                Call GetSampleInfo_AMIS(asRow, SPD)
'
'        Case "EASYS"
'                Call GetSampleInfo_EASYS(asRow, SPD)

        Case "EHWA"
                GetSampleInfo = GetSampleInfo_EHWA(asRow, SPD)
'
'        Case "EONM"
'                Call GetSampleInfo_EONM(asRow, SPD)
'
'        Case "GINUS"
'                Call GetSampleInfo_GINUS(asRow, SPD)
'
'        Case "GSEN"
'                Call GetSampleInfo_MSINFOTEC(asRow, SPD)
'
'        Case "HWASAN"
'                Call GetSampleInfo_HWASAN(asRow, SPD)
'
'        Case "JAINCOM"
'                Call GetSampleInfo_JAINCOM(asRow, SPD)
'
'        Case "JWINFO"
'                Call GetSampleInfo_JWINFO(asRow, SPD)
'
'        Case "KCHART"
'                Call GetSampleInfo_KCHART(asRow, SPD)
'
'        Case "KOMAIN"
'                Call GetSampleInfo_KOMAIN(asRow, SPD)
'
'        Case "KYU"                  '�Ǿ���б�����
'                Call GetSampleInfo_KYU(asRow, SPD)
'
'        Case "MCC"
'                Call GetSampleInfo_MCC(asRow, SPD)
'
'        Case "MEDICHART"
'                Call GetSampleInfo_MEDICHART(asRow, SPD)
'
'        Case "MEDIIT"
'                Call GetSampleInfo_MEDIIT(asRow, SPD)
'
'        Case "MEDITOLISS"                   '�Ƹ�����
'                Call GetSampleInfo_MEDITOLISS(asRow, SPD)
'
'        Case "MOD"
'                Call GetSampleInfo_MOD(asRow, SPD)
'
'        Case "MSINFOTEC"
'                Call GetSampleInfo_MSINFOTEC(asRow, SPD)
'
'        Case "NEOSOFT"
'                Call GetSampleInfo_NEOSOFT(asRow, SPD)
'
'        Case "ONITGUM"                      '�¾�Ƽ ����
'                Call GetSampleInfo_ONITGUM(asRow, SPD)
'
'        Case "ONITEMR"                      '�¾�Ƽ EMR
'                Call GetSampleInfo_ONITEMR(asRow, SPD)
'
        Case "PHILL"
                Call GetSampleInfo_PHILL(asRow, SPD)
                
'        Case "NU"
'                Call GetSampleInfo_NU(asRow, SPD)
                
'        Case "PLIS"                      '�¾�Ƽ EMR
'                Call GetSampleInfo_PLIS(asRow, SPD)
'
'        Case "TWIN"
'                Call GetSampleInfo_TWIN(asRow, SPD)
'
'        Case "SY"
'                Call GetSampleInfo_SY(asRow, SPD)
'
'        Case "UBCARE"
'                Call GetSampleInfo_UBCARE(asRow, SPD)

    End Select


    GetSampleInfo = 1

    Screen.MousePointer = 0


End Function

'-- �˻��� ���� ��������
Function GetSampleInfo_PHILL(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_PHILL = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    strRegDate = Mid(strBarcode, 1, 8)
    lngRegNo = Val(Mid(strBarcode, 9))
    
    Screen.MousePointer = 11
          
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       P.request_date  AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , P.exam_no       AS PID      " & vbCrLf
    SQL = SQL & "     , P.company_code  AS INOUT    " & vbCrLf
    SQL = SQL & "     , P.chart_no      AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , p.personal_id               " & vbCrLf
    SQL = SQL & "     , p.person_name   AS PNAME    " & vbCrLf
    SQL = SQL & "     , P.worker_code               " & vbCrLf
    SQL = SQL & "     , P.patient_kind              " & vbCrLf
    SQL = SQL & "     , P.person_sex    AS SEX      " & vbCrLf
    SQL = SQL & "     , P.person_age    AS AGE      " & vbCrLf
    SQL = SQL & "     , R.exam_order                " & vbCrLf
    SQL = SQL & "     , R.exam_code     AS ITEM     " & vbCrLf
    SQL = SQL & "     , E.exam_ename                " & vbCrLf
    SQL = SQL & "     , R.pro_code      AS ORDERCODE            " & vbCrLf
    SQL = SQL & "  FROM trust P, trures R, examitem E           " & vbCrLf
    SQL = SQL & " WHERE P.request_date  = '" & strRegDate & "'  " & vbCrLf
    SQL = SQL & "   AND P.exam_no       = '" & lngRegNo & "'    " & vbCrLf
    SQL = SQL & "   AND R.exam_code     IN (" & gAllTestCd & ") " & vbCrLf
    SQL = SQL & "   AND R.exam_code     <> 'X999'               " & vbCrLf
    SQL = SQL & "   AND P.request_date  = R.request_date        " & vbCrLf
    SQL = SQL & "   AND P.exam_no       = R.exam_no             " & vbCrLf
    SQL = SQL & "   AND R.exam_code     = E.exam_code           " & vbCrLf
    SQL = SQL & " ORDER BY P.request_date, P.exam_no            "
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Format(Trim(RS.Fields("HOSPDATE")) & "", "####-##-##"), asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("patient_kind")) & "", asRow, colINOUT
                'SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        '-- ó���ڵ�
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDERCODE")) & ""
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                gPatTest(intTestCnt) = Trim(RS.Fields("ITEM"))
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_PHILL = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_PHILL = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_PHILL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfo_SWMC(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_SWMC = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
          
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       PART_JUBSU_DATE         AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , SPECIMEN_SER            AS BARCODE  " & vbCrLf
    SQL = SQL & "     , BUNHO                   AS PID      " & vbCrLf
    SQL = SQL & "     , FKCPL0201               AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , SPECIMEN_CODE           AS SPCCD    " & vbCrLf
    SQL = SQL & "     , SUNAME                  AS PNAME    " & vbCrLf
    SQL = SQL & "     , AGE                     AS AGE      " & vbCrLf
    SQL = SQL & "     , SEX                     AS SEX      " & vbCrLf
    SQL = SQL & "     , GWA                     AS DEPT     " & vbCrLf
    SQL = SQL & "     , HANGMOG_CODE            AS ITEM     " & vbCrLf
    SQL = SQL & "  FROM VW_CPL_INTERFACE_GUMSA_LOAD         " & vbCrLf
    SQL = SQL & " WHERE SPECIMEN_SER = '" & strBarcode & "' " & vbCrLf
    SQL = SQL & "   AND NVL(CONFIRM_YN, 'N') = 'N'          " & vbCrLf
    SQL = SQL & "   AND HANGMOG_CODE IN (" & gAllTestCd & ")" & vbCrLf
    'SQL = SQL & "   AND JANGBI_CODE = '" & gHOSP.MACHCD & "'"
    
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("SPCCD")), asRow, colSPECIMEN
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, Trim(RS.Fields("DEPT")) & "", asRow, colDEPT
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If GetTestNm(Trim(RS.Fields("ITEM"))) = gArrEQPNm(intCol - colSTATE, 6) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_SWMC = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_SWMC = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_SWMC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfo_VHS(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    '-- ���ڵ� ��ȣ�� ���� ��ȸ
    Dim Prm1            As New ADODB.Parameter
    
On Error GoTo DBErr
    
    GetSampleInfo_VHS = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
          
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandTimeout = 15
    AdoCmd.CommandText = "PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEEQP_S01"
    AdoCmd.CommandType = adCmdStoredProc
    
    Set Prm1 = AdoCmd.CreateParameter("IN_SPCNO", adVarChar, adParamInput, 11, strBarcode)
    AdoCmd.Parameters.Append Prm1
    
    Set RS = New ADODB.Recordset
    RS.Open AdoCmd.Execute
    
    Call SetSQLData("���ڵ���ȸ", "PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEEQP_S01 >> " & strBarcode)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
'    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("BLD_COL_DATE")) & "", asRow, colHOSPDATE
                SetText SPD, strBarcode, asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PT_NO")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("ACPNO_1")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PT_NAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .PID = Trim(RS.Fields("PT_NO")) & ""
                    .PNAME = Trim(RS.Fields("PT_NAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If GetTestNm(Trim(RS.Fields("EXAM_CD"))) = gArrEQPNm(intCol - colSTATE, 6) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("EXAM_CD")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
    Set AdoCmd = Nothing
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
        
        'CT�� ó����
        gPatOrdCd = gPatOrdCd & ",'XXXXX','YYYYY','ZZZZZ'"
    End If
    
    GetSampleInfo_VHS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_VHS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_VHS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- �˻��� ���� ��������
'Function GetSampleInfo_NU(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
'    Dim strRegDate      As String
'    Dim strBarcode      As String
'    Dim strPatID        As String
'    Dim strChartNo      As String
'    Dim intCol          As Integer
'    Dim intTestCnt      As Integer
'    Dim lngRegNo            As Long
'
'    Dim sParam      As String
'    Dim sRcvData    As String
'    Dim varRcvData  As Variant
'    Dim varTstCode  As Variant
'    Dim i           As Integer
'    Dim j           As Integer
'
'On Error GoTo DBErr
'
'    GetSampleInfo_NU = -1
'
'    intTestCnt = 0
'    gPatOrdCd = ""
'    ReDim Preserve gPatTest(0)
'
'    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
'    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
'    strPatID = Trim(GetText(SPD, asRow, colPID))
'    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
'
'    If strBarcode = "" Then
'        Exit Function
'    End If
'
'    Screen.MousePointer = 11
'
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00101&"                                       'submit ID
'    sParam = sParam & "business_id=li&"                                             'business_id
'    sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"     '�����ID|����ڵ�
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '����ڵ�
'    sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '����ڵ�
'    sParam = sParam & "bcno=" & strBarcode                                          '���ڵ�
'
''    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("���ڵ���ȸ", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'    If InStr(1, sRcvData, "<?xml version") > 0 Then
'        varRcvData = Split(sRcvData, "CDATA[")
'    End If
'
'    If UBound(varRcvData) >= 0 Then
'        For i = 1 To UBound(varRcvData)
'            varRcvData(i) = Mid(varRcvData(i), 1, InStr(varRcvData(i), "]") - 1)
''            Debug.Print varRcvData(i)
'        Next
'
'        For i = 1 To UBound(varRcvData) 'Step 19
'            With SPD
'                .ReDraw = False
'                intTestCnt = intTestCnt + 1
'
'                'ȯ�� ����/����
'                With mPatient
'                    .SEX = mGetP(varRcvData(6) & "", 1, "/")
'                    .AGE = mGetP(varRcvData(6) & "", 2, "/")
'                End With
'
'                SetText SPD, "1", asRow, colCHECKBOX
'                SetText SPD, Format(Mid(varRcvData(1), 1, 8), "####-##-##"), asRow, colHOSPDATE
'                SetText SPD, varRcvData(2) & "", asRow, colINOUT
'                SetText SPD, varRcvData(3) & "", asRow, colBARCODE
'                SetText SPD, varRcvData(4) & "", asRow, colPID
'                SetText SPD, varRcvData(5) & "", asRow, colPNAME
'                SetText SPD, mPatient.SEX, asRow, colPSEX
'                SetText SPD, mPatient.AGE, asRow, colPAGE
'
'                '��������
'                SetText SPD, CStr(intTestCnt), asRow, colOCNT
'
'                '���������� ����
'                With mOrder
'                    .BarNo = varRcvData(3) & ""
'                    .PID = varRcvData(4) & ""
'                    .PNAME = varRcvData(5) & ""
'                    .Count = CStr(intTestCnt)
'                    .NoOrder = False
'                End With
'
'                '-- ȭ�鿡 ǥ��
'                If Trim(varRcvData(10) & "") <> "" Then
'                    varTstCode = Split(varRcvData(11), "��")
'                    For j = 0 To UBound(varTstCode) - 1
'                        gPatOrdCd = gPatOrdCd & "'" & Trim(varTstCode(j)) & "',"
'
'                        For intCol = colSTATE + 1 To .MaxCols
'                            If Trim(varTstCode(j)) = gArrEQP(intCol - colSTATE, 2) Then
'                                .Row = asRow
'                                .Col = intCol
'                                .BackColor = vbYellow
'                                Call SetText(SPD, "��", asRow, intCol)
'                                Exit For
'                            End If
'                        Next
'
'                        gPatOrdCd = gPatOrdCd & "'" & Trim(varTstCode(j)) & "',"
'                        gPatTest(intTestCnt) = Trim(varTstCode(j))
'                    Next
'                End If
'
'            End With
'            DoEvents
'
'        Next
'    End If
'
'    RS.Close
'
'    If gPatOrdCd <> "" Then
'        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
'    End If
'
'    GetSampleInfo_NU = 1
'
'    Screen.MousePointer = 0
'
'Exit Function
'
'DBErr:
'    GetSampleInfo_NU = -1
'    intTestCnt = 0
'    Screen.MousePointer = 0
'
''    strErrMsg = ""
''    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_NU" & vbNewLine & vbNewLine
''    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
''    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
''    frmErrMsg.txtErr = vbNewLine & strErrMsg
''    frmErrMsg.Show
'
'End Function

Public Sub XmlSelect_Free()
    
    With XmlSelect
        .AGE = ""
        .BCNO = ""
        .EXECprcpuniqno = ""
        .ORDDEPTCD = ""
        .PATNM = ""
        .PID = ""
        .PRCPDD = ""
        .PRGSTNO = ""
        .RETESTYN = ""
        .RSLTSTAT = ""
        .SEX = ""
        .SPCACPTDT = ""
        .SPCCD = ""
        .SPCNM = ""
        .SPCSTAT = ""
        .TCLSCD = ""
        .TESTCD = ""
        .TESTLRGCD = ""
        .WORKNO = ""
    End With
    
End Sub

Public Sub DisplayNode_Info(asPath As String)

    Dim xmlDoc          As New MSXML2.DOMDocument30
    Dim nodeBook        As IXMLDOMElement
    Dim nodeId          As IXMLDOMAttribute
    Dim xNode           As MSXML2.IXMLDOMNode
    Dim namedNodeMap    As IXMLDOMNamedNodeMap
    Dim Child_Node      As MSXML2.IXMLDOMNodeList
    Dim i, j            As Integer
    
    On Error GoTo ErrXML:
    
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    'xmlDoc.Load "D:\������Ʈ\VB\���������������ǿ�\����\Info.xml"
    
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
        MsgBox ("You have error " & myErr.reason)
    Else
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                Exit For
            End If
        Next
        
        For i = 0 To xNode.childNodes.Item(0).childNodes.Length
            'Debug.Print xNode.childNodes.Item(0).childNodes.Item(i).baseName & ":" & xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
            Select Case UCase(xNode.childNodes.Item(0).childNodes.Item(i).baseName)
                Case "AGE":             XmlSelect.AGE = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "BCNO":            XmlSelect.BCNO = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "EXECprcpuniqno":  XmlSelect.EXECprcpuniqno = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "ORDDEPTCD":       XmlSelect.ORDDEPTCD = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "PATNM":           XmlSelect.PATNM = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "PID":             XmlSelect.PID = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "PRCPDD":          XmlSelect.PRCPDD = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "PRGSTNO":         XmlSelect.PRGSTNO = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "RETESTYN":        XmlSelect.RETESTYN = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "RSLTSTAT":        XmlSelect.RSLTSTAT = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "SEX":             XmlSelect.SEX = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "SPCACPTDT":       XmlSelect.SPCACPTDT = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "SPCCD":           XmlSelect.SPCCD = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "SPCNM":           XmlSelect.SPCNM = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "SPCSTAT":         XmlSelect.SPCSTAT = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "TCLSCD":          XmlSelect.TCLSCD = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "TESTCD":          XmlSelect.TESTCD = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "TESTLRGCD":       XmlSelect.TESTLRGCD = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
                Case "WORKNO":          XmlSelect.WORKNO = xNode.childNodes.Item(0).childNodes.Item(i).nodeTypedValue
            End Select
        Next
       
        Set Child_Node = Nothing
        
    End If

ErrXML:
    Exit Sub
    
End Sub

Public Sub DisplayNode_InfoS(asPath As String, asCnt As Integer)

    Dim xmlDoc          As New MSXML2.DOMDocument30
    Dim nodeBook        As IXMLDOMElement
    Dim nodeId          As IXMLDOMAttribute
    Dim xNode           As MSXML2.IXMLDOMNode
    Dim namedNodeMap    As IXMLDOMNamedNodeMap
    Dim Child_Node      As MSXML2.IXMLDOMNodeList
    Dim i, j            As Integer
    Dim intNodeLen      As Integer
    
On Error GoTo ErrXML:
    
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
        MsgBox ("You have error " & myErr.reason)
    Else
        ReDim Preserve XmlSelectS.AGE(asCnt)
        ReDim Preserve XmlSelectS.BCNO(asCnt)
        ReDim Preserve XmlSelectS.EXECprcpuniqno(asCnt)
        ReDim Preserve XmlSelectS.ORDDEPTCD(asCnt)
        ReDim Preserve XmlSelectS.PATNM(asCnt)
        ReDim Preserve XmlSelectS.PID(asCnt)
        ReDim Preserve XmlSelectS.PRCPDD(asCnt)
        ReDim Preserve XmlSelectS.PRGSTNO(asCnt)
        ReDim Preserve XmlSelectS.RETESTYN(asCnt)
        ReDim Preserve XmlSelectS.RSLTSTAT(asCnt)
        ReDim Preserve XmlSelectS.SEX(asCnt)
        ReDim Preserve XmlSelectS.SPCACPTDT(asCnt)
        ReDim Preserve XmlSelectS.SPCCD(asCnt)
        ReDim Preserve XmlSelectS.SPCNM(asCnt)
        ReDim Preserve XmlSelectS.SPCSTAT(asCnt)
        ReDim Preserve XmlSelectS.TCLSCD(asCnt)
        ReDim Preserve XmlSelectS.TESTCD(asCnt)
        ReDim Preserve XmlSelectS.TESTLRGCD(asCnt)
        ReDim Preserve XmlSelectS.WORKNO(asCnt)
            
        '<bcno><![CDATA[3010700030]]></bcno>
        '<patnm><![CDATA[�ڼ���]]></patnm>
        '<prgstno><![CDATA[400321-1******]]></prgstno>
        '<pid><![CDATA[000132623]]></pid>
        '<sex><![CDATA[M]]></sex>
        '<age><![CDATA[78]]></age>
        '<spcnm><![CDATA[Throat swab]]></spcnm>
        '<spccd><![CDATA[023]]></spccd>
        '<tclscd><![CDATA[VB6012A]]></tclscd>
        '<spcstat><![CDATA[4]]></spcstat>
        '<rsltstat><![CDATA[-]]></rsltstat>
        '<workno><![CDATA[20181217I20002]]></workno>
        '<testcd><![CDATA[VB6012A]]></testcd>
        '<execprcpuniqno><![CDATA[2002638354]]></execprcpuniqno>
        '<spcacptdt><![CDATA[20181217094414]]></spcacptdt>
        '<prcpdd><![CDATA[20181217]]></prcpdd>
        '<retestyn><![CDATA[N]]></retestyn>
        '<testlrgcd><![CDATA[I]]></testlrgcd>
        '<orddeptcd><![CDATA[NU]]></orddeptcd>
        
        
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                For intNodeLen = 0 To xNode.childNodes.Length - 1
                    For i = 0 To xNode.childNodes.Item(intNodeLen).childNodes.Length - 1
                        'Debug.Print xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName & ":" & xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue
                        Select Case UCase(xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName)
                            Case "AGE":             XmlSelectS.AGE(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '����           [78]
                            Case "BCNO":            XmlSelectS.BCNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                '��ü��ȣ       [3010700030]
                            Case "EXECprcpuniqno":  XmlSelectS.EXECprcpuniqno(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue      'íƮ��ȣ?      [2002638354]
                            Case "ORDDEPTCD":       XmlSelectS.ORDDEPTCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           'ó��μ��ڵ�?  [NU]
                            
                            Case "PATNM":           XmlSelectS.PATNM(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               'ȯ�ڸ�         [�ڼ���]
                            Case "PID":             XmlSelectS.PID(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 'ȯ�ڹ�ȣ       [000132623]
                            Case "PRCPDD":          XmlSelectS.PRCPDD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              'ó����?        [20181217]
                            Case "PRGSTNO":         XmlSelectS.PRGSTNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '�ֹι�ȣ       [400321-1******]
                            
                            Case "RETESTYN":        XmlSelectS.RETESTYN(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue            '��˿���       [N]
                            Case "RSLTSTAT":        XmlSelectS.RSLTSTAT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue            '�������       [-]
                            
                            Case "SEX":             XmlSelectS.SEX(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '����           [M]
                            
                            Case "SPCACPTDT":       XmlSelectS.SPCACPTDT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '��üä��ð�?  [20181217094414]
                            Case "SPCCD":           XmlSelectS.SPCCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '��ü�ڵ�       [023]
                            Case "SPCNM":           XmlSelectS.SPCNM(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '��ü��         [Throat swab]
                            Case "SPCSTAT":         XmlSelectS.SPCSTAT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '��ü����       [4]
                            
                            Case "TCLSCD":          XmlSelectS.TCLSCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              'ó���ڵ�       [VB6012A]
                            Case "TESTCD":          XmlSelectS.TESTCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '�˻��ڵ�       [VB6012A]
                            Case "TESTLRGCD":       XmlSelectS.TESTLRGCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '����׷��ڵ�?  [I]
                            Case "WORKNO":          XmlSelectS.WORKNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '��ũ��ȣ       [20181217I20002]
                        End Select
                    Next
                    j = j + 1
                Next
            End If
        Next
       
        Set Child_Node = Nothing
        
    End If

    Exit Sub
    
ErrXML:
    Exit Sub
    
End Sub

'-- �˻��� ���� ��������
'Function GetSampleInfo_HDINFO(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
'    Dim strRegDate      As String
'    Dim strBarcode      As String
'    Dim strPatID        As String
'    Dim strChartNo      As String
'    Dim intCol          As Integer
'    Dim intTestCnt      As Integer
'    Dim lngRegNo            As Long
'
'    Dim sParam      As String
'    Dim sRcvData    As String
'    Dim varRcvData  As Variant
'    Dim varTstCode  As Variant
'    Dim i           As Integer
'    Dim j           As Integer
'    Dim strXmlName  As String
'    Dim strNames    As String
'
'On Error GoTo DBErr
'
'    GetSampleInfo_HDINFO = -1
'
'    intTestCnt = 0
'    gPatOrdCd = ""
'    ReDim Preserve gPatTest(0)
'
'    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
'    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
'    strPatID = Trim(GetText(SPD, asRow, colPID))
'    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
'
'    If strBarcode = "" Then
'        Exit Function
'    End If
'
'    Screen.MousePointer = 11
'
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00123&"                                   'submit ID
'    sParam = sParam & "business_id=lis&"                                        'business_id
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '����ڵ�
'    sParam = sParam & "bcno=" & strBarcode                                      '��ü��ȣ(=���ڵ�)
'
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("���ڵ���ȸ", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'
'    If InStr(1, sRcvData, "<?xml version") > 0 Then
'        varRcvData = Split(sRcvData, "<worklist>")
'    End If
'
'    strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & "_" & strBarcode & ".xml"
'
'    Call SetXMLData(strXmlName, sRcvData)
'
'    Call DisplayNode_InfoS(App.PATH & "\Xml\" & strXmlName, UBound(varRcvData))
'
'    Kill App.PATH & "\Xml\" & strXmlName
'
'
'    If UBound(varRcvData) >= 1 Then
'        For i = 0 To UBound(varRcvData) - 1 'Step 19
'            With SPD
'                .ReDraw = False
'
'                intTestCnt = intTestCnt + 1
'
'                'ȯ�� ����/����
'                With mPatient
'                    .SEX = XmlSelectS.SEX(i)
'                    .AGE = XmlSelectS.AGE(i)
'                End With
'
'               ' blnSame = False
'
'                'If blnSame = False Then
'                    SetText SPD, "1", asRow, colCHECKBOX
'                    SetText SPD, XmlSelectS.PRCPDD(i), asRow, colHOSPDATE
'                    'SetText SPD, varRcvData(i + 1) & "", asRow, colINOUT
'                    SetText SPD, XmlSelectS.BCNO(i), asRow, colBARCODE
'                    SetText SPD, XmlSelectS.PID(i), asRow, colPID
'                    SetText SPD, XmlSelectS.PATNM(i), asRow, colPNAME
'                    SetText SPD, XmlSelectS.SEX(i), asRow, colPSEX
'                    SetText SPD, XmlSelectS.AGE(i), asRow, colPAGE
'                    SetText SPD, XmlSelectS.SPCNM(i), asRow, colSPECIMEN
'                    'SetText SPD, varRcvData(i + 6) & "", intRow, colOCNT
'                    'SetText SPD, varRcvData(i + 7) & "", intRow, colCHARTNO
'                    'SetText SPD, varRcvData(i + 8) & "", intRow, colOCNT
'
'                    For intCol = colSTATE + 1 To .MaxCols
'                        'If XmlSelectS.TESTCD(i) = gArrEQP(intCol - colSTATE, 2) Then
'
'                        If GetTestNm(Trim(XmlSelectS.TESTCD(i))) = gArrEQPNm(intCol - colSTATE, 6) Then
'                            .Row = asRow
'                            .Col = intCol
'                            .BackColor = vbYellow
'                            Call SetText(SPD, "��", asRow, intCol)
'                            Exit For
'                        End If
'                    Next
'
'                    gPatOrdCd = gPatOrdCd & "'" & XmlSelectS.TESTCD(i) & "',"
'
'                'End If
'            End With
'        Next
'    Else
'        'MsgBox "��ȸ ����ڰ� �����ϴ�.", vbOKOnly + vbCritical, "��ũ����Ʈ ��ȸ"
'    End If
'
'    If gPatOrdCd <> "" Then
'        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
'    End If
'
'
'    'MsgBox gPatOrdCd
'
'    GetSampleInfo_HDINFO = 1
'
'    Screen.MousePointer = 0
'
'Exit Function
'
'DBErr:
'    GetSampleInfo_HDINFO = -1
'    intTestCnt = 0
'    Screen.MousePointer = 0
'
''    strErrMsg = ""
''    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_NU" & vbNewLine & vbNewLine
''    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
''    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
''    frmErrMsg.txtErr = vbNewLine & strErrMsg
''    frmErrMsg.Show
'
'End Function

'-- �˻��� ���� ��������
'Function GetSampleTest_HDINFO(ByVal asBarcode As String) As Integer
'    Dim sParam      As String
'    Dim sRcvData    As String
'    Dim varRcvData  As Variant
'    Dim strXmlName  As String
'
'On Error GoTo DBErr
'
'    GetSampleTest_HDINFO = -1
'
'    If asBarcode = "" Then
'        Exit Function
'    End If
'
'    Screen.MousePointer = 11
'
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00123&"                                   'submit ID
'    sParam = sParam & "business_id=lis&"                                        'business_id
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '����ڵ�
'    sParam = sParam & "bcno=" & asBarcode                                      '��ü��ȣ(=���ڵ�)
'
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("���ڵ��׽�Ʈ��ȸ", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'    If InStr(1, sRcvData, "<?xml version") > 0 Then
'        varRcvData = Split(sRcvData, "<worklist>")
'    End If
'
'    strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & "_" & asBarcode & ".xml"
'
'    Call SetXMLData(strXmlName, sRcvData)
'
'    Call DisplayNode_InfoS(App.PATH & "\Xml\" & strXmlName, UBound(varRcvData))
'
'    Kill App.PATH & "\Xml\" & strXmlName
'
'    If UBound(varRcvData) >= 1 Then
'        GetSampleTest_HDINFO = 1
'    End If
'
'
'    Screen.MousePointer = 0
'
'Exit Function
'
'DBErr:
'    GetSampleTest_HDINFO = -1
'    Screen.MousePointer = 0
'
'End Function

'-- �˻��� ���� ��������
Function GetSampleInfo_EHWA(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
    Dim sParam      As String
    Dim sRcvData    As String
    Dim varRcvData  As Variant
    Dim varTstCode  As Variant
    Dim i           As Integer
    Dim j           As Integer
    
    Dim sRes        As String
    Dim strHospGbn  As String
    
On Error GoTo DBErr
    
    GetSampleInfo_EHWA = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    ReDim Preserve gPatTest(0)
    
    SetText SPD, "0", asRow, colCHECKBOX
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    'MsgBox strBarcode
    If strBarcode = "" Then
        Exit Function
    End If
        
    gHospCode = "02"
    
    '���ﺴ���� ���ڵ尡 ���Ϸ� ���۵ǰ�, �񵿺����� ���ڵ尡 ����Ϸ� ���۵ȴ�. �񵿹��ڵ�� ������ 13 �̻��̴�!
    If Len(strBarcode) = 11 And IsNumeric(strBarcode) Then
        strHospGbn = Mid(strBarcode, 1, 2)
        If CCur(strHospGbn) > 12 Then
            gHospCode = "02"      '�̴�񵿺���
        Else
            gHospCode = "01"      '�̴뼭�ﺴ��
        End If
    End If
    
    Screen.MousePointer = 11
  
    sRes = Online_XML(gXml_ORDER_SELECT, strBarcode, "GETQUERY", "", "") ' "PKG_MSE_LM_INTERFACE.PC_MSE_ORDER_SELECT"
  
'    sRes = Online_XML(gXml_LOGIN, "", "GETQUERY", txtID.Text, txtPW.Text) ' "PKG_MSE_LM_INTERFACE.PC_MSE_ORDER_SELECT"
  
  
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00101&"                                       'submit ID
'    sParam = sParam & "business_id=li&"                                             'business_id
'    sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"     '�����ID|����ڵ�
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '����ڵ�
'    sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '����ڵ�
'    sParam = sParam & "bcno=" & strBarcode                                          '���ڵ�
        
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("���ڵ���ȸ", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'    If InStr(1, sRcvData, "<?xml version") > 0 Then
'        varRcvData = Split(sRcvData, "CDATA[")
'    End If

    If sRes <> "" Then
'        For i = 1 To UBound(varRcvData)
'            varRcvData(i) = Mid(varRcvData(i), 1, InStr(varRcvData(i), "]") - 1)
''            Debug.Print varRcvData(i)
'        Next
'
        For i = 0 To giIndex
            With SPD
                .ReDraw = False
                
                'ȯ�� ����/����
                With mPatient
                    .SEX = gPatInfo_Select.SEX_TP_CD
                    .AGE = gPatInfo_Select.PT_BRDY_DT
                End With

                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, gPatInfo_Select.ACPT_DTM, asRow, colHOSPDATE
                SetText SPD, gPatInfo_Select.PT_HME_DEPT_CD, asRow, colINOUT
                SetText SPD, strBarcode, asRow, colBARCODE
                SetText SPD, gPatInfo_Select.TH1_SPCM_CD, asRow, colSPECIMEN
                SetText SPD, gPatInfo_Select.PT_NO, asRow, colPID
                SetText SPD, gPatInfo_Select.PT_NM, asRow, colPNAME
                SetText SPD, gPatInfo_Select.SEX_TP_CD, asRow, colPSEX
                SetText SPD, gPatInfo_Select.PT_BRDY_DT, asRow, colPAGE
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT

                '���������� ����
                With mOrder
                    .BarNo = strBarcode
                    .PID = gPatInfo_Select.PT_NO
                    .PNAME = gPatInfo_Select.PT_NM
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With

                '-- ȭ�鿡 ǥ��
                'If Trim(varRcvData(10) & "") <> "" Then
                    For intCol = colSTATE + 1 To .MaxCols
                        If gExam_Select(i).TST_CD = gArrEQP(intCol - colSTATE, 2) Then
                            .Row = asRow
                            .Col = intCol
                            .BackColor = vbYellow
                            'Call SetText(SPD, "��", asRow, intCol)
                            Exit For
                        End If
                    Next
                'End If
                
            End With
            DoEvents
            
        Next
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_EHWA = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_EHWA = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_NU" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
    
End Function

Function SetJudge(asResult As String, asEquipCode As String)

    Select Case gEMR
        Case "AMIS"                         '�ƹ̽�
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EMEDI"                        '�̸޵�
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "BIT"                          '��Ʈ
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)

        Case "BIT70"                        '��Ʈ HIB70
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EASYS"                        '������
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EONM"                         '�̿¿�
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "GSEN"                         '����Ŀ�´����̼���(��íƮ)
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "HWASAN"                       'ȭ��
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JAINCOM"                       '������
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JWINFO"                       '�߿�����
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KCHART"                       '�ٴ����Ʈ
                SetJudge = SetJudge_KCHART(asResult, asEquipCode)
        
        Case "KOMAIN"                       '�߿�����
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KYU"                          '�Ǿ���б�����
                '��ũ����Ʈ ��ɾ���
                'SetJudge =  SetJudge_KYU(asResult,asEquipCode)
        Case "MEDICHART"                    '�޵�íƮ
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDIIT"
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDITOLISS"                    '
                SetJudge = SetJudge_MEDITOLISS(asResult, asEquipCode)
            
        Case "MSINFOTEC"                    'MS������
                SetJudge = SetJudge_MSINFOTEC(asResult, asEquipCode)
                
    End Select
    
End Function

Function SetJudge_LOCAL(asResult As String, asEquipCode As String)
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sLVal As String
    Dim sHVal As String
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResFlag As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    If Not IsNumeric(sEquipRes) Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT REFLOW, REFHIGH                     " & vbCr
    SQL = SQL & "  FROM EQPMASTER                           " & vbCr
    SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'  " & vbCr

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If IsNumeric(Trim(RS_L.Fields("REFLOW")) & "") = True And IsNumeric(Trim(RS_L.Fields("REFHIGH")) & "") = True Then
            sLVal = Trim(RS_L.Fields("REFLOW")) & ""
            sHVal = Trim(RS_L.Fields("REFHIGH")) & ""
            If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
                sResFlag = ""
            ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
                sResFlag = "H"
            ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
                sResFlag = "L"
            End If
        End If
    End If
 
    SetJudge_LOCAL = sResFlag
    
End Function

Function SetJudge_EASYS(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim strLow      As String
    Dim strHigh     As String
    
    SetJudge_EASYS = ""
    
          SQL = "Select REFLOW, REFHIGH                     " & vbCr
    SQL = SQL & "  From EQPMASTER                           " & vbCr
    SQL = SQL & " Where EQUIPCD  = '" & gHOSP.MACHCD & "'   " & vbCr
    SQL = SQL & "   And TESTCODE = '" & asEquipCode & "'    " & vbCr
    
    Set RSJ = New ADODB.Recordset
    Set RSJ = AdoCn_Local.Execute(SQL, , 1)
    If Not RSJ.EOF = True And Not RSJ.BOF = True Then
        strLow = Trim(RSJ.Fields("REFLOW") & "")
        strHigh = Trim(RSJ.Fields("REFHIGH") & "")
        
        If strLow <> "" And strHigh <> "" And asResult <> "" And IsNumeric(strLow) And IsNumeric(strHigh) And IsNumeric(asResult) Then
            If Val(asResult) > Val(strHigh) Then
                SetJudge_EASYS = "H"
            ElseIf Val(asResult) < Val(strLow) Then
                SetJudge_EASYS = "L"
            Else
                SetJudge_EASYS = " "
            End If
        Else
            SetJudge_EASYS = " "
        End If
    Else
        SetJudge_EASYS = ""
    End If
        
    RSJ.Close

End Function

Function SetJudge_MSINFOTEC(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim strAge      As String
    Dim strSex      As String
    Dim stryy, strmm, strdd, strDate  As String
    
On Error GoTo ErrorTrap
    
    SetJudge_MSINFOTEC = ""
    
    asResult = Replace(asResult, "<", "")
    asResult = Replace(asResult, ">", "")
    
    strAge = mPatient.AGE
    strSex = mPatient.SEX
    
    If strAge <> "" Then
        If strAge <= 7 Then
            SQL = "Select YMAX as MAX, YMIN as MIN "
        Else
            If strSex = "M" Then
                     SQL = "Select MMAX as MAX, MMIN as MIN "
            Else
                     SQL = "Select WMAX as MAX, WMIN as MIN "
            End If
        End If
    Else
        SQL = "Select MMAX as MAX, MMIN as MIN "
    End If
    
    SQL = SQL & "  From emr.LABMAST"
    SQL = SQL & " Where ORCD =  '" & asEquipCode & "'"
    
    Set RSJ = AdoCn.Execute(SQL)
    Do Until RSJ.EOF
        If IsNumeric(asResult) And IsNumeric(RSJ.Fields("MAX")) And IsNumeric(RSJ.Fields("MIN")) Then
            If Val(asResult) > Val(RSJ.Fields("MAX")) Then
                SetJudge_MSINFOTEC = "H"
            ElseIf Val(asResult) < Val(RSJ.Fields("MIN")) Then
                SetJudge_MSINFOTEC = "L"
            Else
                SetJudge_MSINFOTEC = " "
            End If
        Else
            SetJudge_MSINFOTEC = " "
        End If
        RSJ.MoveNext
    
    Loop
    
    RSJ.Close

Exit Function

ErrorTrap:
    SetJudge_MSINFOTEC = ""
    
End Function

Function SetJudge_MEDITOLISS(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim strRefVal   As String
    
On Error GoTo ErrorTrap
    
    SetJudge_MEDITOLISS = ""
    
    SQL = ""
    SQL = SQL & "SELECT REFER_VALUE                                 " & vbCr
    SQL = SQL & "  FROM MEDITOLISS..TOTRES                          " & vbCr
    SQL = SQL & " WHERE REQUEST_DATE    = '" & mResult.RsltDate & "'" & vbCr
    SQL = SQL & "   AND EXAM_NO         = '" & mResult.BarNo & "'   " & vbCr
    SQL = SQL & "   AND EXAM_CODE       = '" & asEquipCode & "'     " & vbCr
    
    Set RSJ = AdoCn.Execute(SQL)
    Do Until RSJ.EOF
        strRefVal = RSJ.Fields("REFER_VALUE").Value & ""
        If IsNumeric(asResult) And Len(strRefVal) > 0 Then
            If Val(Trim$(asResult)) < Val(Mid(strRefVal, 1, InStr(strRefVal, "~") - 1)) Then
                SetJudge_MEDITOLISS = "L"
            ElseIf Val(Trim$(asResult)) > Val(Mid(strRefVal, InStr(strRefVal, "~") + 1)) Then
                SetJudge_MEDITOLISS = "H"
            Else
                SetJudge_MEDITOLISS = ""
            End If
        End If
    Loop
                
    RSJ.Close
    
Exit Function

ErrorTrap:
    SetJudge_MEDITOLISS = ""
    
End Function

Function SetJudge_KCHART(asResult As String, asEquipCode As String) As String
    Dim RS1         As ADODB.Recordset
    Dim sEquipCode  As String
    Dim sEquipRes   As String
    Dim sResFlag    As String
    Dim strRefL     As String
    Dim strRefH     As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    strRefL = ""
    strRefH = ""
    
'    SQL = SQL & "  L.����˻�ID AS R, " & vbCrLf
'    SQL = SQL & "  L.��������ID AS P, " & vbCrLf

    '���γ� ����ġ0~����ġ1,
    '���ο� ����ġ2~����ġ3,
    '�ҾƳ� ����ġ4~����ġ5,
    '�Ҿƿ� ����ġ6~����ġ7
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       A.ȯ�ڼ��� AS ����                                          " & vbCr
    SQL = SQL & "     , L.����ġ0, L.����ġ1, L.����ġ2, L.����ġ3                  " & vbCr
    SQL = SQL & "     , L.����ġ4, L.����ġ5, L.����ġ6, L.����ġ7                  " & vbCr
    SQL = SQL & "     , (L.ó���ڵ� + L.�����ڵ�) AS ITEM                           " & vbCr
    SQL = SQL & "  FROM             TB_����˻� L                                   " & vbCr
    SQL = SQL & "       INNER JOIN  TB_�������� J ON (L.��������ID = J.��������ID)  " & vbCr
    SQL = SQL & "       INNER JOIN  TB_�����Ϲ� A ON (J.��������   = A.��������     " & vbCr
    SQL = SQL & "                                AND  J.íƮ��ȣ   = A.íƮ��ȣ     " & vbCr
    SQL = SQL & "                                AND  J.�����ȣ   = A.�����ȣ)    " & vbCr
    SQL = SQL & "  Where L.��ü��ȣ = '" & mResult.BarNo & "'                       " & vbCr
    SQL = SQL & "    AND L.�˻���� < 5                                             " & vbCr
    SQL = SQL & "    AND (L.ó���ڵ� + L.�����ڵ�) = '" & sEquipCode & "'           " & vbCr
                                                                 

     Call SetSQLData("����ġ��ȸ", SQL)
     
     '-- Record Count ������
     AdoCn.CursorLocation = adUseClient
     Set RS1 = AdoCn.Execute(SQL, , 1)
     If Not RS1.EOF = True And Not RS1.BOF = True Then
         Do Until RS1.EOF
            strRefL = ""
            strRefH = ""
            If Trim(RS1.Fields("����")) & "" = "M" Then
                If Trim(RS1.Fields("����ġ0")) & "" <> "" Then
                    strRefL = Trim(RS1.Fields("����ġ0")) & ""
                    strRefH = Trim(RS1.Fields("����ġ1")) & ""
                End If
            Else
                If Trim(RS1.Fields("����")) & "" = "F" Then
                    If Trim(RS1.Fields("����ġ2")) & "" <> "" Then
                        strRefL = Trim(RS1.Fields("����ġ2")) & ""
                        strRefH = Trim(RS1.Fields("����ġ3")) & ""
                    Else
                        strRefL = Trim(RS1.Fields("����ġ0")) & ""
                        strRefH = Trim(RS1.Fields("����ġ1")) & ""
                    End If
                End If
            End If
            RS1.MoveNext
        Loop
    
        If IsNumeric(sEquipRes) And IsNumeric(strRefL) = True And IsNumeric(strRefH) = True Then
            If CCur(sEquipRes) > CCur(strRefL) And CCur(sEquipRes) < CCur(strRefH) Then
                sResFlag = ""
            ElseIf CCur(strRefH) <= CCur(sEquipRes) Then
                sResFlag = "H"
            ElseIf CCur(strRefL) >= CCur(sEquipRes) Then
                sResFlag = "L"
            End If
        End If
    End If
    
    RS1.Clone
    
    SetJudge_KCHART = sResFlag
    
End Function


Function SetResult(asResult As String, asEquipCode As String)
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
    Dim sPoint As Integer
    Dim sResType As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT RESPREC, REFLOW, REFHIGH " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'"

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If IsNumeric(Trim(RS_L.Fields("RESPREC")) & "") = True Then
            sPoint = CInt(Trim(RS_L.Fields("RESPREC")))
            sResType = ""
            For i = 0 To sPoint
                If i = 0 Then
                    sResType = "#0"
                ElseIf i = 1 Then
                    sResType = sResType & ".0"
                Else
                    sResType = sResType & "0"
                End If
            Next
            sResult = Format(sEquipRes, sResType)
        Else
            sResult = sEquipRes
        End If
    End If
 
    SetResult = sResult
    
End Function

Function SetCutOffResult(asResult As String, asEquipCode As String, asTestCode As String) As String
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
'    Dim sPoint As Integer
'    Dim sResType As String
    
    Dim dblLow      As Double
    Dim dblHigh     As Double
    Dim strLComp    As String
    Dim strHComp    As String
    
    sResult = ""
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT RESULTTYPE, COLIN, COLCOMP, COLOUT, COHIN, COHCOMP, COHOUT, COMOUT   " & vbCr
    SQL = SQL & "  FROM EQPMASTER                                                " & vbCr
    SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                     " & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'                       " & vbCr
    SQL = SQL & "   AND TESTCODE    = '" & asTestCode & "'                       " & vbCr

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If Trim(RS_L.Fields("COLCOMP") & "") <> "" And Trim(RS_L.Fields("COLIN") & "") <> "" Then
            If IsNumeric(Trim(RS_L.Fields("COLIN") & "")) Then
                dblLow = CCur(RS_L.Fields("COLIN"))
                strLComp = Trim(RS_L.Fields("COLCOMP"))
                If strLComp = "<" Then
                    If CCur(asResult) < dblLow Then
                        sResult = Trim(RS_L.Fields("COLOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                ElseIf strLComp = "<=" Then
                    If CCur(asResult) <= dblLow Then
                        sResult = Trim(RS_L.Fields("COLOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                End If
            End If
        ElseIf Trim(RS_L.Fields("COHCOMP") & "") <> "" And Trim(RS_L.Fields("COHIN") & "") <> "" Then
            If IsNumeric(Trim(RS_L.Fields("COHIN") & "")) Then
                dblHigh = CCur(RS_L.Fields("COHIN"))
                strHComp = Trim(RS_L.Fields("COHCOMP"))
                If strHComp = ">" Then
                    If CCur(asResult) < dblLow Then
                        sResult = Trim(RS_L.Fields("COHOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                ElseIf strHComp = ">=" Then
                    If CCur(asResult) >= dblHigh Then
                        sResult = Trim(RS_L.Fields("COHOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                End If
            End If
        End If
    End If
    
    If sResult <> "" Then
        Select Case Trim(RS_L.Fields("RESULTTYPE") & "")
            Case "���Ծ���"
                    sResult = Trim(asResult)
            Case "����"
                    sResult = Trim(asResult)
            Case "����"
                    sResult = Trim(sResult)
            Case "����(����)"
                    sResult = asResult & "(" & Trim(sResult) & ")"
            Case "����(����)"
                    sResult = sResult & "(" & Trim(asResult) & ")"
        End Select
    End If
    
    RS_L.Close
    
    SetCutOffResult = sResult
    
End Function


Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String

    With frmMain
        sExamDate = Format(Now, "yyyymmdd")
        If Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) = "" Then
            Exit Function
        End If

        SQL = ""
        SQL = SQL & "DELETE FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO     = '" & gHOSP.MACHCD & "' " & vbCrLf
        SQL = SQL & "   AND EXAMDATE    = '" & Trim(GetText(.spdOrder, asRow1, colEXAMDATE)) & "' " & vbCrLf
        SQL = SQL & "   AND EXAMTIME    = '" & Trim(GetText(.spdOrder, asRow1, colEXAMTIME)) & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) & vbCrLf
        SQL = SQL & "   AND HOSPDATE    = '" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "' " & vbCrLf
        SQL = SQL & "   AND EQUIPCODE   = '" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'" & vbCrLf
        SQL = SQL & "   AND EXAMCODE    = '" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'" & vbCrLf

        If DBExec(AdoCn_Local, SQL) Then
            SQL = ""
            SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
            SQL = SQL & "  EQUIPNO"                         '����ڵ�
            SQL = SQL & ", EXAMDATE"                        '�˻�����
            SQL = SQL & ", EXAMTIME"                        '�˻�ð�
            SQL = SQL & ", SAVESEQ"                         '�������(��¥��)
            SQL = SQL & ", HOSPDATE" & vbCrLf               '������������
            
            SQL = SQL & ", BARCODE"                         '��ü��ȣ
            SQL = SQL & ", PID"                             '���Ϲ�ȣ(������ȣ)
            SQL = SQL & ", CHARTNO"                         'íƮ��ȣ
            SQL = SQL & ", SPECIMEN"                        '��ü
            SQL = SQL & ", DEPT" & vbCrLf                   '�Ƿڰ�
            
            SQL = SQL & ", INOUT"                           '��/��
            SQL = SQL & ", ERYN"                            '���޿���
            SQL = SQL & ", RETESTYN"                        '��˿���
            SQL = SQL & ", PNAME"                           '�̸�
            SQL = SQL & ", PSEX" & vbCrLf                   '����(M,F)
            
            SQL = SQL & ", PAGE"                            '����
            SQL = SQL & ", EXAMUID"                         '�˻���ID
            SQL = SQL & ", DISKNO"                          'Rack
            SQL = SQL & ", POSNO"                           'Pos
            SQL = SQL & ", EQPSEQNO" & vbCrLf               '���˻��ȣ
            '============================================================
            
            SQL = SQL & ", SEQNO"                           '�˻��Ϸù�ȣ
            SQL = SQL & ", EQUIPCODE"                       '�˻�ä��
            SQL = SQL & ", ORDERCODE"                       '����ó���ڵ�
            SQL = SQL & ", EXAMCODE"                        '�����˻��ڵ�
            SQL = SQL & ", EXAMCODESUB" & vbCrLf            '�����˻��ڵ�(SUB)"
            
            SQL = SQL & ", EXAMNAME"                        '�˻��
            SQL = SQL & ", EQUIPRESULT"                     '�����"
            SQL = SQL & ", RESULT"                          '�Ҽ�������"
            SQL = SQL & ", PREVRESULT"                      '�������"
            SQL = SQL & ", REFJUDGE" & vbCrLf               '����(H,L)
            
            SQL = SQL & ", REFFLAG"                         'flag
            SQL = SQL & ", REFVALUE"                        '����ġ
            SQL = SQL & ", PANICVALUE"                      'Delta
            SQL = SQL & ", DELTAVALUE"                      'Panic
            SQL = SQL & ", SENDFLAG"                        '���۱���(0:������,1:����)"
            SQL = SQL & ", SENDDATE)" & vbCrLf               '��������
            
            SQL = SQL & " VALUES (" & vbCrLf
            SQL = SQL & "'" & gHOSP.MACHCD & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colEXAMDATE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colEXAMTIME)) & "'"
            SQL = SQL & "," & Trim(GetText(.spdOrder, asRow1, colSAVESEQ))
            SQL = SQL & ",'" & Mid(Trim(GetText(.spdOrder, asRow1, colHOSPDATE)), 1, 10) & "'" & vbCrLf
            
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPID)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colCHARTNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colSPECIMEN)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colDEPT)) & "'" & vbCrLf
            
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colINOUT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colER)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colRT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPNAME)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPSEX)) & "'" & vbCrLf
            
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPAGE)) & "'"
            SQL = SQL & ",'" & gHOSP.USERID & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colRACKNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPOSNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colSEQNO)) & "'" & vbCrLf
            '============================================================
            
            SQL = SQL & "," & Trim(GetText(.spdResult, asRow2, colRSEQNO))
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRORDERCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRSUBCD)) & "'" & vbCrLf
            
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTNM)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRMACHRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRLISRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRPREVRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRJUDGE)) & "'" & vbCrLf
            
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRFLAG)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRREF)) & "'"
            SQL = SQL & ",''"
            SQL = SQL & ",''"
            SQL = SQL & ",'0'"
            SQL = SQL & ",'')" & vbCrLf
            
            'SetRawData "[��������]" & SQL
            
            If Not DBExec(AdoCn_Local, SQL) Then
                Exit Function
            End If

        End If
    End With

End Function

'-- ���� �˻��� ��¥�� Max + 1 ��ȣ�� �����´�
Public Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1

          SQL = "SELECT MAX(SAVESEQ) as SEQ FROM PATRESULT  "
    SQL = SQL & " WHERE EXAMDATE = '" & strDate & "' " & vbCrLf

    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        If Trim(RS.Fields("SEQ") & "") = "" Then
            getMaxTestNum = 1
        Else
            getMaxTestNum = Trim(RS.Fields("SEQ")) + 1
        End If
    Else
        getMaxTestNum = 1
    End If

    If getMaxTestNum >= 99999 Then
        getMaxTestNum = 99999
    End If

    RS.Close

End Function
'
