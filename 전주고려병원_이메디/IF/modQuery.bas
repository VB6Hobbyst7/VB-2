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
    
    Dim strEqpcd        As String
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
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
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
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_MCC(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim dblBarno        As Double
    Dim prm1            As New ADODB.Parameter
    Dim prm2            As New ADODB.Parameter
    Dim prm3            As New ADODB.Parameter
    Dim prm4            As New ADODB.Parameter
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
        SaveTransData_MCC = -1
        
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
        
        If IsNumeric(strBarcode) Then
            dblBarno = CDbl(strBarcode)
        Else
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
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
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
                    SQL = ""
                    SQL = SQL & "Exec UP_LIS_INTERFACE_U$001 " & dblBarno & "," & strTestCd & "," & sResult & "," & gHOSP.MACHCD
                        
                    Set AdoCmd = New ADODB.Command
                    Set AdoCmd.ActiveConnection = AdoCn
                    With AdoCmd
                        .CommandTimeout = 15
                        .CommandText = "UP_LIS_INTERFACE_U$001"
                        .CommandType = adCmdStoredProc
                        
                        Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, dblBarno)      '���ڵ��ȣ
                        .Parameters.Append prm1
    
                        Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       'ó���ڵ�
                        .Parameters.Append prm2
    
                        Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '�����
                        .Parameters.Append prm3
    
                        Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '����ڵ�
                        .Parameters.Append prm4
    
                        .Execute
                        
                    End With
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_MCC = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_MCC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


Function SaveTransData_TWIN(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim dblBarno        As Double
    Dim prm1            As New ADODB.Parameter
    Dim prm2            As New ADODB.Parameter
    Dim prm3            As New ADODB.Parameter
    Dim prm4            As New ADODB.Parameter
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
        SaveTransData_TWIN = -1
        
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
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
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
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "Update TW_HSP_OCS.TWEXAM_RESULTC           " & vbCrLf
                    SQL = SQL & "   Set STATUS      = '4'                   " & vbCrLf  '�˻����
                    SQL = SQL & "     , RESULT      = '" & sResult & "'     " & vbCrLf  '�˻���
                    SQL = SQL & "     , RESULTDATE  = TRUNC(SYSDATE)        " & vbCrLf  '�˻����۽ð�
                    SQL = SQL & " Where SPECNO      = '" & strBarcode & "'  " & vbCrLf  '��ü��ȣ
'                    SQL = SQL & "   And MASTERCODE  = 'LH1P01'    " & vbCrLf  '�������ڵ� LH1P01
                    SQL = SQL & "   And SUBCODE     = '" & strTestCd & "'   " & vbCrLf  '�˻��ڵ�
                    SQL = SQL & "   And STATUS      <= '3'                  " & vbCrLf  '�˻����(=��ü����)
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                
                    '-- ���¾�����Ʈ
                    SQL = ""
                    SQL = SQL & "Update TW_HSP_OCS.TWEXAM_SPECMST           " & vbCrLf
                    SQL = SQL & "   Set STATUS     = '3'                    " & vbCrLf '�˻���� [������(3:�����Ȯ��, 4:�κ�����)]
                    SQL = SQL & "     , RESULTDATE = TRUNC(SYSDATE)         " & vbCrLf
                    SQL = SQL & " Where SPECNO     = '" & strBarcode & "'   " & vbCrLf '��ü��ȣ
                    SQL = SQL & "   And STATUS     <= '3'                   " & vbCrLf '�˻���� [3:��ü����]
                    
                    Call SetSQLData("��������", SQL, "A")
                    AdoCn.Execute SQL
                    
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_TWIN = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_TWIN = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_TWIN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_EMEDI(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim dblBarno        As Double
    Dim prm1            As New ADODB.Parameter
    Dim prm2            As New ADODB.Parameter
    Dim prm3            As New ADODB.Parameter
    Dim prm4            As New ADODB.Parameter
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
        SaveTransData_EMEDI = -1
        
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
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
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
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "Update RESULTOFNUM Set                                     " & vbCrLf
                    SQL = SQL & "   TEXTRESULTVAL       = '" & sResult & "'                 " & vbCrLf
                    SQL = SQL & " , RESULTINDATE        = '" & Format(Now, "yyyymmdd") & "' " & vbCrLf
                    SQL = SQL & " , RESULTINTIME        = '" & Format(Now, "hhmm") & "'     " & vbCrLf
                    SQL = SQL & " , RESULTFLAG          = 1                                 " & vbCrLf
                    SQL = SQL & " , PRINTFLAG           = 5                                 " & vbCrLf
                    SQL = SQL & " Where SPCMNO          = '" & strBarcode & "'              " & vbCrLf
                    SQL = SQL & "   And RESULTITEMCODE  = '" & strTestCd & "'               " & vbCrLf
                    SQL = SQL & "   And RESULTFLAG      < 1                                 " & vbCrLf
                    'SQL = SQL & "   And RESULTITEMCODE IN "
                    'SQL = SQL & "      (SELECT RESULTITEMCODE FROM LABRSLTITEMBASES"
                    'SQL = SQL & "        WHERE   CHANNELNO = '" & strTestCd & "')"
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                
                    '-- ���º���
                    SQL = ""
                    SQL = SQL & "Update REGISTINFOS Set                         " & vbCrLf
                    SQL = SQL & "   RESULTSTATE         = 1                     " & vbCrLf
                    SQL = SQL & " , RSVACPTSTATE        = 4                     " & vbCrLf
                    SQL = SQL & " , DIVIDFLAG           = 1                     " & vbCrLf
                    SQL = SQL & " Where SPCMNO          = '" & strBarcode & "'  " & vbCrLf
                    SQL = SQL & "   And DIVIDFLAG       < 3                     " & vbCrLf
                    SQL = SQL & "   And CLAS            = 4                     " & vbCrLf
                    
                    Call SetSQLData("���º���", SQL, "A")
                    AdoCn.Execute SQL
                
                    '-- ���º���
                    SQL = ""
                    SQL = SQL & "Update ORDERINFOS Set                                                  " & vbCrLf
                    SQL = SQL & "   ORDERSTEPSTATE          = 4                                         " & vbCrLf
                    SQL = SQL & " Where (PATID,ORDERDATE,DEPTCODE,SLIPNO,SLIPSEQ) IN                    " & vbCrLf
                    SQL = SQL & "       (Select PATID,ORDERDATE,DEPTCODE,SLIPNO,SLIPSEQ From REGISTINFOS" & vbCrLf
                    SQL = SQL & "         Where SPCMNO          = '" & strBarcode & "'                  " & vbCrLf
                    SQL = SQL & "           And DIVIDFLAG       < 3                                     " & vbCrLf
                    SQL = SQL & "           And CLAS            = 4 )                                   " & vbCrLf
                    
                    Call SetSQLData("���º���", SQL, "A")
                    AdoCn.Execute SQL
                
                    
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_EMEDI = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_EMEDI = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_EMEDI" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
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
        Case "PHILL"
                Call GetWorkList_PHILL(pFrom, pTo, SPD)

        Case "MSINFOTEC"                    'MS������
                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)

        Case "EMEDI"
                Call GetWorkList_EMEDI(pFrom, pTo, SPD)
                
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
                    
'
                    'SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
'                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
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

    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

RST:

End Sub


Public Sub GetWorkList_EMEDI(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intCol      As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       R.RESULTDATE        AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , P.PATID             AS PID      " & vbCrLf
    SQL = SQL & "     , P.PATNAME           AS PNAME    " & vbCrLf
    SQL = SQL & "     , R.SPCMNO            AS BARCODE  " & vbCrLf
    SQL = SQL & "     , R.RESULTITEMCODE    AS ITEM     " & vbCrLf
    SQL = SQL & "  From RESULTOFNUM R, REGISTINFOS L, PATMST P " & vbCrLf
    SQL = SQL & " Where R.RESULTDATE BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    SQL = SQL & "   And R.RESULTITEMCODE IN (" & gAllTestCd & ")            " & vbCrLf
    SQL = SQL & "   And R.ACPTDATE  =   L.ACPTDATE                          " & vbCrLf
    SQL = SQL & "   And R.ACPTSEQ   =   L.ACPTSEQ                           " & vbCrLf
    SQL = SQL & "   And R.PATID     =   P.PATID                             " & vbCrLf
    SQL = SQL & "   And (R.TEXTRESULTVAL = '' OR R.TEXTRESULTVAL IS NULL)   " & vbCrLf
    SQL = SQL & " ORDER BY R.RESULTDATE,P.PATID                             " & vbCrLf

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                .ReDraw = False
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                
                    For intCol = colSTATE + 1 To SPD.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                            SPD.Row = .MaxRows
                            SPD.Col = intCol
                            SPD.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    
                    For intCol = colSTATE + 1 To SPD.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                            SPD.Row = .MaxRows
                            SPD.Col = intCol
                            SPD.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    End If

    RS.Close

    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetWorkList_EMEDI" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

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

    SPD.RowHeight(-1) = 12
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

        Case "MCC"
                Call GetSampleInfo_MCC(asRow, SPD)

        Case "TWIN"
                Call GetSampleInfo_TWIN(asRow, SPD)

        Case "EMEDI"
                Call GetSampleInfo_EMEDI(asRow, SPD)


    End Select


    GetSampleInfo = 1

    Screen.MousePointer = 0


End Function


Public Function GetEquipExamCode_XN1000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    Dim strCBC      As String
    Dim strDIFF     As String
    
    GetEquipExamCode_XN1000 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        '-- ������ ���� ��� CBC/ DIFF �˻��ϵ��� �Ѵ�.
        If strExamCode = "" Then
            strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
            strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
        End If
        
        If strExamCode <> "" Then
            strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
        End If
        
        GetEquipExamCode_XN1000 = strExamCode
        
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
                If strSendCH = "WBC" Or strSendCH = "RBC" Or strSendCH = "HGB" Or _
                    strSendCH = "HCT" Or strSendCH = "MCV" Or strSendCH = "MCH" Or strSendCH = "MCHC" Or _
                    strSendCH = "PLT" Or strSendCH = "RDW-SD" Or strSendCH = "RDW-CV" Or strSendCH = "PDW" Or _
                    strSendCH = "MPV" Or strSendCH = "P-LCR" Or strSendCH = "PCT" Or strSendCH = "NRBC#" Or strSendCH = "NRBC%" Then
                    
                    strCBC = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
                    
                End If
    
                If strSendCH = "NEUT#" Or strSendCH = "LYMPH#" Or strSendCH = "MONO#" Or strSendCH = "EO#" Or strSendCH = "BASO#" Or _
                    strSendCH = "NEUT%" Or strSendCH = "LYMPH%" Or strSendCH = "MONO%" Or strSendCH = "EO%" Or strSendCH = "BASO%" Or _
                    strSendCH = "IG#" Or strSendCH = "IG%" Then
                   
                    '-- ^^^^LYMPH#\�� �ΰ��� ������ ETB �� ��񿡼� �ν����� ���ϱ� ����..(�� �ڸ��� 230)
                    'strDIFF = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                    strDIFF = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                    
                End If
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    strExamCode = strCBC & strDIFF
    
    '-- ������ ���� ��� CBC/ DIFF �˻��ϵ��� �Ѵ�.
    If strExamCode = "" Then
        strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
        strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
    End If
    
    If strExamCode <> "" Then
        GetEquipExamCode_XN1000 = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
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
    
On Error GoTo ErrHandle
    
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

ErrHandle:
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
Function GetSampleInfo_MCC(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_MCC = -1
    
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
    SQL = SQL & "       READING_YMD     AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , BCODE_NO        AS BARCODE      " & vbCrLf
    SQL = SQL & "     , PTNT_NO         AS PID          " & vbCrLf
    SQL = SQL & "     , PTNT_NM         AS PNAME        " & vbCrLf
    SQL = SQL & "     , AGE             AS AGE          " & vbCrLf
    SQL = SQL & "     , SEX             AS SEX          " & vbCrLf
    SQL = SQL & "     , IO_GB           AS INOUT        " & vbCrLf
    SQL = SQL & "     , ORD_CD          AS ITEM         " & vbCrLf
    SQL = SQL & "     , SP_CD           AS SPCCD        " & vbCrLf
    SQL = SQL & "  FROM LIS_INTERFACE1_V                " & vbCrLf
    SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "' " & vbCrLf
    SQL = SQL & "   AND STS_CD = '0'" & vbCrLf                      '0 ����, 1:�������
    SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCrLf
    SQL = SQL & " ORDER BY ORD_CD " & vbCrLf
        
        
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
                SetText SPD, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "�Կ�", "�ܷ�"), asRow, colINOUT
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
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
    
    GetSampleInfo_MCC = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MCC = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_MCC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfo_TWIN(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_TWIN = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    ReDim Preserve gPatTest(0)
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
'    SQL = ""
'    SQL = SQL & "SELECT DISTINCT "
'    SQL = SQL & "       READING_YMD     AS HOSPDATE     " & vbCrLf
'    SQL = SQL & "     , BCODE_NO        AS BARCODE      " & vbCrLf
'    SQL = SQL & "     , PTNT_NO         AS PID          " & vbCrLf
'    SQL = SQL & "     , PTNT_NM         AS PNAME        " & vbCrLf
'    SQL = SQL & "     , AGE             AS AGE          " & vbCrLf
'    SQL = SQL & "     , SEX             AS SEX          " & vbCrLf
'    SQL = SQL & "     , IO_GB           AS INOUT        " & vbCrLf
'    SQL = SQL & "     , ORD_CD          AS ITEM         " & vbCrLf
'    SQL = SQL & "     , SP_CD           AS SPCCD        " & vbCrLf
'    SQL = SQL & "  FROM LIS_INTERFACE1_V                " & vbCrLf
'    SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "' " & vbCrLf
'    SQL = SQL & "   AND STS_CD = '0'" & vbCrLf                      '0 ����, 1:�������
'    SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCrLf
'    SQL = SQL & " ORDER BY ORD_CD " & vbCrLf
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       C.JOBDATE                               AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , C.SPECNO                                AS BARCODE      " & vbCrLf
    SQL = SQL & "     , C.PTNO                                  AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , C.JOBNO                                 AS PID          " & vbCrLf
    SQL = SQL & "     , DECODE(C.GBIO,'I','�Կ�','O','�ܷ�')    AS INOUT        " & vbCrLf
    SQL = SQL & "     , C.SNAME                                 AS PNAME        " & vbCrLf
    SQL = SQL & "     , C.SEX                                   AS SEX          " & vbCrLf
    SQL = SQL & "     , C.AGE                                   AS AGE          " & vbCrLf
    SQL = SQL & "     , A.MASTERCODE                            AS ORDERCODE    " & vbCrLf
    SQL = SQL & "     , A.SUBCODE                               AS ITEM         " & vbCrLf
    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A                             " & vbCrLf
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_MASTER  B                             " & vbCrLf
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_SPECMST C                             " & vbCrLf
    SQL = SQL & " Where A.SPECNO = '" & strBarcode & "'                         " & vbCrLf
    'SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'                     " & vbCrLf '����ڵ�
    'SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")                    " & vbCrLf
    SQL = SQL & "   AND A.SUBCODE IN (" & gAllTestCd & ")                       " & vbCrLf
    SQL = SQL & "   AND C.STATUS  <= '3'                                        " & vbCrLf '�˻����
    SQL = SQL & "   And C.SPECNO  = A.SPECNO                                    " & vbCrLf
    SQL = SQL & "   And A.MASTERCODE = B.MASTERCODE                             " & vbCrLf
    SQL = SQL & " ORDER BY C.JOBDATE, C.SPECNO                                  " & vbCrLf
        
        
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
                'SetText SPD, Format(Trim(RS.Fields("HOSPDATE")) & "", "####-##-##"), asRow, colHOSPDATE
                SetText SPD, Mid(Trim(RS.Fields("HOSPDATE")) & "", 1, 10), asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("INOUT")), asRow, colINOUT
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
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
    
    GetSampleInfo_TWIN = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_TWIN = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_TWIN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- �˻��� ���� ��������
Function GetSampleInfo_EMEDI(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_EMEDI = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    ReDim Preserve gPatTest(0)
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
'    If strRegDate = "" Then
'        Exit Function
'    End If
    
    Screen.MousePointer = 11
        
        'R.RESULTDATE    = '" & strRegDate & "'  " & vbCrLf
    'SQL = SQL & "   And "
    '"
    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       R.RESULTDATE        AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , P.PATID             AS PID      " & vbCrLf
    SQL = SQL & "     , P.PATNAME           AS PNAME    " & vbCrLf
    SQL = SQL & "     , R.SPCMNO            AS BARCODE  " & vbCrLf
    SQL = SQL & "     , R.RESULTITEMCODE    AS ITEM     " & vbCrLf
    SQL = SQL & "  From RESULTOFNUM R, REGISTINFOS L, PATMST P  " & vbCrLf
    SQL = SQL & " Where R.SPCMNO        = '" & strBarcode & "'  " & vbCrLf
    SQL = SQL & "   And R.ACPTDATE      =   L.ACPTDATE          " & vbCrLf
    SQL = SQL & "   And R.ACPTSEQ       =   L.ACPTSEQ           " & vbCrLf
    SQL = SQL & "   And R.PATID         =   P.PATID             " & vbCrLf
    SQL = SQL & "   And R.RESULTITEMCODE IN (" & gAllTestCd & ")" & vbCrLf
    SQL = SQL & "   And (R.TEXTRESULTVAL = '' OR R.TEXTRESULTVAL IS NULL)   " & vbCrLf
    SQL = SQL & " ORDER BY R.RESULTDATE,P.PATID                             " & vbCrLf
        
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
                SetText SPD, Mid(Trim(RS.Fields("HOSPDATE")) & "", 1, 10), asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
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
    
    GetSampleInfo_EMEDI = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_EMEDI = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_EMEDI" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
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
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "'" & vbCrLf
            
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
