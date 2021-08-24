Attribute VB_Name = "modQuery"
Option Explicit

Public SQL          As String
Public RS           As ADODB.Recordset

'-- �˻縶���� ��ȸ
Public Sub GetTestList()
    Dim intRow          As Long
    
    frmMain.spdTest.MaxRows = 0
    intRow = 0
    gAllTestCd = ""
    Erase gArrEQP
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT EQUIPCD,SEQNO,TESTCODE,SENDCHANNEL,RSLTCHANNEL, " & vbCr
    SQL = SQL & " TESTNAME,ABBRNAME,RESPREC,REFLOW,REFHIGH," & vbCr
    SQL = SQL & " RESULTTYPE,CUTOFFUSE,COLIN,COLCOMP,COLOUT," & vbCr
    SQL = SQL & " COHIN,COHCOMP,COHOUT,COMOUT" & vbCr
    SQL = SQL & " ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument, QCReagent, QCUnit, QCTemp " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "
    
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdTest
            .MaxRows = AdoRs_Local.RecordCount
            ReDim Preserve gArrEQP(AdoRs_Local.RecordCount, 15)
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("EQUIPCD").Value & "", intRow, colLMACHCODE)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("SEQNO").Value & "", intRow, colLSEQNO)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("TESTCODE").Value & "", intRow, colLTESTCD)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("SENDCHANNEL").Value & "", intRow, colLOCHANNEL)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("RSLTCHANNEL").Value & "", intRow, colLRCHANNEL)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("TESTNAME").Value & "", intRow, colLTESTNM)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("ABBRNAME").Value & "", intRow, colLABBRNM)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("RESPREC").Value & "", intRow, colLRESSPEC)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("REFLOW").Value & "", intRow, colLLOW)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("REFHIGH").Value & "", intRow, colLHIGH)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("RESULTTYPE").Value & "", intRow, colLRSTTYPE)
                Call SetText(frmMain.spdTest, IIf(AdoRs_Local.Fields("CUTOFFUSE").Value & "" = "Y", "1", "0"), intRow, colLCUTUSE)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COLIN").Value & "", intRow, colLCOLIN)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COLCOMP").Value & "", intRow, colLCOLCOMP)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COLOUT").Value & "", intRow, colLCOLOUT)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COMOUT").Value & "", intRow, colLCOMOUT)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COHIN").Value & "", intRow, colLCOHIN)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COHCOMP").Value & "", intRow, colLCOHCOMP)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COHOUT").Value & "", intRow, colLCOHOUT)
                '--QC
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCLab").Value & "", intRow, colLQCLab)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCLot").Value & "", intRow, colLQCLot)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCAnalyte").Value & "", intRow, colLQCAnalyte)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCMethod").Value & "", intRow, colLQCMethod)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCInstrument").Value & "", intRow, colLQCInstrument)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCReagent").Value & "", intRow, colLQCReagent)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCUnit").Value & "", intRow, colLQCUnit)
                '-- �Ҽ�����ȯ���� ���
                'Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCTemp").Value & "", intRow, colLQCTemp)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCTemp").Value & "", intRow, colLUseResSpec)
               ' Call SetText(frmMain.spdTest, IIf(AdoRs_Local.Fields("QCTemp").Value & "" = "1", "���", "�̻��"), intRow, colLUseResSpec)

                
                
                gArrEQP(intRow, 1) = AdoRs_Local.Fields("SEQNO").Value & ""
                gArrEQP(intRow, 2) = AdoRs_Local.Fields("TESTCODE").Value & ""
                gArrEQP(intRow, 3) = AdoRs_Local.Fields("SENDCHANNEL").Value & ""
                gArrEQP(intRow, 4) = AdoRs_Local.Fields("RSLTCHANNEL").Value & ""
                gArrEQP(intRow, 5) = AdoRs_Local.Fields("ABBRNAME").Value & ""
                gArrEQP(intRow, 6) = AdoRs_Local.Fields("RESPREC").Value & ""
                gArrEQP(intRow, 7) = AdoRs_Local.Fields("REFLOW").Value & ""
                gArrEQP(intRow, 8) = AdoRs_Local.Fields("REFHIGH").Value & ""
                gArrEQP(intRow, 9) = AdoRs_Local.Fields("RESULTTYPE").Value & ""
                gArrEQP(intRow, 10) = AdoRs_Local.Fields("CUTOFFUSE").Value & ""
                gArrEQP(intRow, 11) = AdoRs_Local.Fields("COLCOMP").Value & "" & AdoRs_Local.Fields("COLIN").Value
                gArrEQP(intRow, 12) = AdoRs_Local.Fields("COLOUT").Value & ""
                gArrEQP(intRow, 13) = AdoRs_Local.Fields("COMOUT").Value & ""
                gArrEQP(intRow, 14) = AdoRs_Local.Fields("COHCOMP").Value & "" & AdoRs_Local.Fields("COHIN").Value
                gArrEQP(intRow, 15) = AdoRs_Local.Fields("COHOUT").Value & ""
                
                If Trim(AdoRs_Local.Fields("TESTCODE").Value) <> "" Then
                    If intRow = 1 Or gAllTestCd = "" Then
                        gAllTestCd = "'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                    Else
                        gAllTestCd = gAllTestCd & ",'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                    End If
                End If
                
                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 12
        End With
    End If
End Sub

'-- �˻���������� ��ȸ
Public Sub GetOrderMST()
    Dim intRow          As Long
    
    gAllOrdCd = ""
    intRow = 0
    
    SQL = ""
    SQL = SQL & "SELECT ORDERCODE FROM ORDMASTER " & vbCr
    SQL = SQL & " ORDER BY ORDERCODE "
    
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdOrdMst
            .MaxRows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(frmMain.spdOrdMst, AdoRs_Local.Fields("ORDERCODE").Value & "", intRow, 1)
                
                If Trim(AdoRs_Local.Fields("ORDERCODE").Value) <> "" Then
                    If intRow = 1 Or gAllTestCd = "" Then
                        gAllOrdCd = "'" & AdoRs_Local.Fields("ORDERCODE").Value & "'"
                    Else
                        gAllOrdCd = gAllOrdCd & ",'" & AdoRs_Local.Fields("ORDERCODE").Value & "'"
                    End If
                End If
                
                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 12
        End With
    End If
End Sub

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
    
    'Set AdoRs_Local = New ADODB.Recordset
    
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            GetTestNm = AdoRs_Local.Fields("ITEMNM").Value & ""
            AdoRs_Local.MoveNext
        Loop
    End If
    
    'AdoCn_Local.Close
    
End Function

'-- �˻��ڵ�� �˻�� ��ȸ
Public Function GetTestCd(ByVal pItem As String, Optional pFull As Boolean) As String
    Dim intRow          As Long

    GetTestCd = ""

    SQL = ""
    SQL = SQL & "SELECT b.TESTCODE AS ITEMCD            " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER a, TESTMASTER b       " & vbCrLf
    SQL = SQL & " WHERE a.RSLTCHANNEL = b.RSLTCHANNEL   " & vbCrLf
    SQL = SQL & "   AND a.ABBRNAME = '" & pItem & "'    " & vbCrLf

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            GetTestCd = GetTestCd & "'" & AdoRs_Local.Fields("ITEMCD").Value & "',"
            AdoRs_Local.MoveNext
        Loop
    End If
    
    If GetTestCd <> "" Then
        GetTestCd = Mid(GetTestCd, 1, Len(GetTestCd) - 1)
    End If
    
    AdoRs_Local.Close
    Set AdoRs_Local = Nothing
    
End Function

'-- �˻��׸� ��ȸ
Public Function GetTest(ByVal pTestCd As String) As String
    
On Error GoTo RST
    GetTest = ""
    
    SQL = ""
    SQL = SQL & "Select ORD_NM "
    SQL = SQL & "  From LIS_ORD_LIST_V" & vbCr
    SQL = SQL & " Where ord_cd = '" & pTestCd & "'" & vbCr
  
    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            GetTest = Trim(RS.Fields("ORD_NM")) & ""
            RS.MoveNext
        Loop
    End If

    RS.Close
    
Exit Function

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Function

'-- QC ���� ����
Public Sub SetQCList_Header()
    Dim i   As Integer
    
On Error GoTo RST
    
    SQL = ""
    SQL = SQL & "Delete From QCHEADER "
  
    Call DBExec(AdoCn_Local, SQL)

    With frmQCMaster.spdHeader
        For i = 1 To .MaxRows
            If Trim(GetText(frmQCMaster.spdHeader, i, 1)) = "" Then
                Exit Sub
            End If
            SQL = ""
            SQL = SQL & "Insert Into QCHEADER (LotID,MachID,InstrumentID) " & vbCr
            SQL = SQL & " Values (" & vbCr
            SQL = SQL & "'" & GetText(frmQCMaster.spdHeader, i, 1) & "'," & vbCr
            SQL = SQL & "'" & GetText(frmQCMaster.spdHeader, i, 2) & "'," & vbCr
            SQL = SQL & "'" & GetText(frmQCMaster.spdHeader, i, 3) & "'" & vbCr
            SQL = SQL & ") " & vbCr
        
            Call DBExec(AdoCn_Local, SQL)
        Next
    End With
    
Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_SetQCList_Header" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- QC ���� ����
Public Sub SetQCList_Detail(ByVal pQCID As String)
    Dim i   As Integer
    
On Error GoTo RST
    
    SQL = ""
    SQL = SQL & "Delete From QCDETAIL "
    SQL = SQL & " Where InstrumentID  ='" & pQCID & "'"
  
    Call DBExec(AdoCn_Local, SQL)

    With frmQCMaster.spdQCID
        For i = 1 To .MaxRows
            SQL = ""
            SQL = SQL & "Insert Into QCDETAIL (InstrumentID, QCLevel, ID) " & vbCr
            SQL = SQL & " Values (" & vbCr
            SQL = SQL & "'" & pQCID & "'," & vbCr
            SQL = SQL & "'" & GetText(frmQCMaster.spdQCID, i, 2) & "'," & vbCr
            SQL = SQL & "'" & GetText(frmQCMaster.spdQCID, i, 3) & "'" & vbCr
            SQL = SQL & ") " & vbCr
        
            Call DBExec(AdoCn_Local, SQL)
        Next
    End With
    
Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "SetQCList_Detail" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- QC ���� ��ȸ
Public Sub GetQCList_Header()
    Dim i   As Integer
    
On Error GoTo RST
    
    SQL = ""
    SQL = SQL & "Select LotID,MachID,InstrumentID "
    SQL = SQL & "  From QCHEADER " & vbCr
  
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        i = 1
        frmQCMaster.spdHeader.MaxRows = RS.RecordCount
        Do Until RS.EOF
            Call SetText(frmQCMaster.spdHeader, Trim(RS.Fields("LotID")) & "", i, 1)
            Call SetText(frmQCMaster.spdHeader, Trim(RS.Fields("MachID")) & "", i, 2)
            Call SetText(frmQCMaster.spdHeader, Trim(RS.Fields("InstrumentID")) & "", i, 3)
            i = i + 1
            RS.MoveNext
        Loop
    End If
    
    frmQCMaster.spdHeader.RowHeight(-1) = 14
    RS.Close
    
Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_GetQCList_Header" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- QC ���� ��ȸ -��
Public Sub GetQCList_QCID(ByVal strInst As String)
    Dim i   As Integer
    
On Error GoTo RST
    frmQCMaster.spdQCID.MaxRows = 0
    
    SQL = ""
    SQL = SQL & "Select InstrumentID,QCLevel,ID "
    SQL = SQL & "  From QCDetail " & vbCr
    SQL = SQL & " Where InstrumentID = '" & strInst & "'" & vbCr
    SQL = SQL & " Order By  InstrumentID,QCLevel,ID "
    
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        i = 1
        frmQCMaster.spdQCID.MaxRows = RS.RecordCount
        Do Until RS.EOF
            Call SetText(frmQCMaster.spdQCID, Trim(RS.Fields("InstrumentID")) & "", i, 1)
            Call SetText(frmQCMaster.spdQCID, Trim(RS.Fields("QCLevel")) & "", i, 2)
            Call SetText(frmQCMaster.spdQCID, Trim(RS.Fields("ID")) & "", i, 3)
            i = i + 1
            RS.MoveNext
        Loop
    End If
    
    frmQCMaster.spdQCID.RowHeight(-1) = 14
    RS.Close
    
Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & "GetQCList_QCID" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


'-- QC ���� ��ȸ (���������� QC ���� �𸦶�..)
Public Function strQCFlag(ByVal strInst As String, ByVal strQCBarCd As String) As String
    
On Error GoTo RST
    
    strQCFlag = ""

    SQL = ""
    SQL = SQL & "Select Count(*) AS CNT  "
    SQL = SQL & "  From QCDetail " & vbCr
    SQL = SQL & " Where InstrumentID = '" & strInst & "'" & vbCr
    SQL = SQL & "   And ID = '" & strQCBarCd & "'" & vbCr
    
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        If IsNull(RS.Fields("CNT")) Or RS.Fields("CNT") = 0 Then
            strQCFlag = ""
        Else
            strQCFlag = "QC"
        End If
    End If
    RS.Close
    
Exit Function

RST:
     
                strErrMsg = "��    ġ : " & "strQCFlag" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Function

'-- QC ���� ����Ʈ ��ȸ(����)
Public Sub GetQCList_Detail(ByVal pEqpCD As String, ByVal pLotID As String, ByVal pInstID As String)
    Dim i   As Integer
    
On Error GoTo RST
    frmQCMaster.spdDetail.MaxRows = 0
    
    SQL = ""
    SQL = SQL & "SELECT Distinct b.AnalyteID, a.lablottestid,c.name,b.MethodID,b.ReagentID, b.UnitID, b.TemperatureID " & vbCr
    SQL = SQL & "  FROM LabLotTest a, test b, analyte c" & vbCr
    SQL = SQL & " WHERE a.Labid = '" & pEqpCD & "'" & vbCr
    SQL = SQL & "   AND a.Lotid = '" & pLotID & "'" & vbCr
    SQL = SQL & "   AND b.InstrumentID = '" & pInstID & "'" & vbCr
    SQL = SQL & "   AND a.testid = b.testid " & vbCr
    SQL = SQL & "   AND b.AnalyteID = c.AnalyteID " & vbCr
    SQL = SQL & " ORDER BY a.lablottestid"
    
    '-- Record Count ������
    AdoCn_QC.CursorLocation = adUseClient
    Set RS = AdoCn_QC.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        i = 1
        frmQCMaster.spdDetail.MaxRows = RS.RecordCount
        Do Until RS.EOF
            Call SetText(frmQCMaster.spdDetail, Trim(RS.Fields("AnalyteID")) & "", i, 1)
            Call SetText(frmQCMaster.spdDetail, Trim(RS.Fields("lablottestid")) & "", i, 2)
            Call SetText(frmQCMaster.spdDetail, Trim(RS.Fields("name")) & "", i, 3)
            Call SetText(frmQCMaster.spdDetail, pInstID, i, 4)
            Call SetText(frmQCMaster.spdDetail, Trim(RS.Fields("MethodID")) & "", i, 5)
            Call SetText(frmQCMaster.spdDetail, Trim(RS.Fields("ReagentID")) & "", i, 6)
            Call SetText(frmQCMaster.spdDetail, Trim(RS.Fields("UnitID")) & "", i, 7)
            Call SetText(frmQCMaster.spdDetail, Trim(RS.Fields("TemperatureID")) & "", i, 8)
           ' Call SetText(frmQCMaster.spdDetail, Trim(RS.Fields("name")) & "", i, 8)
            i = i + 1
            RS.MoveNext
        Loop
    End If
    
    frmQCMaster.spdDetail.RowHeight(-1) = 14
    RS.Close
    
Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & "GetQCList_Detail" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- QC ��� ����Ʈ ��ȸ
Public Function GetQCResult_Detail(ByVal pEqpCD As String, ByVal pQCChannel As String, ByVal pAnalyID As String, ByVal pResult As String, Optional ByVal pMethod As String = "") As String
    Dim i               As Integer
    Dim strLotID        As String
    Dim strMachID       As String
    Dim strQCLevel      As String
    Dim strQCTmp        As String
    Dim strQCTemp()     As Variant
    Dim strQCVal        As String
    Dim strDtTM         As String
    Dim strRun          As String
    Dim strQCBuf        As String
    Dim varQCBuf        As Variant
    Dim FindFile        As String
    Dim intCnt          As Integer
    Dim strMethodID     As String
    Dim strReagentID    As String
    Dim strUnitID       As String
    Dim strTemperatureID    As String
    Dim strClip         As String
    
'Point|201708311040|1|1|506927|28550|294|98|1906|6|21|6|sa|||3.26|
'Point|201708311040|2|1|506927|28550|294|238|1906|6|43|6|sa|||2.75|
'Point|201708311040|1|1|506927|28550|295|98|1906|6|20|6|sa|||2.17|
'Point|201708311040|1|1|506927|28550|296|975|1906|6|15|6|sa|||5.7|
'Point|201708311040|1|1|506927|28550|297|103|1906|6|93|6|sa|||14.9|
'Point|201708311040|1|1|506927|28550|307|98|1906|6|23|6|sa|||68.6|
'Point|201708311040|1|1|506927|28550|351|103|1906|6|24|6|sa|||26.3|
'Point|201708311040|1|1|506927|28550|352|103|1906|6|15|6|sa|||38.4|
'Point|201708311040|1|1|506927|28550|1340|98|1906|6|15|6|sa|||36.8|
'Point|201708311040|1|1|506927|28550|355|103|1906|6|93|6|sa|||16.1|
'Point|201708311040|1|1|506927|28550|356|98|1906|6|21|6|sa|||56|
'Point|201708311040|1|1|506927|28550|363|98|1906|6|23|6|sa|||9.9|
'Point|201708311040|1|1|506927|28550|398|103|1906|6|93|6|sa|||43.0|
'Point|201708311040|1|1|506927|28550|385|103|1906|6|93|6|sa|||36.2|
'Point|201708311040|1|1|506927|28550|387|103|1906|6|93|6|sa|||13.0|
'Point|201708311040|1|1|506927|28550|563|103|1906|6|93|6|sa|||1.5|
'Point|201708311040|1|1|506927|28550|558|103|1906|6|93|6|sa|||0.4|
'Point|201708311040|2|1|506927|28550|294|238|1906|6|43|6|sa|||2.75|

'On Error GoTo RST

    GetQCResult_Detail = ""
    strQCVal = ""
    strDtTM = Format(Now, "yyyymmddhhmm")
    strRun = "1"
    
    If InStr(pAnalyID, ",") > 0 Then
        pAnalyID = Trim(mGetP(pAnalyID, 2, ","))
    End If
    
    SQL = ""
    SQL = SQL & "SELECT Distinct h.LotID, h.MachID,d.QCLevel " & vbCr
    SQL = SQL & "  FROM QCHEADER h,QCDETAIL d" & vbCr
    SQL = SQL & " WHERE d.ID = '" & pQCChannel & "'" & vbCr
    SQL = SQL & "   AND h.InstrumentID = d.InstrumentID "
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strLotID = Trim(RS.Fields("LotID")) & ""
        strMachID = Trim(RS.Fields("MachID")) & ""
        strQCLevel = Trim(RS.Fields("QCLevel")) & ""
        'RS.MoveNext
    End If
    
    RS.Close
    
    intCnt = 0
    
    SQL = ""
    SQL = SQL & "SELECT Distinct b.AnalyteID, a.lablottestid,c.name,b.MethodID,b.ReagentID, b.UnitID, b.TemperatureID " & vbCr
    SQL = SQL & "  FROM LabLotTest a, test b, analyte c" & vbCr
    SQL = SQL & " WHERE a.Labid = '" & pEqpCD & "'" & vbCr
    SQL = SQL & "   AND a.Lotid = '" & strLotID & "'" & vbCr
    SQL = SQL & "   AND b.InstrumentID = '" & strMachID & "'" & vbCr
    SQL = SQL & "   AND b.AnalyteID = '" & pAnalyID & "'" & vbCr
    SQL = SQL & "   AND a.testid = b.testid " & vbCr
    SQL = SQL & "   AND b.AnalyteID = c.AnalyteID " & vbCr
    If pMethod <> "" Then
        SQL = SQL & "   AND b.MethodID  ='" & pMethod & "'" & vbCr
    End If
    SQL = SQL & " ORDER BY a.lablottestid"
    
   ' Erase strQCTemp
    
    '-- Record Count ������
    AdoCn_QC.CursorLocation = adUseClient
    Set RS = AdoCn_QC.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strMethodID = Trim(RS.Fields("MethodID"))
        strMachID = strMachID
        strReagentID = Trim(RS.Fields("ReagentID"))
        strUnitID = Trim(RS.Fields("UnitID"))
        strTemperatureID = Trim(RS.Fields("TemperatureID"))
    End If
    
    RS.Close
    
    strQCVal = ""
    strRun = 1
            

    With frmMain.spdQcResult
        For i = 1 To .DataRowCnt
            If GetText(frmMain.spdQcResult, i, 7) = pAnalyID Then
                strRun = strRun + 1
                Exit For
            End If
        Next
        .MaxRows = .MaxRows + 1
        
        Call SetText(frmMain.spdQcResult, "Point", .MaxRows, 1)
        Call SetText(frmMain.spdQcResult, strDtTM, .MaxRows, 2)
        Call SetText(frmMain.spdQcResult, strRun, .MaxRows, 3)
        Call SetText(frmMain.spdQcResult, strQCLevel, .MaxRows, 4)
        Call SetText(frmMain.spdQcResult, pEqpCD, .MaxRows, 5)
        Call SetText(frmMain.spdQcResult, strLotID, .MaxRows, 6)
        Call SetText(frmMain.spdQcResult, pAnalyID, .MaxRows, 7)
        Call SetText(frmMain.spdQcResult, strMethodID, .MaxRows, 8)
        Call SetText(frmMain.spdQcResult, strMachID, .MaxRows, 9)
        Call SetText(frmMain.spdQcResult, strReagentID, .MaxRows, 10)
        Call SetText(frmMain.spdQcResult, strUnitID, .MaxRows, 11)
        Call SetText(frmMain.spdQcResult, strTemperatureID, .MaxRows, 12)
        Call SetText(frmMain.spdQcResult, "sa", .MaxRows, 13)
        Call SetText(frmMain.spdQcResult, "", .MaxRows, 14)
        Call SetText(frmMain.spdQcResult, "", .MaxRows, 15)
        Call SetText(frmMain.spdQcResult, pResult, .MaxRows, 16)
    End With

    'GetQCResult_Detail = strQCVal
    
Exit Function

RST:
     
                strErrMsg = "��    ġ : " & "GetQCResult_Detail" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Function


'-- QC ��� ����Ʈ ��ȸ
Public Function GetQCResult_Detail_Type2(ByVal pEqpCD As String, ByVal pQCChannel As String, ByVal pAnalyID As String, ByVal pResult As String) As String
    Dim i               As Integer
    Dim strLotID        As String
    Dim strMachID       As String
    Dim strQCLevel      As String
    Dim strQCTmp        As String
    Dim strQCTemp()     As Variant
    Dim strQCVal        As String
    Dim strDtTM         As String
    Dim strRun          As String
    Dim strQCBuf        As String
    Dim varQCBuf        As Variant
    Dim FindFile        As String
    Dim intCnt          As Integer
    
'Point|201708311040|1|1|506927|28550|294|98|1906|6|21|6|sa|||3.26|
'Point|201708311040|2|1|506927|28550|294|238|1906|6|43|6|sa|||2.75|
'Point|201708311040|1|1|506927|28550|295|98|1906|6|20|6|sa|||2.17|
'Point|201708311040|1|1|506927|28550|296|975|1906|6|15|6|sa|||5.7|
'Point|201708311040|1|1|506927|28550|297|103|1906|6|93|6|sa|||14.9|
'Point|201708311040|1|1|506927|28550|307|98|1906|6|23|6|sa|||68.6|
'Point|201708311040|1|1|506927|28550|351|103|1906|6|24|6|sa|||26.3|
'Point|201708311040|1|1|506927|28550|352|103|1906|6|15|6|sa|||38.4|
'Point|201708311040|1|1|506927|28550|1340|98|1906|6|15|6|sa|||36.8|
'Point|201708311040|1|1|506927|28550|355|103|1906|6|93|6|sa|||16.1|
'Point|201708311040|1|1|506927|28550|356|98|1906|6|21|6|sa|||56|
'Point|201708311040|1|1|506927|28550|363|98|1906|6|23|6|sa|||9.9|
'Point|201708311040|1|1|506927|28550|398|103|1906|6|93|6|sa|||43.0|
'Point|201708311040|1|1|506927|28550|385|103|1906|6|93|6|sa|||36.2|
'Point|201708311040|1|1|506927|28550|387|103|1906|6|93|6|sa|||13.0|
'Point|201708311040|1|1|506927|28550|563|103|1906|6|93|6|sa|||1.5|
'Point|201708311040|1|1|506927|28550|558|103|1906|6|93|6|sa|||0.4|
'Point|201708311040|2|1|506927|28550|294|238|1906|6|43|6|sa|||2.75|

On Error GoTo RST

    GetQCResult_Detail_Type2 = ""
    strQCVal = ""
    strDtTM = Format(Now, "yyyymmddhhmm")
    strRun = "1"
    
    If InStr(pAnalyID, ",") > 0 Then
        pAnalyID = Trim(mGetP(pAnalyID, 2, ","))
    End If
    
    SQL = ""
    SQL = SQL & "SELECT Distinct h.LotID, h.MachID,d.QCLevel " & vbCr
    SQL = SQL & "  FROM QCHEADER h,QCDETAIL d" & vbCr
    SQL = SQL & " WHERE d.ID = '" & pQCChannel & "'" & vbCr
    SQL = SQL & "   AND h.InstrumentID = d.InstrumentID "
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            strLotID = Trim(RS.Fields("LotID")) & ""
            strMachID = Trim(RS.Fields("MachID")) & ""
            strQCLevel = Trim(RS.Fields("QCLevel")) & ""
            RS.MoveNext
        Loop
    End If
    
    RS.Close
    
    intCnt = 0
    
    SQL = ""
    SQL = SQL & "SELECT Distinct b.AnalyteID, a.lablottestid,c.name,b.MethodID,b.ReagentID, b.UnitID, b.TemperatureID " & vbCr
    SQL = SQL & "  FROM LabLotTest a, test b, analyte c" & vbCr
    SQL = SQL & " WHERE a.Labid = '" & pEqpCD & "'" & vbCr
    SQL = SQL & "   AND a.Lotid = '" & strLotID & "'" & vbCr
    SQL = SQL & "   AND b.InstrumentID = '" & strMachID & "'" & vbCr
    SQL = SQL & "   AND b.AnalyteID = '" & pAnalyID & "'" & vbCr
    SQL = SQL & "   AND a.testid = b.testid " & vbCr
    SQL = SQL & "   AND b.AnalyteID = c.AnalyteID " & vbCr
    SQL = SQL & " ORDER BY a.lablottestid"
    
    Erase strQCTemp
    
    '-- Record Count ������
    AdoCn_QC.CursorLocation = adUseClient
    Set RS = AdoCn_QC.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            ReDim Preserve strQCTemp(intCnt)
            strQCTemp(intCnt) = Trim(RS.Fields("MethodID")) & "|" & strMachID & "|" & Trim(RS.Fields("ReagentID")) & "|" & Trim(RS.Fields("UnitID")) & "|" & Trim(RS.Fields("TemperatureID")) & "|"
            'strQCTmp = Trim(RS.Fields("MethodID")) & "|" & strMachID & "|" & Trim(RS.Fields("ReagentID")) & "|" & Trim(RS.Fields("UnitID")) & "|" & Trim(RS.Fields("TemperatureID")) & "|"
            intCnt = intCnt + 1
            RS.MoveNext
        Loop
    End If
    
    RS.Close
    
    strQCVal = ""
    If intCnt > 0 Then
        For i = 0 To UBound(strQCTemp)
            strRun = i + 1
            If strQCTemp(i) <> "" Then
                strQCVal = strQCVal & "Point" & "|"
                strQCVal = strQCVal & strDtTM & "|"         ' Date Time     // yyyymmddhhmm
                strQCVal = strQCVal & strRun & "|"             ' run           // 1,2,3,4
                strQCVal = strQCVal & strQCLevel & "|"      ' level         // 1,2,3
                strQCVal = strQCVal & pEqpCD & "|"          ' lab           // 447834(�����ڵ�� ��ü ����?)
                strQCVal = strQCVal & strLotID & "|"        ' lot           // 159792(�Է�)
                strQCVal = strQCVal & pAnalyID & "|"        ' analyte       // �˻��׸񸶴� ����,  Cyfra 21-1 : pAnalyte = "222"
                strQCVal = strQCVal & strQCTemp(i)              ' MethodID, InstrumentID, ReagentID, UnitID, TemperatureID
                strQCVal = strQCVal & "sa" & "|"
                strQCVal = strQCVal & "" & "|"
                strQCVal = strQCVal & "" & "|"
                strQCVal = strQCVal & pResult & "|"
                strQCVal = strQCVal & vbCrLf
            End If
        Next
    End If
        
    GetQCResult_Detail_Type2 = strQCVal
    
Exit Function

RST:
     
                strErrMsg = "��    ġ : " & "GetQCResult_Detail" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Function

Public Function f_subSet_RefVal(ByVal strORCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAge As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
    Dim rs_svr As ADODB.Recordset

On Error GoTo ErrorTrap
    
    strRSLT = Replace(strRSLT, "<", "")
    strRSLT = Replace(strRSLT, ">", "")
    f_subSet_RefVal = " "
    
    f_subSet_RefVal = ""
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
    SQL = SQL & " Where ORCD =  '" & strORCD & "'"
    
    Set rs_svr = AdoCn.Execute(SQL)
    Do Until rs_svr.EOF
        If IsNumeric(strRSLT) And IsNumeric(rs_svr.Fields("MAX")) And IsNumeric(rs_svr.Fields("MIN")) Then
            If Val(strRSLT) > Val(rs_svr.Fields("MAX")) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRSLT) < Val(rs_svr.Fields("MIN")) Then
                f_subSet_RefVal = "L"
            Else
                f_subSet_RefVal = " "
            End If
        Else
            f_subSet_RefVal = " "
        End If
        rs_svr.MoveNext
    
    Loop
    
Exit Function

ErrorTrap:
     
End Function

'-- ��ũ����Ʈ ��ȸ
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim getBarcode  As String
    Dim strBarcode  As String
    Dim blnSame     As Boolean
    Dim strItems    As String
    Dim intOCnt     As Integer
    Dim strDateFr8  As String
    Dim strDateTo8  As String
    Dim strDateFr10 As String
    Dim strDateTo10 As String
    
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    Select Case gOCS
        Case "KYU"
'            intBcNow = DateDiff("d", "1999-01-01", Format(Now, "yyyy-mm-dd"))
'            intBcFive = Mid(strBarcode, 1, 5) '06351
'            intBcAdd = intBcFive - intBcNow
'            strADT = Format(Now + intBcAdd, "yyyymmdd")
'            strSlip1 = Mid(strBarcode, 6, 2)  '23
'            strSlip2 = Mid(strBarcode, 8, 5)  '00001
    
    
        '      -- ���̺� ���
'            SQL = ""
'            SQL = SQL & "SELECT DISTINCT To_Char(R.jeobsudt, 'yyyymmdd') as HOSPDATE, R.slipno1 as RACK, R.slipno2 as POS, R.ptno as PID, p.sname as PNAME, R.itemcd as ITEM " & vbCr
'            SQL = SQL & "  FROM twexam_general_sub R, twexam_general O, twbas_patient p" & vbCr
'            SQL = SQL & " WHERE r.verify <> 'Y' " & vbCr
'            SQL = SQL & "   AND O.gbch = 'Y' " & vbCr
'            SQL = SQL & "   AND R.jeobsudt = to_date('" & strADT & "','yyyymmdd')" & vbCr
'            SQL = SQL & "   AND R.slipno1 = '" & gHOSP.PARTCD & "'" & vbCr
'            'SQL = SQL & "   AND R.slipno2 = '" & strSlip2 & "'" & vbCr
'            SQL = SQL & "   AND R.itemcd IN (" & gAllTestCd & ")" & vbCr
'            SQL = SQL & "   AND R.jeobsudt = O.jeobsudt" & vbCr
'            SQL = SQL & "   AND R.slipno1 = O.slipno1" & vbCr
'            SQL = SQL & "   AND R.slipno2 = O.slipno2" & vbCr
'            SQL = SQL & "   AND R.PTNO = O.PTNO" & vbCr
'            SQL = SQL & "   AND R.PTNO = p.PTNO"
'            SQL = SQL & " ORDER BY HOSPDATE, POS "
'''
'''      SQL.Text:=' Select distinct                              '+#13#10+
'''                '       To_Char(R.jeobsudt, ''yyyymmdd'') ADT  '+#13#10+  //��Ʈ��������
'''                '       , R.slipno1                            '+#13#10+  //�۾���ȣ
'''                '       , R.slipno2                            '+#13#10+
'''                '       , R.ptno                               '+#13#10+
'''                '       , O.deptcode                           '+#13#10+
'''                '       , O.status                             '+#13#10+
'''                '       , p.sname                              '+#13#10+
'''                '       , p.jumin1||p.jumin2 as jno            '+#13#10+
'''                '  From twexam_general_sub R                   '+#13#10+
'''                '     , twexam_general O                       '+#13#10+
'''                '     , twbas_patient p                        '+#13#10+
'''                '  where r.verify <> ''Y''                     '+#13#10+
'''                '    and O.gbch = ''Y''                        '+#13#10+
'''                '    and R.slipno1 = ''13''                    '+#13#10+
'''                '    and R.itemcd in '+EList                    +#13#10+
'''                '    and R.jeobsudt Between to_date('''+WF+'000000'', ''yyyymmddHH24miss'')  '+#13#10+
'''                '                       and to_date('''+WT+'235959'', ''yyyymmddHH24miss'')  '+#13#10+
'''                '    and R.jeobsudt = O.jeobsudt               '+#13#10+
'''                '    and R.slipno1  = O.slipno1                '+#13#10+
'''                '    and R.slipno2  = O.slipno2                '+#13#10+
'''                '    and R.PTNO     = O.PTNO                   '+#13#10+
'''                '    and R.PTNO     = p.PTNO                   '+#13#10+
'''                '  order by adt, R.slipno2 ';
                
            Call SetSQLData("��ũ��ȸ", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                frmWorkList.spdWork.MaxRows = 0
                strItems = ""
                Do Until RS.EOF
                    iCnt = iCnt + 1
                    With frmWorkList.spdWork
                        .ReDraw = False
                        
                        For i = 1 To frmWorkList.spdWork.DataRowCnt
                            strDate = GetText(frmWorkList.spdWork, i, colHOSPDATE)
                            strBarcode = GetText(frmWorkList.spdWork, i, colBARCODE)
                            If Trim(RS("HOSPDATE")) = strDate And Trim(RS("BARCODE")) = strBarcode Then
                                blnSame = True
                            End If
                        Next
                        
                        If blnSame = False Then
                            .MaxRows = .MaxRows + 1
                            intRow = .MaxRows
                                
                            SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                            SetText frmWorkList.spdWork, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                            SetText frmWorkList.spdWork, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                            'SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
                            SetText frmWorkList.spdWork, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
'                            SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                            SetText frmWorkList.spdWork, Trim(RS.Fields("INOUT")) & "", intRow, colINOUT
                            SetText frmWorkList.spdWork, frmWorkList.txtSeq.Text, intRow, colSEQNO
                            SetText frmWorkList.spdWork, GetSampleITEM(intRow), intRow, colITEMS
                            
                            frmWorkList.txtSeq.Text = frmWorkList.txtSeq.Text + 1
                        
                        End If
                    End With
                    
                    blnSame = False
                
                    DoEvents
                    
                    RS.MoveNext
                Loop
                frmWorkList.chkAll.Value = "1"
            Else
                frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
                    
                    
        Case "BIGUBCARE"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT i.IntOdrDte AS HOSPDATE"
            'SQL = SQL & ", i.IntLabNum AS PID "               ' �˻��ȣ"
            SQL = SQL & ", i.IntLabNum AS BARCODE "           ' �˻��ȣ"
            SQL = SQL & ", i.IntChtNum AS CHARTNO "           ' ��Ʈ��ȣ"
            SQL = SQL & ", i.IntPatNam AS PNAME "             ' ȯ�ڸ�"
'            SQL = SQL & ", i.IntSexTyp AS SEX"                ' ����"
            'SQL = SQL & ", i.IntOdrStt AS STATE               ' ����"
            SQL = SQL & ", i.IntEmgYon AS INOUT"              ' ���޿���"
            SQL = SQL & "  FROM interfacedb..IntRst i, aphdb..rstinf r " & vbCr
            SQL = SQL & " WHERE r.RstOdrStt not in ('OC') " & vbCr
            SQL = SQL & "   AND (r.rstrstval = '' or rstrstval is null)" & vbCr
            If gHOSP.MACHNM <> "HITACHI7080" Then
                SQL = SQL & "   AND i.intodrtyp = '" & gHOSP.PARTCD & "'" & vbCr  ''HEMO'
            End If
            SQL = SQL & "   AND i.IntOdrDte BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
            SQL = SQL & "   AND i.IntLabCod + cast(IntLabseq as varchar(3)) IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND i.intlabnum = r.rstlabnum" & vbCr
            SQL = SQL & "   AND i.intodrdte = r.rstodrdte" & vbCr
            SQL = SQL & "   AND i.intlabseq = r.rstlabseq" & vbCr
            SQL = SQL & "   AND i.intlabcod = r.rstodrcod" & vbCr
     
            Call SetSQLData("��ũ��ȸ", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                frmWorkList.spdWork.MaxRows = 0
                strItems = ""
                Do Until RS.EOF
                    iCnt = iCnt + 1
                    With frmWorkList.spdWork
                        .ReDraw = False
                        
                        For i = 1 To frmWorkList.spdWork.DataRowCnt
                            strDate = GetText(frmWorkList.spdWork, i, colHOSPDATE)
                            strBarcode = GetText(frmWorkList.spdWork, i, colBARCODE)
                            If Trim(RS("HOSPDATE")) = strDate And Trim(RS("BARCODE")) = strBarcode Then
                                blnSame = True
                            End If
                        Next
                        
                        If blnSame = False Then
                            .MaxRows = .MaxRows + 1
                            intRow = .MaxRows
                                
                            SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                            SetText frmWorkList.spdWork, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                            SetText frmWorkList.spdWork, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                            'SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
                            SetText frmWorkList.spdWork, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
'                            SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                            SetText frmWorkList.spdWork, Trim(RS.Fields("INOUT")) & "", intRow, colINOUT
                            SetText frmWorkList.spdWork, frmWorkList.txtSeq.Text, intRow, colSEQNO
                            SetText frmWorkList.spdWork, GetSampleITEM(intRow), intRow, colITEMS
                            
                            frmWorkList.txtSeq.Text = frmWorkList.txtSeq.Text + 1
                        
                        End If
                    End With
                    
                    blnSame = False
                
                    DoEvents
                    
                    RS.MoveNext
                Loop
                frmWorkList.chkAll.Value = "1"
            Else
                frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
        Case "BIT"
            SQL = ""
            SQL = SQL & " SELECT DISTINCT SUBSTRING(O.OCMACPDTM,1,8) AS HOSPDATE" & vbCr
            SQL = SQL & "        ,R.RESOCMNUM AS BARCODE "
            SQL = SQL & "        ,O.OCMCHTNUM AS CHARTNO "
            SQL = SQL & "        ,R.RESOCMNUM AS PID " & vbCr
            SQL = SQL & "        ,P.PBSPATNAM AS PNAME"
            SQL = SQL & "        ,P.PBSSEXTYP AS SEX"
            SQL = SQL & "        ,R.ResLabCod AS ITEM " & vbCr
            SQL = SQL & "        ,R.ResOdrSeq, R.ResSeq, R.ResSubSeq " & vbCr
            SQL = SQL & "   FROM RESINF AS R, OCMINF AS O, PBSINF AS P, LABMST AS E" & vbCr
            SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & pFrom & "000000" & "' AND '" & pTo & "235959" & "'" & vbCr
            SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')" & vbCr
            SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM" & vbCr
            SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM" & vbCr
            SQL = SQL & "   AND R.RESLABCOD = E.LABCOD" & vbCr
            '-- ���������
            If frmWorkList.chkSave.Value = "0" Then
                SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCr         '--  'I':�߰� 'F' �Ϸ�"
                SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)" & vbCr
            End If
            SQL = SQL & " ORDER BY HOSPDATE, CHARTNO, PID "
    
            Call SetSQLData("��ũ��ȸ", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                frmWorkList.spdWork.MaxRows = 0
                strItems = ""
                Do Until RS.EOF
                    iCnt = iCnt + 1
                    With frmWorkList.spdWork
                        .ReDraw = False
                        
                        For i = 1 To frmWorkList.spdWork.DataRowCnt
                            strDate = GetText(frmWorkList.spdWork, i, colHOSPDATE)
                            strBarcode = GetText(frmWorkList.spdWork, i, colBARCODE)
                            If Trim(RS("HOSPDATE")) = strDate And Trim(RS("BARCODE")) = strBarcode Then
                                blnSame = True
                            End If
                        Next
                        
                        If blnSame = False Then
                            .MaxRows = .MaxRows + 1
                            intRow = .MaxRows
                                
                            SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                            SetText frmWorkList.spdWork, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                            SetText frmWorkList.spdWork, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
                            SetText frmWorkList.spdWork, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                            SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                            SetText frmWorkList.spdWork, frmWorkList.txtSeq.Text, intRow, colSEQNO
                            SetText frmWorkList.spdWork, GetSampleITEM(intRow), intRow, colITEMS
                            
                            frmWorkList.txtSeq.Text = frmWorkList.txtSeq.Text + 1
                        
                        End If
                    End With
                    
                    blnSame = False
                
                    DoEvents
                    
                    RS.MoveNext
                Loop
                frmWorkList.chkAll.Value = "1"
            Else
                frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
        
        Case "NAVY"
            SQL = ""
            SQL = SQL & "Select DISTINCT ORDDATE as HOSPDATE"
            SQL = SQL & ", SPCID as BARCODE "
            SQL = SQL & ", PATNO as PID "
            SQL = SQL & ", WORKCODE as CHARTNO "
            'SQL = SQL & ", ORDSEQNO as SEQ " & vbCr
            'SQL = SQL & ", EXAMCODE as ITEM " & vbCr
            SQL = SQL & "  From SLXWORKT " & vbCr
            SQL = SQL & " Where ORDDATE BETWEEN  '" & pFrom & "' And '" & pTo & "'" & vbCr
            SQL = SQL & "   And HOSPID   = '" & gHOSP.HOSPCD & "'" & vbCr         ' �δ��ڵ�
            SQL = SQL & "   And ROOMCODE = '" & gHOSP.PARTCD & "'" & vbCr         ' �Һ�
            SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")" & vbCr        ' �˻��ڵ�
            SQL = SQL & "   And (RSLTTEXT = '' or RSLTTEXT is null) " & vbCr    ' ���Ȯ������
            'SQL = SQL & "   And PROCSTAT = 'N' " & vbCr   '-- ���Ȯ������
            SQL = SQL & " ORDER BY ORDDATE, SPCID, PATNO, WORKCODE, ORDSEQNO "
            
            Call SetSQLData("��ũ��ȸ", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                SPD.MaxRows = 0
                strItems = ""
                Do Until RS.EOF
                    iCnt = iCnt + 1
                    With SPD
                        .ReDraw = False
                        
                        For i = 1 To SPD.DataRowCnt
                            strDate = GetText(SPD, i, colHOSPDATE)
                            strBarcode = GetText(SPD, i, colBARCODE)
                            If Trim(RS("HOSPDATE")) = strDate And Trim(RS("BARCODE")) = strBarcode Then
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
                            'SetText SPD, Trim(RS.Fields("SEQ")) & "", intRow, colSEQNO
                            SetText SPD, GetSampleITEM(intRow), intRow, colITEMS
                            
'                            SetText SPD, "1", intRow, colCHECKBOX
'                            SetText SPD, "20180417", intRow, colHOSPDATE
'                            SetText SPD, "1234567890", intRow, colBARCODE
'                            SetText SPD, "9876543210", intRow, colCHARTNO
'                            SetText SPD, "4561230789", intRow, colPID
                            'SetText SPD, "126", intRow, colSEQNO
                            
                            frmWorkList.txtSeq.Text = frmWorkList.txtSeq.Text + 1
                        
                        End If
                    End With
                    
                    blnSame = False
                
                    DoEvents
                    
                    RS.MoveNext
                Loop
                frmWorkList.chkAll.Value = "1"
            Else
                frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
    End Select

     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- ��ũ����Ʈ ��ȸ
Public Sub GetWorkList_Main(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim getBarcode  As String
    Dim strBarcode  As String
    Dim blnSame     As Boolean
    Dim strItems    As String
    Dim intOCnt     As Integer
    Dim strDateFr8  As String
    Dim strDateTo8  As String
    Dim strDateFr10 As String
    Dim strDateTo10 As String
    
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    Select Case gOCS
        Case "BIT"
            SQL = ""
            SQL = SQL & " SELECT DISTINCT SUBSTRING(O.OCMACPDTM,1,8) AS HOSPDATE" & vbCr
            SQL = SQL & "        ,R.RESOCMNUM AS BARCODE "
            SQL = SQL & "        ,O.OCMCHTNUM AS CHARTNO "
            SQL = SQL & "        ,R.RESOCMNUM AS PID " & vbCr
            SQL = SQL & "        ,P.PBSPATNAM AS PNAME"
            SQL = SQL & "        ,P.PBSSEXTYP AS SEX"
            SQL = SQL & "        ,R.ResLabCod AS ITEM " & vbCr
            SQL = SQL & "        ,R.ResOdrSeq, R.ResSeq, R.ResSubSeq " & vbCr
            SQL = SQL & "   FROM RESINF AS R, OCMINF AS O, PBSINF AS P, LABMST AS E" & vbCr
            SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & pFrom & "000000" & "' AND '" & pTo & "235959" & "'" & vbCr
            SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')" & vbCr
            SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM" & vbCr
            SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM" & vbCr
            SQL = SQL & "   AND R.RESLABCOD = E.LABCOD" & vbCr
            '-- ���������
            If frmWorkList.chkSave.Value = "0" Then
                SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCr         '--  'I':�߰� 'F' �Ϸ�"
                SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)" & vbCr
            End If
            SQL = SQL & " ORDER BY HOSPDATE, CHARTNO, PID "
    
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                SPD.MaxRows = 0
                strItems = ""
                Do Until RS.EOF
                    iCnt = iCnt + 1
                    With SPD
                        .ReDraw = False
                        
                        For i = 1 To SPD.DataRowCnt
                            strDate = GetText(SPD, i, colHOSPDATE)
                            strBarcode = GetText(SPD, i, colBARCODE)
                            If Trim(RS("HOSPDATE")) = strDate And Trim(RS("BARCODE")) = strBarcode Then
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
                            SetText SPD, frmWorkList.txtSeq.Text, intRow, colSEQNO
                            
                        End If
                    End With
                    
                    blnSame = False
                
                    DoEvents
                    
                    RS.MoveNext
                Loop
            Else
                frmMain.lblCommStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
            End If
            
            RS.Close
            
            
        Case "NAVY"
            SQL = ""
            SQL = SQL & "Select DISTINCT ORDDATE as HOSPDATE"
            SQL = SQL & ", SPCID as BARCODE "
            SQL = SQL & ", PATNO as PID "
            SQL = SQL & ", WORKCODE as CHARTNO "
            'SQL = SQL & ", ORDSEQNO as SEQ " & vbCr
            'SQL = SQL & ", EXAMCODE as ITEM " & vbCr
            SQL = SQL & "  From SLXWORKT " & vbCr
            SQL = SQL & " Where ORDDATE BETWEEN  '" & pFrom & "' And '" & pTo & "'" & vbCr
            SQL = SQL & "   And HOSPID   = '" & gHOSP.HOSPCD & "'" & vbCr         ' �δ��ڵ�
            SQL = SQL & "   And ROOMCODE = '" & gHOSP.PARTCD & "'" & vbCr         ' �Һ�
            SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")" & vbCr        ' �˻��ڵ�
            SQL = SQL & "   And (RSLTTEXT = '' or RSLTTEXT is null) " & vbCr    ' ���Ȯ������
            'SQL = SQL & "   And PROCSTAT = 'N' " & vbCr   '-- ���Ȯ������
            SQL = SQL & " ORDER BY ORDDATE, SPCID, PATNO, WORKCODE, ORDSEQNO "
            
            Call SetSQLData("��ũ��ȸ", SQL)
            
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                SPD.MaxRows = 0
                strItems = ""
                Do Until RS.EOF
                    iCnt = iCnt + 1
                    With SPD
                        .ReDraw = False
                        
                        For i = 1 To SPD.DataRowCnt
                            strDate = GetText(SPD, i, colHOSPDATE)
                            strBarcode = GetText(SPD, i, colBARCODE)
                            If Trim(RS("HOSPDATE")) = strDate And Trim(RS("BARCODE")) = strBarcode Then
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
                            'SetText SPD, Trim(RS.Fields("SEQ")) & "", intRow, colSEQNO
                            'SetText SPD, GetSampleITEM(intRow), intRow, colITEMS
                            
'                            SetText SPD, "1", intRow, colCHECKBOX
'                            SetText SPD, "20180417", intRow, colHOSPDATE
'                            SetText SPD, "1234567890", intRow, colBARCODE
'                            SetText SPD, "9876543210", intRow, colCHARTNO
'                            SetText SPD, "4561230789", intRow, colPID
'                            'SetText SPD, "126", intRow, colSEQNO
                        
                        End If
                    End With
                    
                    blnSame = False
                
                    DoEvents
                    
                    RS.MoveNext
                Loop
            Else
                frmMain.lblCommStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
            End If
            
            RS.Close
            
    End Select

     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


'-- ����� ��ȸ
Public Sub GetResultList(ByVal pFrom As String, ByVal pTo As String, ByVal pDateType As Integer, ByVal pOpt As Integer)
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
    SQL = SQL & "SELECT DISTINCT SAVESEQ,EXAMDATE,HOSPDATE,EQUIPNO,BARCODE,SAMPLETYPE,DISKNO,POSNO" & vbCr
    SQL = SQL & ",CHARTNO,INOUT,PID,PNAME,PSEX,PAGE,PJUMIN,SENDFLAG,SENDDATE,EXAMUID,HOSPITAL " & vbCr
    '-- �˻���
    SQL = SQL & ",SEQNO,EXAMNAME,RESULT,REFJUDGE" & vbCr
    
    SQL = SQL & "  FROM PATRESULT " & vbCr
    '-- �˻�����
    If pDateType = 0 Then
        SQL = SQL & " WHERE EXAMDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
    '-- ��������
    Else
        SQL = SQL & " WHERE HOSPDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
    End If
    
    '-- ����
    If pOpt = 1 Then
        SQL = SQL & "   AND SENDFLAG = '2' " & vbCr
    '-- ������
    ElseIf pOpt = 2 Then
        SQL = SQL & "   AND SENDFLAG <> '2' " & vbCr
    End If
    
    SQL = SQL & "   AND EXAMCODE IN (" & gAllTestCd & ") " & vbCr
    
    If pDateType = 0 Then
        SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,BARCODE,SEQNO"
    Else
        SQL = SQL & " ORDER BY HOSPDATE,SAVESEQ,BARCODE,SEQNO "
    End If
    
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmMain.spdROrder.MaxRows = 0
        strItems = ""
        Do Until RS.EOF
            iCnt = iCnt + 1
            With frmMain.spdROrder
                .ReDraw = False
                
'                If iCnt = 1 Then
'                    .MaxRows = .MaxRows + 1
'                    intRow = .MaxRows
'                End If
                
                strSaveSeq = GetText(frmMain.spdROrder, intRow, colSAVESEQ)
                strExamDate = GetText(frmMain.spdROrder, intRow, colEXAMDATE)
                
                'Debug.Print Trim(RS.Fields("SAVESEQ"))
                'Debug.Print Trim(RS.Fields("EXAMDATE"))
                If strSaveSeq <> Trim(RS.Fields("SAVESEQ")) & "" Or strExamDate <> Trim(RS.Fields("EXAMDATE")) & "" Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                    
                    SetText frmMain.spdROrder, "1", intRow, colCHECKBOX
                    SetText frmMain.spdROrder, Trim(RS.Fields("SAVESEQ")) & "", intRow, colSAVESEQ
                    SetText frmMain.spdROrder, Trim(RS.Fields("EXAMDATE")) & "", intRow, colEXAMDATE
                    SetText frmMain.spdROrder, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText frmMain.spdROrder, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText frmMain.spdROrder, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText frmMain.spdROrder, Trim(RS.Fields("DISKNO")) & "", intRow, colRACKNO
                    SetText frmMain.spdROrder, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText frmMain.spdROrder, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText frmMain.spdROrder, Trim(RS.Fields("PSEX")) & "", intRow, colPSEX
                    SetText frmMain.spdROrder, Trim(RS.Fields("PAGE")) & "", intRow, colPAGE
                    SetText frmMain.spdROrder, Trim(RS.Fields("PJUMIN")) & "", intRow, colPJUMIN
                    SetText frmMain.spdROrder, Trim(RS.Fields("INOUT")) & "", intRow, colINOUT
                    SetText frmMain.spdROrder, Trim(RS.Fields("EQUIPNO")) & "", intRow, colKEY1
                    
                    
                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                    Case "0"
                            SetText frmMain.spdROrder, "�����", intRow, colSTATE
                    Case "2"
                            SetText frmMain.spdROrder, "���ۿϷ�", intRow, colSTATE
                    End Select
                    SetText frmMain.spdROrder, GetSampleITEM(intRow), intRow, colITEMS
                
                End If
                
                For intCol = colSTATE + 1 To .MaxCols
                    .Row = 0
                    .Col = intCol
                    If Trim(RS.Fields("EXAMNAME")) & "" = Trim(.Text) Then
                        SetText frmMain.spdROrder, Trim(RS.Fields("RESULT")) & "", intRow, intCol
                        If Trim(RS.Fields("REFJUDGE")) & "" <> "" Then
                            SetForeColor frmMain.spdROrder, intRow, intRow, intCol, intCol, 255, 0, 0
                        End If
                        Exit For
                    End If
                
                Next
                
            End With
            DoEvents
            
            RS.MoveNext
        Loop
        frmMain.chkRAll.Value = "1"
    Else
        'frmMain.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
        frmMain.chkRAll.Value = "0"
    End If
    
    RS.Close
     
    frmMain.spdROrder.RowHeight(-1) = 12
    frmMain.spdROrder.ReDraw = True
    
    Call frmMain.GetPatTRestResult_Search(1)
    
    Screen.MousePointer = 0

End Sub

'-- �˻��� ITEM ��������
Function GetSampleITEM(ByVal asRow As Long) As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strRegDate      As String
    Dim strChartNo      As String
    Dim lngExamNo       As Long
    Dim strItems        As String
    Dim strSpcYY        As String
    Dim strSpcNo        As String
    
    GetSampleITEM = ""
    
    strRegDate = Trim(GetText(frmWorkList.spdWork, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(frmWorkList.spdWork, asRow, colBARCODE))
    strPatID = Trim(GetText(frmWorkList.spdWork, asRow, colPID))
    strChartNo = Trim(GetText(frmWorkList.spdWork, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
        
    Select Case gOCS
        Case "BIGUBCARE"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT i.IntLabCod + cast(IntLabseq as varchar(3)) AS ITEM "
            SQL = SQL & "  from interfacedb..IntRst i, aphdb..rstinf r " & vbCr
            SQL = SQL & " WHERE r.RstOdrStt not in ('OC') " & vbCr
            SQL = SQL & "   AND (r.rstrstval = '' or rstrstval is null)" & vbCr
            If gHOSP.MACHNM <> "HITACHI7080" Then
                SQL = SQL & "   AND i.intodrtyp = '" & gHOSP.PARTCD & "'" & vbCr  ''HEMO'
            End If
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
            SQL = SQL & " SELECT DISTINCT R.ResLabCod AS ITEM " & vbCr
            SQL = SQL & "   FROM RESINF AS R " & vbCr
            SQL = SQL & " WHERE LTRIM(RTRIM(R.RESOCMNUM)) = '" & strBarcode & "'" & vbLf
            SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCr         '--  'I':�߰� 'F' �Ϸ�"
            SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)" & vbCr
            
        Case "NAVY"
            SQL = ""
            SQL = SQL & "Select DISTINCT EXAMCODE as ITEM " & vbCr
            SQL = SQL & "  From SLXWORKT " & vbCr
            SQL = SQL & " Where ORDDATE =  '" & strRegDate & "'" & vbCr
            SQL = SQL & "   And HOSPID = '" & gHOSP.HOSPCD & "'" & vbCr         ' �δ��ڵ�
            SQL = SQL & "   And ROOMCODE = '" & gHOSP.PARTCD & "'" & vbCr         ' �Һ�
            SQL = SQL & "   And SPCID    = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   And WORKCODE = '" & strChartNo & "'" & vbCr
            SQL = SQL & "   And PATNO    = '" & strPatID & "'" & vbCr
            SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")" & vbCr        ' �˻��ڵ�
    End Select
            
    Call SetSQLData("ITEM��ȸ", SQL)
    
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
    
End Function

'-- �˻��� ITEM ��������
Function GetSampleITEM_Main(ByVal asRow As Long) As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strRegDate      As String
    Dim strChartNo      As String
    Dim lngExamNo       As Long
    Dim strItems        As String
    Dim strSpcYY        As String
    Dim strSpcNo        As String
    
    GetSampleITEM_Main = ""
    
    strRegDate = Trim(GetText(frmMain.spdOrder, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(frmMain.spdOrder, asRow, colBARCODE))
    strPatID = Trim(GetText(frmMain.spdOrder, asRow, colPID))
    strChartNo = Trim(GetText(frmMain.spdOrder, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
        
    Select Case gOCS
        Case "NAVY"
            SQL = ""
            SQL = SQL & "Select DISTINCT EXAMCODE as ITEM " & vbCr
            SQL = SQL & "  From SLXWORKT " & vbCr
            SQL = SQL & " Where ORDDATE =  '" & strRegDate & "'" & vbCr
            SQL = SQL & "   And HOSPID = '" & gHOSP.HOSPCD & "'" & vbCr         ' �δ��ڵ�
            SQL = SQL & "   And ROOMCODE = '" & gHOSP.PARTCD & "'" & vbCr         ' �Һ�
            SQL = SQL & "   And SPCID    = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   And WORKCODE = '" & strChartNo & "'" & vbCr
            SQL = SQL & "   And PATNO    = '" & strPatID & "'" & vbCr
            SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")" & vbCr        ' �˻��ڵ�

    End Select
            
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
    
    GetSampleITEM_Main = strItems
    
    RS.Close
    
End Function


'Public Function GetTimeFull() As String
''Server�� ���� �ð��� �����´�
''Return = 10:00:00
'    SQL = "select convert(char(8),getdate(),108) "
'    db_select_Var gServer, SQL, GetTimeFull
'End Function
'
'Public Function GetTimeShort() As String
''Server�� ���� �ð��� �����´�
''Return = 10:00
'    SQL = "select convert(char(5),getdate(),108) "
'    db_select_Var gServer, SQL, GetTimeShort
'End Function


'Public Function GetDateFull_ORCL() As String
'    Dim s As String
'    Dim t As String
'
'
''Oracle : Server�� ���� ��¥�� �����´�
'    SQL = " Select To_Char(SysDate, 'mm/dd/yyyy hh24:mi:ss') From Dual "
'
'    db_select_Var gServer, SQL, s
'
'    If Not IsDate(s) Then
'        s = Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:nn:ss")
'    End If
'
'    GetDateFull_ORCL = s
'End Function


'-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ���
Function GetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetEquipExamCode_AU480 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmMain.spdOrder, intRow, colBARCODE))    '2 ���� ���ڵ� ��ȣ
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    frmMain.vasTemp.MaxRows = 0
    
    
    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "SELECT Distinct SENDCHANNEL "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPCD  = '" & Trim(argEquipCode) & "' "
    SQL = SQL & "   AND TESTCODE in (" & Trim(gPatOrdCd) & ")"
    
'    Res = GetDBSelectRow(gLocal, SQL)
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strExamCode = ""
        Do Until RS.EOF
            If Trim(RS.Fields("SENDCHANNEL") & "") <> "" Then
                strExamCode = strExamCode & "0" & Trim(RS.Fields("SENDCHANNEL") & "") & "0"
            Else
                Exit Do
            End If
            
            RS.MoveNext
        Loop
    End If
    
    GetEquipExamCode_AU480 = strExamCode
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfo(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim lngExamNo       As Long
    Dim intTestCnt      As Integer
    
    Dim intCol     As Integer
    
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    
    Dim strSpcYY    As String
    Dim strSpcNo    As String
    
    Dim intBcNow    As Integer
    Dim intBcFive   As Integer
    Dim intBcAdd    As Integer
    Dim strADT      As String
    Dim strSlip1    As String
    Dim strSlip2    As String
    
    Dim sqlRet      As Integer
    
    '-- ���ڵ� ��ȣ�� ���� ��ȸ
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    
    Dim strDate As String
    
On Error GoTo DBErr
    
    GetSampleInfo = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    
    'strPatID = Trim(GetText(SPD, asRow, colPID))
    'strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    
            
    If strBarcode = "" And Len(strBarcode) <> 12 Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    Select Case gOCS
        Case "KYU"
            strDate = Format(Now, "yyyy-mm-dd")
            intBcNow = DateDiff("d", "1999-01-01", strDate)
            intBcFive = Mid(strBarcode, 1, 5) '06351
            intBcAdd = intBcFive - intBcNow
            'strADT = Format(Now + intBcAdd, "yyyymmdd")
            strADT = Format(Now + intBcAdd, "yyyy-mm-dd")
            strSlip1 = Mid(strBarcode, 6, 2)  ''10���� �����ϸ� TLA���ν����� �¿�
                                                '�������� EXAM_INTERFACE_S
            strSlip2 = Mid(strBarcode, 8, 5)  '00001
           
            '-- SP ���
            Set AdoCmd = New ADODB.Command
            Set AdoCmd.ActiveConnection = AdoCn
        
            AdoCmd.CommandTimeout = 15
            If strSlip1 = "10" Then
                AdoCmd.CommandText = "TW_MIS_EXAM.EXAM_TLA_INTERFACE_S"
            Else
                AdoCmd.CommandText = "EXAM_INTERFACE_S"
            End If
            
            AdoCmd.CommandType = adCmdStoredProc
            
            If strSlip1 = "10" Then
                Set prm1 = AdoCmd.CreateParameter("I_JEOBSUDT", adDate, adParamInput, 10, strADT)
                AdoCmd.Parameters.Append prm1
                Set prm2 = AdoCmd.CreateParameter("I_BARCODE", adDouble, adParamInput, 12, strBarcode)
                AdoCmd.Parameters.Append prm2
            Else
                Set prm1 = AdoCmd.CreateParameter("I_JEOBSUDT", adDate, adParamInput, 10, strADT)
                AdoCmd.Parameters.Append prm1
                Set prm2 = AdoCmd.CreateParameter("I_SLIPNO1", adInteger, adParamInput, 2, strSlip1)
                AdoCmd.Parameters.Append prm2
                Set prm3 = AdoCmd.CreateParameter("I_SLIPNO2", adInteger, adParamInput, 5, strSlip2)
                AdoCmd.Parameters.Append prm3
            End If
            
            Set RS = New ADODB.Recordset
            RS.Open AdoCmd.Execute
            
            Call SetText(SPD, strADT, asRow, colHOSPDATE)
            
            intTestCnt = 0
            
'            If Not RS.EOF = True And Not RS.BOF = True Then
'                Do Until RS.EOF
'                    With vaSpread1
'                        .ReDraw = False
'                        intTestCnt = intTestCnt + 1
'                        vaSpread1.MaxRows = intTestCnt
'
'                        SetText vaSpread1, Trim(RS.Fields("ptno")) & "", intTestCnt, 1
'                        SetText vaSpread1, Trim(RS.Fields("sname")) & "", intTestCnt, 2
'                        SetText vaSpread1, Trim(RS.Fields("sex")) & "", intTestCnt, 3
'                        SetText vaSpread1, Trim(RS.Fields("ageyy")) & "", intTestCnt, 4
'                        SetText vaSpread1, Trim(RS.Fields("deptcode")) & "", intTestCnt, 5
'        '                SetText vaSpread1, Trim(RS.Fields("gber")) & "", intTestCnt, 6
'                        SetText vaSpread1, Trim(RS.Fields("slipno1")) & "", intTestCnt, 7
'                        SetText vaSpread1, Trim(RS.Fields("slipno2")) & "", intTestCnt, 8
'                        SetText vaSpread1, Trim(RS.Fields("itemcd")) & "", intTestCnt, 9
'                        SetText vaSpread1, Trim(RS.Fields("itemnm")) & "", intTestCnt, 10
'                        SetText vaSpread1, Trim(RS.Fields("geomchc1")) & "", intTestCnt, 11
'                        SetText vaSpread1, Trim(RS.Fields("status")) & "", intTestCnt, 12
'                        SetText vaSpread1, Trim(RS.Fields("result1")) & "", intTestCnt, 13
'
'                    End With
'                    DoEvents
'
'                    RS.MoveNext
'                Loop
'            End If
'
'            RS.Close
            
            
            SetText SPD, "0", asRow, colCHECKBOX
            
            intTestCnt = 0
            
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    With SPD
                        .ReDraw = False
                        intTestCnt = intTestCnt + 1
                        'SPD.MaxRows = intTestCnt
                        SetText SPD, "1", asRow, colCHECKBOX
                        SetText SPD, strDate, asRow, colHOSPDATE
                        SetText SPD, Trim(RS.Fields("ptno")) & "", asRow, colPID
                        mOrder.PID = Trim(RS.Fields("ptno")) & ""
                        SetText SPD, Trim(RS.Fields("slipno1")) & "", asRow, colRACKNO
                        SetText SPD, Trim(RS.Fields("slipno2")) & "", asRow, colPOSNO
                        SetText SPD, Trim(RS.Fields("sname")) & "", asRow, colPNAME
                        mOrder.PName = Trim(RS.Fields("sname")) & ""
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT

                        frmMain.txtRcv.Text = frmMain.txtRcv.Text & Trim(RS.Fields("itemcd")) & vbCr
                        
                        '-- ȭ�鿡 ǥ��
                        Dim strTestNm   As String
                        frmMain.spdPatOrder.MaxRows = 0
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("itemcd")) = gArrEQP(intCol - colSTATE, 2) Then
                                strTestNm = GetTestNm(Trim(RS.Fields("itemcd")))
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "��", asRow, intCol)

                                '-- ������ �˻��� �˻��ڵ带 ã�Ƴ��´�
                                frmMain.spdPatOrder.MaxRows = frmMain.spdPatOrder.MaxRows + 1
                                Call SetText(frmMain.spdPatOrder, strTestNm, frmMain.spdPatOrder.MaxRows, 1)
                                Call SetText(frmMain.spdPatOrder, Trim(RS.Fields("itemcd")), frmMain.spdPatOrder.MaxRows, 2)

                                Exit For
                            End If
                        Next
                        gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("itemcd")) & "',"
                        
                    End With
                    DoEvents
                    
                    RS.MoveNext
                Loop
            End If
            
            RS.Close
    
            '-- ���̺� ���
'            SQL = ""
'            SQL = SQL & "SELECT DISTINCT To_Char(R.jeobsudt, 'yyyymmdd') as HOSPDATE, R.slipno1 as RACK, R.slipno2 as POS, R.ptno as PID, p.sname as PNAME, R.itemcd as ITEM " & vbCr
'            SQL = SQL & "  FROM twexam_general_sub R, twexam_general O, twbas_patient p" & vbCr
'            SQL = SQL & " WHERE r.verify <> 'Y' " & vbCr
'            SQL = SQL & "   AND O.gbch = 'Y' " & vbCr
'            SQL = SQL & "   AND R.jeobsudt = to_date('" & strADT & "','yyyymmdd')" & vbCr
'            SQL = SQL & "   AND R.slipno1 = '" & strSlip1 & "'" & vbCr
'            SQL = SQL & "   AND R.slipno2 = '" & strSlip2 & "'" & vbCr
'            SQL = SQL & "   AND R.itemcd IN (" & gAllTestCd & ")" & vbCr
'            SQL = SQL & "   AND R.jeobsudt = O.jeobsudt" & vbCr
'            SQL = SQL & "   AND R.slipno1 = O.slipno1" & vbCr
'            SQL = SQL & "   AND R.slipno2 = O.slipno2" & vbCr
'            SQL = SQL & "   AND R.PTNO = O.PTNO" & vbCr
'            SQL = SQL & "   AND R.PTNO = p.PTNO"
'
'            Call SetSQLData("���ڵ���ȸ", SQL)
'
'            '-- Record Count ������
'            AdoCn.CursorLocation = adUseClient
'            Set RS = AdoCn.Execute(SQL, , 1)
'
'
'            SetText SPD, "0", asRow, colCHECKBOX
'
'            If Not RS.EOF = True And Not RS.BOF = True Then
'                Do Until RS.EOF
'                    With SPD
'                        .ReDraw = False
'                        intTestCnt = intTestCnt + 1
'                        SetText SPD, "1", asRow, colCHECKBOX
'                        SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
'                        SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
'                        mOrder.PID = Trim(RS.Fields("PID")) & ""
'                        SetText SPD, Trim(RS.Fields("RACK")) & "", asRow, colRACKNO
'                        SetText SPD, Trim(RS.Fields("POS")) & "", asRow, colPOSNO
'                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
'                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
'                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
'
'                        '-- ȭ�鿡 ǥ��
'                        For intCol = colSTATE + 1 To .MaxCols
'                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
'                                .Row = asRow
'                                .Col = intCol
'                                .BackColor = vbYellow
'                                Call SetText(SPD, "��", asRow, intCol)
'
'                                Exit For
'                            End If
'                        Next
'                        gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
'
'                    End With
'                    DoEvents
'
'                    RS.MoveNext
'                Loop
'            End If
'
'            RS.Close
        
        Case "BIGUBCARE"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT i.IntOdrDte AS HOSPDATE"
            SQL = SQL & ", i.IntLabNum AS BARCODE "           ' �˻��ȣ"
            SQL = SQL & ", i.IntChtNum AS CHARTNO "           ' ��Ʈ��ȣ"
            SQL = SQL & ", i.IntPatNam AS PNAME "             ' ȯ�ڸ�"
            SQL = SQL & ", i.IntSexTyp AS SEX"                ' ����"
            'SQL = SQL & ", i.IntOdrStt AS STATE               ' ����"
            SQL = SQL & ", i.IntEmgYon AS INOUT"              ' ���޿���"
            SQL = SQL & ", i.IntLabCod + cast(IntLabseq as varchar(3))  AS ITEM "
            SQL = SQL & ", i.IntLabSeq AS SUBITEM "
            SQL = SQL & "  from interfacedb..IntRst i, aphdb..rstinf r " & vbCr
            SQL = SQL & " WHERE r.RstOdrStt not in ('OC') " & vbCr
            SQL = SQL & "   AND (r.rstrstval = '' or rstrstval is null)" & vbCr
            If gHOSP.MACHNM <> "HITACHI7080" Then
                SQL = SQL & "   AND i.intodrtyp = '" & gHOSP.PARTCD & "'" & vbCr  ''HEMO'
            End If
            
            'SQL = SQL & "   AND i.IntOdrDte BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
            SQL = SQL & "   AND i.IntLabNum = '" & strBarcode & "'" & vbCr
            'SQL = SQL & "   AND i.IntLabCod + cast(IntLabseq as varchar(2)) IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND i.IntLabCod + cast(IntLabseq as varchar(3)) IN (" & gAllTestCd & ")" & vbCr

            SQL = SQL & "   AND i.intlabnum = r.rstlabnum" & vbCr
            SQL = SQL & "   AND i.intodrdte = r.rstodrdte" & vbCr
            SQL = SQL & "   AND i.intlabseq = r.rstlabseq" & vbCr
            SQL = SQL & "   AND i.intlabcod = r.rstodrcod" & vbCr
            
            Call SetSQLData("���ڵ���ȸ", SQL)
            
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            
            SetText SPD, "0", asRow, colCHECKBOX
            
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    With SPD
                        .ReDraw = False
                        intTestCnt = intTestCnt + 1
                        SetText SPD, "1", asRow, colCHECKBOX
                        SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                        'SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                        SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                        mOrder.PID = Trim(RS.Fields("CHARTNO")) & ""
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                        SetText SPD, Trim(RS.Fields("INOUT")) & "", asRow, colINOUT
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                
                        '-- ȭ�鿡 ǥ��
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "��", asRow, intCol)
                                
                                '-- �������� SEQ
                                gArrEQP(intCol - colSTATE, 7) = Trim(RS.Fields("SUBITEM")) & ""   '�������� ��ȣ's
                                
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
            
        Case "BIT"
            SQL = ""
            SQL = SQL & " SELECT DISTINCT SUBSTRING(O.OCMACPDTM,1,8) AS HOSPDATE" & vbCr
            SQL = SQL & "        ,R.RESOCMNUM AS BARCODE "
            SQL = SQL & "        ,O.OCMCHTNUM AS CHARTNO "
            SQL = SQL & "        ,R.RESOCMNUM AS PID " & vbCr
            SQL = SQL & "        ,P.PBSPATNAM AS PNAME"
            SQL = SQL & "        ,P.PBSSEXTYP AS SEX"
            SQL = SQL & "        ,R.ResLabCod AS ITEM " & vbCr
            SQL = SQL & "        ,R.ResOdrSeq, R.ResSeq, R.ResSubSeq " & vbCr
            SQL = SQL & "   FROM RESINF AS R, OCMINF AS O, PBSINF AS P, LABMST AS E" & vbCr
            SQL = SQL & " WHERE ltrim(rtrim(R.RESOCMNUM)) = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')" & vbCr
            SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM" & vbCr
            SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM" & vbCr
            SQL = SQL & "   AND R.RESLABCOD = E.LABCOD" & vbCr
            
            Call SetSQLData("���ڵ���ȸ", SQL)
            
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            
            SetText SPD, "0", asRow, colCHECKBOX
            
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    With SPD
                        .ReDraw = False
                        intTestCnt = intTestCnt + 1
                        SetText SPD, "1", asRow, colCHECKBOX
                        SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                        SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                        SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                
                        '-- ȭ�鿡 ǥ��
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "��", asRow, intCol)
                                
                                '-- �������� SEQ
                                gArrEQP(intCol - colSTATE, 7) = Trim(RS.Fields("ResOdrSeq")) & "|" & Trim(RS.Fields("ResSeq")) & "|" & Trim(RS.Fields("ResSubSeq"))   '�������� ��ȣ's
                                
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
            
        Case "NAVY"
            SQL = ""
            SQL = SQL & "Select DISTINCT ORDDATE as HOSPDATE"
            SQL = SQL & ", SPCID as BARCODE "
            SQL = SQL & ", PATNO as PID "
            SQL = SQL & ", WORKCODE as CHARTNO "
            SQL = SQL & ", EXAMCODE as ITEM " & vbCr
            SQL = SQL & "  From SLXWORKT " & vbCr
            SQL = SQL & " Where ORDDATE  = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   And HOSPID   = '" & gHOSP.HOSPCD & "'" & vbCr         ' �δ��ڵ�
            SQL = SQL & "   And ROOMCODE = '" & gHOSP.PARTCD & "'" & vbCr         ' �Һ�
            SQL = SQL & "   And SPCID    = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   And WORKCODE = '" & strChartNo & "'" & vbCr
            SQL = SQL & "   And PATNO    = '" & strPatID & "'" & vbCr
            SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")" & vbCr        ' �˻��ڵ�
            
            Call SetSQLData("���ڵ���ȸ", SQL)
            
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            
            SetText SPD, "0", asRow, colCHECKBOX
            
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    With SPD
                        .ReDraw = False
                        intTestCnt = intTestCnt + 1
                        SetText SPD, "1", asRow, colCHECKBOX
                        SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                        SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                        SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                        SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                        'SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                
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
                        
                    End With
                    DoEvents
                    
                    RS.MoveNext
                Loop
            End If
            
            RS.Close
    End Select
            

    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
    With frmMain
        sExamDate = Format(.dtpToday, "yyyymmdd")
        If Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) = "" Then
            Exit Function
        End If
        
        SQL = ""
        SQL = SQL & "DELETE FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EXAMDATE = '" & sExamDate & "' " & vbCr
        SQL = SQL & "   AND EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCr
        SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) & vbCr
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "' " & vbCr
        SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'" & vbCr
        SQL = SQL & "   AND EXAMCODE = '" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'"
        
        If DBExec(AdoCn_Local, SQL) Then
            SQL = ""
            SQL = SQL & "INSERT INTO PATRESULT (" & vbCr
            SQL = SQL & "SAVESEQ"                           '�������(��¥��)
            SQL = SQL & ", EXAMDATE"                        '�˻�����"
            SQL = SQL & ", HOSPDATE"                        '������������"
            SQL = SQL & ", EQUIPNO"                         '����ڵ�"
            SQL = SQL & ", BARCODE" & vbCrLf                '��ü��ȣ
            SQL = SQL & ", EQUIPCODE"                       '�˻�ä��"
            SQL = SQL & ", ORDERCODE"                       '����ó���ڵ�"
            SQL = SQL & ", EXAMCODE"                        '�����˻��ڵ�"
            SQL = SQL & ", EXAMSUBCODE"                     '�����˻��ڵ�(SUB)"
            SQL = SQL & ", EXAMNAME"                        '�˻��
            SQL = SQL & ", SEQNO" & vbCrLf                  '�˻��Ϸù�ȣ"
            SQL = SQL & ", SAMPLETYPE"                      '��ü����"
            SQL = SQL & ", INOUT"                           '��/��
            SQL = SQL & ", DISKNO"                          'Rack (VERSACELL ������ ���� �˻�����ڵ带 �����Ѵ�..)
            SQL = SQL & ", POSNO"                           'Pos
            SQL = SQL & ", EQUIPRESULT"                     '�����"
            SQL = SQL & ", RESULT" & vbCrLf                 'LIS ���"
            SQL = SQL & ", REFJUDGE"                        '����
            SQL = SQL & ", REFFLAG"                         'flag
            SQL = SQL & ", REFVALUE"                        '����ġ
            SQL = SQL & ", CHARTNO"                         'íƮ��ȣ
            SQL = SQL & ", PID"                             '���Ϲ�ȣ(������ȣ)"
            SQL = SQL & ", PNAME" & vbCrLf
            SQL = SQL & ", PSEX"
            SQL = SQL & ", PAGE"
            SQL = SQL & ", PJUMIN"
            SQL = SQL & ", PANICVALUE"
            SQL = SQL & ", DELTAVALUE" & vbCrLf
            SQL = SQL & ", SENDFLAG"                        '���۱���(0:������,1:����)"
            SQL = SQL & ", SENDDATE"
            SQL = SQL & ", EXAMUID"
            SQL = SQL & ", HOSPITAL)" & vbCrLf
            SQL = SQL & " VALUES (" & vbCrLf
            SQL = SQL & Trim(GetText(.spdOrder, asRow1, colSAVESEQ))
            SQL = SQL & ",'" & sExamDate & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "'"
            SQL = SQL & ",'" & gHOSP.MACHCD & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRORDERCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRSUBCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTNM)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRSEQNO)) & "'"
            SQL = SQL & ",''"                                                   '��ü����
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colINOUT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colRACKNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPOSNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRMACHRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRLISRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRJUDGE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRFLAG)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRREF)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colCHARTNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPID)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPNAME)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPSEX)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPAGE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPJUMIN)) & "'"
            'SQL = SQL & ",'" & mOrder.AccSeq & "'"                              'panic (accseq ��ü���)
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colKEY1)) & "'"                              'panic (accseq ��ü���)
            SQL = SQL & ",''"                                                   'delta
            SQL = SQL & ",'0'"                                                  '���۱���(0:������,1:����)
            SQL = SQL & ",''"
            SQL = SQL & ",'" & gHOSP.USERID & "'"
            SQL = SQL & ",'" & gHOSP.HOSPNM & "')"
            
            If Not DBExec(AdoCn_Local, SQL) Then
                'SaveQuery SQL
                Exit Function
            End If
            
        End If
        
    End With
    
End Function

Function SetLocalDB_R(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
    With frmMain
        sExamDate = Trim(GetText(.spdROrder, asRow1, colEXAMDATE))
        If Trim(GetText(.spdROrder, asRow1, colSAVESEQ)) = "" Then
            Exit Function
        End If
        
        SQL = ""
        SQL = SQL & "DELETE FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EXAMDATE = '" & sExamDate & "' " & vbCr
        SQL = SQL & "   AND EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCr
        SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(.spdROrder, asRow1, colSAVESEQ)) & vbCr
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.spdROrder, asRow1, colBARCODE)) & "' " & vbCr
        SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(.spdRResult, asRow2, colRCHANNEL)) & "'" & vbCr
        SQL = SQL & "   AND EXAMCODE = '" & Trim(GetText(.spdRResult, asRow2, colRTESTCD)) & "'"
        
        If DBExec(AdoCn_Local, SQL) Then
            SQL = ""
            SQL = SQL & "INSERT INTO PATRESULT (" & vbCr
            SQL = SQL & "SAVESEQ"                           '�������(��¥��)
            SQL = SQL & ", EXAMDATE"                        '�˻�����"
            SQL = SQL & ", HOSPDATE"                        '������������"
            SQL = SQL & ", EQUIPNO"                         '����ڵ�"
            SQL = SQL & ", BARCODE" & vbCrLf                '��ü��ȣ
            SQL = SQL & ", EQUIPCODE"                       '�˻�ä��"
            SQL = SQL & ", ORDERCODE"                       '����ó���ڵ�"
            SQL = SQL & ", EXAMCODE"                        '�����˻��ڵ�"
            SQL = SQL & ", EXAMSUBCODE"                     '�����˻��ڵ�(SUB)"
            SQL = SQL & ", EXAMNAME"                        '�˻��
            SQL = SQL & ", SEQNO" & vbCrLf                  '�˻��Ϸù�ȣ"
            SQL = SQL & ", SAMPLETYPE"                      '��ü����"
            SQL = SQL & ", INOUT"                           '��/��
            SQL = SQL & ", DISKNO"                          'Rack (VERSACELL ������ ���� �˻�����ڵ带 �����Ѵ�..)
            SQL = SQL & ", POSNO"                           'Pos
            SQL = SQL & ", EQUIPRESULT"                     '�����"
            SQL = SQL & ", RESULT" & vbCrLf                 'LIS ���"
            SQL = SQL & ", REFJUDGE"                        '����
            SQL = SQL & ", REFFLAG"                         'flag
            SQL = SQL & ", REFVALUE"                        '����ġ
            SQL = SQL & ", CHARTNO"                         'íƮ��ȣ
            SQL = SQL & ", PID"                             '���Ϲ�ȣ(������ȣ)"
            SQL = SQL & ", PNAME" & vbCrLf
            SQL = SQL & ", PSEX"
            SQL = SQL & ", PAGE"
            SQL = SQL & ", PJUMIN"
            SQL = SQL & ", PANICVALUE"
            SQL = SQL & ", DELTAVALUE" & vbCrLf
            SQL = SQL & ", SENDFLAG"                        '���۱���(0:������,1:����)"
            SQL = SQL & ", SENDDATE"
            SQL = SQL & ", EXAMUID"
            SQL = SQL & ", HOSPITAL)" & vbCrLf
            SQL = SQL & " VALUES (" & vbCrLf
            SQL = SQL & Trim(GetText(.spdROrder, asRow1, colSAVESEQ))
            SQL = SQL & ",'" & sExamDate & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colHOSPDATE)) & "'"
            SQL = SQL & ",'" & gHOSP.MACHCD & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colBARCODE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRCHANNEL)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRORDERCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRTESTCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRSUBCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRTESTNM)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRSEQNO)) & "'"
            SQL = SQL & ",''"                                                   '��ü����
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colINOUT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colRACKNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPOSNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRMACHRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRLISRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRJUDGE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRFLAG)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdRResult, asRow2, colRREF)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colCHARTNO)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPID)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPNAME)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPSEX)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPAGE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPJUMIN)) & "'"
            SQL = SQL & ",''"                                                   'panic
            SQL = SQL & ",''"                                                   'delta
            SQL = SQL & ",'0'"                                                  '���۱���(0:������,1:����)
            SQL = SQL & ",''"
            SQL = SQL & ",'" & gHOSP.USERID & "'"
            SQL = SQL & ",'" & gHOSP.HOSPNM & "')"
            
            If Not DBExec(AdoCn_Local, SQL) Then
                'SaveQuery SQL
                Exit Function
            End If
            
        End If
        
    End With
    
End Function


Function SetLocalDB_TV(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
    With frmMain
        sExamDate = Trim(GetText(.spdROrder, asRow1, colEXAMDATE))
        If Trim(GetText(.spdROrder, asRow1, colSAVESEQ)) = "" Then
            Exit Function
        End If
        
        SQL = ""
        SQL = SQL & "SELECT COUNT(*) AS CNT FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EXAMDATE = '" & sExamDate & "' " & vbCr
        SQL = SQL & "   AND EQUIPNO = '" & Trim(GetText(.spdROrder, asRow1, colKEY1)) & "' " & vbCr
        SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(.spdROrder, asRow1, colSAVESEQ)) & vbCr
        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(.spdROrder, asRow1, colBARCODE)) & "' " & vbCr
        SQL = SQL & "   AND EXAMCODE = '24HRS-V' " & vbCr
        Set RS = AdoCn_Local.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            If Trim(RS.Fields("CNT") & "") = 0 Then
                'insert into
            Else
                'update
                GoTo UPDATE
            End If
        End If
            
        
        SQL = ""
        SQL = SQL & "INSERT INTO PATRESULT (" & vbCr
        SQL = SQL & "SAVESEQ"                           '�������(��¥��)
        SQL = SQL & ", EXAMDATE"                        '�˻�����"
        SQL = SQL & ", HOSPDATE"                        '������������"
        SQL = SQL & ", EQUIPNO"                         '����ڵ�"
        SQL = SQL & ", BARCODE" & vbCrLf                '��ü��ȣ
        SQL = SQL & ", EQUIPCODE"                       '�˻�ä��"
        SQL = SQL & ", ORDERCODE"                       '����ó���ڵ�"
        SQL = SQL & ", EXAMCODE"                        '�����˻��ڵ�"
        SQL = SQL & ", EXAMSUBCODE"                     '�����˻��ڵ�(SUB)"
        SQL = SQL & ", EXAMNAME" & vbCrLf               '�˻��
        SQL = SQL & ", SEQNO"                           '�˻��Ϸù�ȣ"
        SQL = SQL & ", SAMPLETYPE"                      '��ü����"
        SQL = SQL & ", INOUT"                           '��/��
        SQL = SQL & ", DISKNO"                          'Rack (VERSACELL ������ ���� �˻�����ڵ带 �����Ѵ�..)
        SQL = SQL & ", POSNO" & vbCrLf                  'Pos
        SQL = SQL & ", EQUIPRESULT"                     '�����"
        SQL = SQL & ", RESULT"                          'LIS ���"
        SQL = SQL & ", REFJUDGE"                        '����
        SQL = SQL & ", REFFLAG"                         'flag
        SQL = SQL & ", REFVALUE" & vbCrLf               '����ġ
        SQL = SQL & ", CHARTNO"                         'íƮ��ȣ
        SQL = SQL & ", PID"                             '���Ϲ�ȣ(������ȣ)"
        SQL = SQL & ", PNAME"
        SQL = SQL & ", PSEX"
        SQL = SQL & ", PAGE" & vbCrLf
        SQL = SQL & ", PJUMIN"
        SQL = SQL & ", PANICVALUE"
        SQL = SQL & ", DELTAVALUE"
        SQL = SQL & ", SENDFLAG"                        '���۱���(0:������,1:����)"
        SQL = SQL & ", SENDDATE" & vbCrLf
        SQL = SQL & ", EXAMUID"
        SQL = SQL & ", HOSPITAL)" & vbCrLf
        SQL = SQL & " VALUES (" & vbCrLf
        SQL = SQL & Trim(GetText(.spdROrder, asRow1, colSAVESEQ))
        SQL = SQL & ",'" & sExamDate & "'"
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colHOSPDATE)) & "'"
        SQL = SQL & ",'" & gHOSP.MACHCD & "'"
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colBARCODE)) & "'" & vbCr
        SQL = SQL & ",''"
        SQL = SQL & ",''"
        SQL = SQL & ",'24HRS-V'"
        SQL = SQL & ",''"
        SQL = SQL & ",'Total Volum'" & vbCr
        SQL = SQL & ",'123'"
        SQL = SQL & ",''"                                                   '��ü����
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colINOUT)) & "'"
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colRACKNO)) & "'"
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPOSNO)) & "'" & vbCr
        SQL = SQL & ",'" & asEquipResult & "'"
        SQL = SQL & ",'" & asEquipResult & "'"
        SQL = SQL & ",''"
        SQL = SQL & ",''"
        SQL = SQL & ",''" & vbCr
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colCHARTNO)) & "'"
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPID)) & "'"
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPNAME)) & "'"
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPSEX)) & "'"
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPAGE)) & "'" & vbCr
        SQL = SQL & ",'" & Trim(GetText(.spdROrder, asRow1, colPJUMIN)) & "'"
        SQL = SQL & ",''"                                                   'panic
        SQL = SQL & ",''"                                                   'delta
        SQL = SQL & ",'0'"                                                  '���۱���(0:������,1:����)
        SQL = SQL & ",''" & vbCr
        SQL = SQL & ",'" & gHOSP.USERID & "'"
        SQL = SQL & ",'" & gHOSP.HOSPNM & "')"
        
        If Not DBExec(AdoCn_Local, SQL) Then
            'SaveQuery SQL
            'Exit Function
        End If
            
'        Call CalProcess(gRow)
        
        Exit Function
UPDATE:
        SQL = ""
        SQL = SQL & "UPDATE PATRESULT SET"
        SQL = SQL & " EQUIPRESULT = '" & asEquipResult & "'"                                            '�����
        SQL = SQL & ",RESULT      = '" & asEquipResult & "'" & vbCr                                     'LIS ���
        SQL = SQL & " WHERE SAVESEQ  = " & Trim(GetText(.spdROrder, asRow1, colSAVESEQ)) & vbCr         '�������(��¥��)
        SQL = SQL & "   AND EXAMDATE = '" & sExamDate & "'" & vbCr                                      '�˻�����
        SQL = SQL & "   AND HOSPDATE = '" & Trim(GetText(.spdROrder, asRow1, colHOSPDATE)) & "'" & vbCr '������������
        SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCr                                   '����ڵ�
        SQL = SQL & "   AND BARCODE  = '" & Trim(GetText(.spdROrder, asRow1, colBARCODE)) & "'" & vbCr  '��ü��ȣ
        SQL = SQL & "   AND EXAMCODE = '24HRS-V'"
        If Not DBExec(AdoCn_Local, SQL) Then
            'SaveQuery SQL
            'Exit Function
        End If
        
'        Call CalProcess(gRow)
        
    End With
    
End Function

'-- ��갪 ó��
'01   Serum (SST)
'02   EDTA
'03   S.citrate
'04   Urine
'05   CSF
'07   Stool
'11  Pleural fluid
'20  ����
'22  Biopsy
Public Function CalProcess(ByVal SPDORD As Object, ByVal SPDRST As Object, ByVal pTestCd As String, Optional ByVal pTV As String)
    Dim RS              As ADODB.Recordset
    Dim RS_L            As ADODB.Recordset
    Dim strBarcode      As String
    Dim intRow          As Integer
    Dim strSex          As String
    Dim strAge          As String
    Dim strSpc          As String
    Dim strResult       As String
    Dim strCalTestCd    As String
    Dim strCalResult    As String
    Dim strPreResult    As String
    
    Dim strIntBase      As String
    Dim lsOrderCode     As String
    Dim lsTestCode      As String
    Dim lsTestName      As String
    Dim lsSeqNo         As String
    Dim lsRstRow        As Integer
    Dim intCol          As Integer
    Dim Res             As Integer
    Dim ActiveRow       As Integer
    Dim strPtId         As String
    
    If pTV = "" Then
        ActiveRow = gRow ' SPDORD.ActiveRow
    Else
        ActiveRow = frmMain.lblPatInfo(3).Caption
    End If
    
    strAge = ""
    strSex = ""
    strSpc = ""
    strResult = ""
    strCalResult = ""
    
    strBarcode = Trim(GetText(SPDORD, ActiveRow, colBARCODE))
        
    If Not IsNumeric(strBarcode) Then
        Exit Function
    End If
    
    '1. ��� ��� �˻��׸��� ã�´�.
    Select Case pTestCd
        Case "C3730N1"  ' : ������ = C3730N1
                strCalTestCd = "URR"        '��Ұ�����
        Case "C3750"    'Creatine
                strCalTestCd = "EGFR"       'MDRD eGFR
        Case "C3791" 'NA
                strCalTestCd = "C3791N1"    'NA(24�ð���)
        Case "C3792" 'K
                strCalTestCd = "C3792N1"    'K(24�ð���)
        Case "C3793" 'Cl
                strCalTestCd = "C3793N1"    'Cl(24�ð���)
        Case "C2200-1" 'micro TP
                strCalTestCd = "C2200-2"    'UTP(24�ð���)
        Case "C3730" 'BUN
                strCalTestCd = "C3730-2"    'BUN(24�ð���)
        Case "C3750N1" 'Crea
                strCalTestCd = "C3750N1"   'Crea(24�ð���)
        Case "C3750N3"  'Crea(��ȸ��)
                strCalTestCd = "C7230"      'MicroALB retio
        Case "C2302N6"    'M.alb
                strCalTestCd = "C7230"      'MicroALB retio
        Case Else
                Exit Function
    End Select
    
    '1. ���ó���׸��� �ִ��� ã�´�.
          SQL = ""
    SQL = SQL & "SELECT COUNT(*) AS CNT" & vbCr
    SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
    SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
    SQL = SQL & "   AND ORD_CD = '" & strCalTestCd & "'"
        
    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        If IsNull(RS.Fields("CNT")) Or RS.Fields("CNT") = 0 Then
            Exit Function
        End If
    End If
    RS.Close
    
    
    '2. ����� ȯ�������� ������� �����´�.
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & " READING_YMD AS HOSPDATE"
    SQL = SQL & ", BCODE_NO AS BARCODE"
    SQL = SQL & ", PTNT_NO AS PID"
    SQL = SQL & " ,PTNT_NM AS PNAME"
    SQL = SQL & " ,AGE AS AGE"
    SQL = SQL & " ,SEX AS SEX"
    SQL = SQL & " ,IO_GB AS INOUT"
    SQL = SQL & " ,ORD_CD AS ITEM" & vbCr
    SQL = SQL & " ,SP_CD AS SPCCD" & vbCr
    SQL = SQL & " ,RESULT_NM AS RESULT" & vbCr
    SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
    SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
    SQL = SQL & "   AND ORD_CD = '" & pTestCd & "'"
    'SQL = SQL & "   AND STS_CD = '0'" & vbCr    '0 ����, 1:�������
        
    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strResult = Trim(RS.Fields("RESULT")) & ""
        strAge = Trim(RS.Fields("AGE")) & ""
        strSex = Trim(RS.Fields("SEX")) & ""
        strSpc = Trim(RS.Fields("SPCCD")) & ""
        strPtId = Trim(RS.Fields("PID")) & ""
    End If
    
    If strResult = "" Then
        '�������� ��ã�� ���..(����)
        SQL = ""
        SQL = SQL & "SELECT RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE BARCODE = '" & strBarcode & "'" & vbCr
        SQL = SQL & "   AND EXAMCODE = '" & pTestCd & "'"

        '-- Record Count ������
        AdoCn_Local.CursorLocation = adUseClient
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            strResult = Trim(RS_L.Fields("RESULT")) & ""
        End If

        RS_L.Close
    End If
    
    RS.Close
    
    Select Case pTestCd
        Case "URR"    '��Ұ��ҿ�
            'ȯ�ڹ�ȣ,������,���ؽð�,ó���ڵ��Դϴ�.
            '���������[C3730N2]�� �����´�.
            SQL = "SELECT [dbo].FUN_H7LIS_PRE_RESULT4('" & strPtId & "', '" & Format(Now, "yyyymmdd") & "', '" & Format(Now, "hhmm") & "', 'C3730N2')"
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                strPreResult = Trim(RS.Fields(0)) & ""
            End If
            If strPreResult = "" Then
                Exit Function
            Else
                If IsNumeric(strResult) And CCur(strResult) > 0 And IsNumeric(strPreResult) And CCur(strPreResult) > 0 Then
                    strCalResult = 1 - (strResult / strPreResult)
                Else
                    Exit Function
                End If
            End If
        Case "C3750"    'Creatine   ==> eGFR ���
            If IsNumeric(strResult) And CCur(strResult) > 0 And strSex <> "" And strAge <> "" Then
                '18�� �̻� ����
                If strAge > 18 Then
                    If strSex = "M" Then
                        strCalResult = 186 * (strResult ^ -1.154) * (strAge ^ -0.203)
                    ElseIf strSex = "F" Then
                        strCalResult = 186 * (strResult ^ -1.154) * (strAge ^ -0.203) * 0.742
                    End If
                    
                    If strCalResult <> "" Then
                        strCalResult = Format(strCalResult, "##0.00")
                    End If
                End If
            Else
                strCalResult = ""
            End If
            
        Case "C3791", "C3792", "C3793"  'NA,K,Cl
            If IsNumeric(strResult) Then
                strCalResult = strResult * CCur(pTV)
                strCalResult = Format(strCalResult, "#,##0.0")
            Else
                strCalResult = ""
            End If
            
        Case "C2200-1" 'micro TP
            If IsNumeric(strResult) Then
                strCalResult = strResult * 10 * CCur(pTV)
                strCalResult = Format(strCalResult, "#,##0.0")
            Else
                strCalResult = ""
            End If
            
        Case "C3730" 'BUN
            If IsNumeric(strResult) Then
                strCalResult = strResult * 10 * CCur(pTV)
                strCalResult = Format(strCalResult, "#,##0.0")
            Else
                strCalResult = ""
            End If
            
        Case "C3750N1" 'Crea
            If IsNumeric(strResult) Then
                strCalResult = (strResult * 10 * CCur(pTV)) / 1000
                strCalResult = Format(strCalResult, "#,##0.00")
            Else
                strCalResult = ""
            End If

        Case "C3750N3"      'Crea(��ȸ��) �̸� M.alb ����� �����´�.
            SQL = "SELECT [dbo].FUN_H7LIS_PRE_RESULT4('" & strPtId & "', '" & Format(Now, "yyyymmdd") & "', '" & Format(Now, "hhmm") & "', 'C2302N6')"
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                strPreResult = Trim(RS.Fields(0)) & ""
            End If
            
            If strPreResult = "" Then
                '���� ó�泪�� �ڵ忩�� ��ã�� ���..(����)
                SQL = ""
                SQL = SQL & "SELECT RESULT_NM AS RESULT" & vbCr
                SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
                SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   AND ORD_CD = 'C2302N6'"

                '-- Record Count ������
                AdoCn.CursorLocation = adUseClient
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    strPreResult = Trim(RS.Fields("RESULT")) & ""
                End If

                RS.Close
                
                If strPreResult = "" Then
                    '�������� ��ã�� ���..(����)
                    SQL = ""
                    SQL = SQL & "SELECT RESULT " & vbCr
                    SQL = SQL & "  FROM PATRESULT " & vbCr
                    SQL = SQL & " WHERE BARCODE = '" & strBarcode & "'" & vbCr
                    SQL = SQL & "   AND EXAMCODE = 'C2302N6'"
    
                    '-- Record Count ������
                    AdoCn_Local.CursorLocation = adUseClient
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        strPreResult = Trim(RS_L.Fields("RESULT")) & ""
                    End If
    
                    RS_L.Close
                End If
                
                
                If strPreResult = "" Then
                    Exit Function
                Else
                    If strResult <> "" Then
                        If IsNumeric(strResult) And CCur(strResult) > 0 And IsNumeric(strPreResult) And CCur(strPreResult) > 0 Then
                            strCalResult = (strPreResult / strResult) * 1000
                            strCalResult = Format(strCalResult, "###0.0")
                        Else
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                End If
            Else
                If IsNumeric(strResult) And CCur(strResult) > 0 And IsNumeric(strPreResult) And CCur(strPreResult) > 0 Then
                    strCalResult = (strPreResult / strResult) * 1000
                    strCalResult = Format(strCalResult, "###0.0")
                Else
                    Exit Function
                End If
            End If
        
        Case "C2302N6"       'M.alb �̸�  Crea(��ȸ��)����� �����´�.
            SQL = "SELECT [dbo].FUN_H7LIS_PRE_RESULT4('" & strPtId & "', '" & Format(Now, "yyyymmdd") & "', '" & Format(Now, "hhmm") & "', 'C3750N3')"
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                strPreResult = Trim(RS.Fields(0)) & ""
            End If
            If strPreResult = "" Then
                '���� ó�泪�� �ڵ忩�� ��ã�� ���..(����)
                SQL = ""
                SQL = SQL & "SELECT RESULT_NM AS RESULT" & vbCr
                SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
                SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   AND ORD_CD = 'C3750N3'"

                '-- Record Count ������
                AdoCn.CursorLocation = adUseClient
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    strPreResult = Trim(RS.Fields("RESULT")) & ""
                End If

                RS.Close
                
                If strPreResult = "" Then
                    '�������� ��ã�� ���..(����)
                    SQL = ""
                    SQL = SQL & "SELECT RESULT " & vbCr
                    SQL = SQL & "  FROM PATRESULT " & vbCr
                    SQL = SQL & " WHERE BARCODE = '" & strBarcode & "'" & vbCr
                    SQL = SQL & "   AND EXAMCODE = 'C3750N3'"
    
                    '-- Record Count ������
                    AdoCn_Local.CursorLocation = adUseClient
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        strPreResult = Trim(RS_L.Fields("RESULT")) & ""
                    End If
    
                    RS_L.Close
                End If
                
                If strPreResult = "" Then
                    Exit Function
                Else
                    If IsNumeric(strResult) And CCur(strResult) > 0 And IsNumeric(strPreResult) And CCur(strPreResult) > 0 Then
                        strCalResult = (strResult / strPreResult) * 1000
                        strCalResult = Format(strCalResult, "###0.0")
                    Else
                        Exit Function
                    End If
                End If
            
            Else
                If IsNumeric(strResult) And CCur(strResult) > 0 And IsNumeric(strPreResult) And CCur(strPreResult) > 0 Then
                    strCalResult = (strResult / strPreResult) * 1000
                    strCalResult = Format(strCalResult, "###0.0")
                Else
                    Exit Function
                End If
            End If
        
        
    End Select
    
    If strCalResult <> "" Then
        SQL = ""
        SQL = SQL & "SELECT RSLTCHANNEL,TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
        SQL = SQL & "  FROM EQPMASTER" & vbCr
        SQL = SQL & " WHERE TESTCODE = '" & strCalTestCd & "'"
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            strIntBase = Trim(RS_L.Fields("RSLTCHANNEL")) & ""
            lsTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
            lsTestName = Trim(RS_L.Fields("TESTNAME")) & ""
            lsSeqNo = Trim(RS_L.Fields("SEQNO")) & ""

            '-- ���Row �߰�
            lsRstRow = SPDRST.DataRowCnt + 1
            If SPDRST.MaxRows < lsRstRow Then
                SPDRST.MaxRows = lsRstRow
            End If
    
            '����� ǥ��
            For intCol = colSTATE + 1 To SPDORD.MaxCols
                If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                    SetText SPDORD, strResult, ActiveRow, intCol
                    Exit For
                End If
            Next
    
            '-- ��� List
            SetText SPDRST, lsSeqNo, lsRstRow, colRSEQNO                '����
            SetText SPDRST, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
            SetText SPDRST, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
            SetText SPDRST, lsTestName, lsRstRow, colRTESTNM            '�˻��
            SetText SPDRST, strIntBase, lsRstRow, colRCHANNEL           '���ä��
            SetText SPDRST, strCalResult, lsRstRow, colRMACHRESULT      '�����
            SetText SPDRST, strCalResult, lsRstRow, colRLISRESULT       'LIS���
            SetText SPDRST, "", lsRstRow, colRJUDGE                     '����
            SetText SPDRST, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
            
            '-- ���� ����
            If pTV = "" Then
                SetLocalDB ActiveRow, lsRstRow, "1", ""
            Else
                SetLocalDB_R ActiveRow, lsRstRow, "1", ""
            End If
            
            '-- ���Count
            If GetText(SPDORD, ActiveRow, colRCNT) = "" Then
                SetText SPDORD, "1", ActiveRow, colRCNT
            Else
                SetText SPDORD, GetText(SPDORD, ActiveRow, colRCNT) + 1, ActiveRow, colRCNT
            End If
            
            
            If gHOSP.MACHCD = "B05" Then
                If pTV = "" Then
                    Res = SaveTransData_MCC_VERSACELL(ActiveRow)
                Else
                    Res = SaveTransData_MCC_VERSACELL_R(ActiveRow)
                End If
            Else
                '1800
                If pTV = "" Then
                    Res = SaveTransData_MCC(ActiveRow)
                Else
                    Res = SaveTransData_MCC_R(ActiveRow)
                End If
            End If
            
            If Res = -1 Then
                '-- ���� ����
                SetForeColor SPDORD, ActiveRow, ActiveRow, 1, colSTATE, 255, 0, 0
                SetText SPDORD, "Failed", ActiveRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor SPDORD, ActiveRow, ActiveRow, 1, colSTATE, 202, 255, 112
                SetText SPDORD, "����Ϸ�", ActiveRow, colSTATE
                SetText SPDORD, "0", ActiveRow, colCHECKBOX
                
                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(SPDORD, ActiveRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(SPDORD, ActiveRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(SPDORD, ActiveRow, colSAVESEQ)) & vbCrLf
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
        End If
    End If
    
End Function


'-- ���� �˻��� ��¥�� Max + 1 ��ȣ�� �����´�
Public Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
          SQL = "SELECT MAX(SAVESEQ) as SEQ FROM PATRESULT  "
    SQL = SQL & " WHERE MID(EXAMDATE,1,8) = '" & strDate & "' " & vbCrLf
    
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

