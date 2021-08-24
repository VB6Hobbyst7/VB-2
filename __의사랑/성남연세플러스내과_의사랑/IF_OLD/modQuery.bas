Attribute VB_Name = "modQuery"
Option Explicit

Public SQL          As String
Public RS           As ADODB.Recordset

Dim blnSameRecord As Boolean

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
    Dim RS2             As ADODB.Recordset
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
    
    'Call SetSQLData("ITEM��ȸ2", SQL)
    
    Set RS2 = New ADODB.Recordset
    
    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS2 = AdoCn_Local.Execute(SQL, , 1)
    If Not RS2.EOF = True And Not RS2.BOF = True Then
        Do Until RS2.EOF
            GetTestNm = RS2.Fields("ITEMNM").Value & ""
            RS2.MoveNext
        Loop
    End If
    
    RS2.Close
    
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
    frmErrMsg.Show vbModal
    
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
    frmErrMsg.Show vbModal
    
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
    frmErrMsg.Show vbModal
    
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
    frmErrMsg.Show vbModal
    
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
    frmErrMsg.Show vbModal
    
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
    frmErrMsg.Show vbModal
    
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
    frmErrMsg.Show vbModal
    
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
    frmErrMsg.Show vbModal
    
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
    frmErrMsg.Show vbModal
    
    Screen.MousePointer = 0
    
End Function


'-- ��ũ����Ʈ ��ȸ
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String)
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
    Dim strDateFr8  As String
    Dim strDateTo8  As String
    Dim strDateFr10 As String
    Dim strDateTo10 As String
    Dim sqlRet      As Integer
    
    Dim varXML      As Variant
    Dim strCommDate As String
    Dim strBarno    As String
    Dim intCnt      As Integer
    Dim varTmp      As Variant
    Dim pGrid_Point As Integer
    Dim pEqipType   As String
    Dim strBarNum   As String
    Dim strPJumin   As String
    Dim strInsNo    As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    Select Case gOCS
            
        Case "KOMAIN"
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = New ADODB.Recordset
            'EMRLIS2
            'r.BCID, r.Hcode, r.Serial, c.PtName, r.Orderdate, ErYn
            SQL = "Exec AP_INF_Bar_Order '" & gHOSP.MACHCD & "','" & pFrom & "','" & pTo & "'"
            
            RS.Open AdoCn.Execute("Exec AP_INF_Bar_Order '" & gHOSP.MACHCD & "','" & pFrom & "','" & pTo & "'", sqlRet)
            
            Call SetSQLData("��ũ��ȸ", SQL)
            
            If Not RS.EOF = True And Not RS.BOF = True Then
                strItems = ""
                Do Until RS.EOF
                    frmWorkList.txtQuery.Text = SQL
                    iCnt = iCnt + 1
                    With frmWorkList.spdWork
                        .ReDraw = False
                        
                        For i = 1 To frmWorkList.spdWork.DataRowCnt
                            strDate = GetText(frmWorkList.spdWork, i, colHOSPDATE)
                            strBarcode = GetText(frmWorkList.spdWork, i, colBARCODE)
                            If Trim(RS.Fields("ORDERDATE")) = strDate And Trim(RS.Fields("BCID")) = strBarcode Then
                                blnSame = True
                            End If
                        Next
                        
                        If blnSame = False Then
                            .MaxRows = .MaxRows + 1
                            intRow = .MaxRows
                                
                            SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                            SetText frmWorkList.spdWork, Trim(RS.Fields("ORDERDATE")) & "", intRow, colHOSPDATE
                            SetText frmWorkList.spdWork, Trim(RS.Fields("BCID")) & "", intRow, colBARCODE
                            SetText frmWorkList.spdWork, Trim(RS.Fields("Hcode")) & "", intRow, colPID
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PtName")) & "", intRow, colPNAME
                                                        
                            '�˻��׸� ��ȸ
                            SetText frmWorkList.spdWork, frmWorkList.txtSeq.Text, intRow, colSEQNO
                            'SetText frmWorkList.spdWork, RS.Fields("CNT"), intRow, colOCNT
                            
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
            
         Case "EASYS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & " a.ACC_YMD     AS HOSPDATE"
            SQL = SQL & ", a.RECEPT_NO  AS BARCODE"
            SQL = SQL & ", a.PTNT_NO    AS PID"
            SQL = SQL & " ,c.PTNT_NM    AS PNAME"
            'SQL = SQL & " ,a.AGE AS AGE"
            'SQL = SQL & " ,a.SEX AS SEX"
            'SQL = SQL & " ,a.IO_GB AS INOUT "
            SQL = SQL & ", COUNT(a.ORD_CD) AS CNT " & vbCr
            SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c " & vbCr
            SQL = SQL & " WHERE a.ACC_YMD between '" & pFrom & "' AND '" & pTo & "'" & vbCr
            SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND a.STS_CD    = 'A'" & vbCr                                                'A:����, R:�������
            SQL = SQL & "   AND a.SUTAK_CD  = ''" & vbCr
            SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO " & vbCr
            SQL = SQL & "   AND a.ptnt_no    = c.ptnt_no " & vbCr
            SQL = SQL & " GROUP BY a.ACC_YMD, a.RECEPT_NO, a.PTNT_NO, c.PTNT_NM " & vbCr
            SQL = SQL & " ORDER BY a.ACC_YMD, a.PTNT_NO "
           
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
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                            'SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                            'SetText frmWorkList.spdWork, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                            'SetText frmWorkList.spdWork, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "�Կ�", "�ܷ�"), intRow, colINOUT
                            SetText frmWorkList.spdWork, frmWorkList.txtSeq.Text, intRow, colSEQNO
                            SetText frmWorkList.spdWork, RS.Fields("CNT"), intRow, colOCNT
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
            

         Case "UBCARE"
'            SQL = ""
'            SQL = SQL & "SELECT DISTINCT "
'            SQL = SQL & " a.ACC_YMD     AS HOSPDATE"
'            SQL = SQL & ", a.RECEPT_NO  AS BARCODE"
'            SQL = SQL & ", a.PTNT_NO    AS PID"
'            SQL = SQL & " ,c.PTNT_NM    AS PNAME"
'            'SQL = SQL & " ,a.AGE AS AGE"
'            'SQL = SQL & " ,a.SEX AS SEX"
'            'SQL = SQL & " ,a.IO_GB AS INOUT "
'            SQL = SQL & ", COUNT(a.ORD_CD) AS CNT " & vbCr
'            SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c " & vbCr
'            SQL = SQL & " WHERE a.ACC_YMD between '" & pFrom & "' AND '" & pTo & "'" & vbCr
'            SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ") " & vbCr
'            SQL = SQL & "   AND a.STS_CD    = 'A'" & vbCr                                                'A:����, R:�������
'            SQL = SQL & "   AND a.SUTAK_CD  = ''" & vbCr
'            SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO " & vbCr
'            SQL = SQL & "   AND a.ptnt_no    = c.ptnt_no " & vbCr
'            SQL = SQL & " GROUP BY a.ACC_YMD, a.RECEPT_NO, a.PTNT_NO, c.PTNT_NM " & vbCr
'            SQL = SQL & " ORDER BY a.ACC_YMD, a.PTNT_NO "
            pEqipType = "C"
            
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "  commdate AS HOSPDATE"
            SQL = SQL & ", barcode AS BARCODE"
            SQL = SQL & ", chartno as PID"
            SQL = SQL & ", patname AS PNAME"
            SQL = SQL & ", patsex AS SEX"
            SQL = SQL & ", patage AS AGE"
            SQL = SQL & ", remark "
            SQL = SQL & "  FROM PAT_RES "
            SQL = SQL & " WHERE commdate between '" & pFrom & "' AND '" & pTo & "'"
            SQL = SQL & "   AND EXAMID IN (" & gAllTestCd & ") " & vbCr
            With frmWorkList
                If .chkPart(1).Value = "1" Then
                    SQL = SQL & "  and (mid(chartno,1,1) <> 'G' and mid(chartno,1,1) <> 'C') "
                ElseIf .chkPart(2).Value = "1" Then
                    SQL = SQL & "  and mid(chartno,1,1) = 'G' "
                ElseIf .chkPart(3).Value = "1" Then
                    SQL = SQL & "  and mid(chartno,1,1) = 'C' "
                End If
            
                If .chkSave.Value = "0" Then
                    SQL = SQL & "   and result = '' "
                End If
            End With
            
            SQL = SQL & " Order by commdate,remark "
            
            
            Call SetSQLData("��ũ��ȸ", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count ������
            AdoCn_Local.CursorLocation = adUseClient
            Set RS = AdoCn_Local.Execute(SQL, , 1)
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
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                            'SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                            'SetText frmWorkList.spdWork, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                            'SetText frmWorkList.spdWork, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "�Կ�", "�ܷ�"), intRow, colINOUT
                            SetText frmWorkList.spdWork, frmWorkList.txtSeq.Text, intRow, colSEQNO
                            'SetText frmWorkList.spdWork, RS.Fields("CNT"), intRow, colOCNT
                            SetText frmWorkList.spdWork, GetSampleITEM(intRow), intRow, colITEMS
                            
                            frmWorkList.txtSeq.Text = frmWorkList.txtSeq.Text + 1
                        
                        End If
                    End With
                    
                    blnSame = False
                
                    DoEvents
                    
                    RS.MoveNext
                Loop
                
                RS.Close
                
                '-- XML �ϱ�
                varXML = f_subSet_XMLWorkList(pFrom, pTo)
                
                If blnSameRecord = False Then
                    'MsgBox "�˻� ����ڰ� �����ϴ�.", vbOKOnly + vbInformation, App.Title
                    Exit Sub
                End If
                
                If UBound(varXML) < 1 Then
                    'MsgBox "�˻� ����ڰ� �����ϴ�.", vbOKOnly + vbInformation, App.Title
                    Exit Sub
                Else
                    strBarno = ""
            
                    With frmWorkList.spdWork
                        For intCnt = 0 To UBound(varXML) - 1
                            varTmp = Split(varXML(intCnt), ",")
                                            
                            '-- ���ä�ΰ�ã��
                            SQL = ""
                            SQL = SQL & " SELECT RSLTCHANNEL,TESTNAME "
                            SQL = SQL & "   FROM EQPMASTER"
                            SQL = SQL & "  WHERE TESTCODE = '" & Trim(varTmp(8)) & "' "
                                                        
                            Set RS = AdoCn_Local.Execute(SQL, , 1)
                            
                            If Not RS.EOF = True And Not RS.BOF = True Then
                                '-- ���� ���� ���
                                XMLInData.ComExamID = Trim(RS.Fields("RSLTCHANNEL").Value & "")
                            
                                XMLInData.Company = varTmp(0)
                                XMLInData.HospCode = varTmp(1)
                                XMLInData.ChartNo = varTmp(2)
                                XMLInData.PatName = varTmp(3)
                                XMLInData.PatJumin = varTmp(4)
                                XMLInData.PatNo = varTmp(5)
                                XMLInData.CommDate = varTmp(6)
                                XMLInData.ExamNo = varTmp(7)
                                XMLInData.ExamID = varTmp(8)
                                XMLInData.Specimen = varTmp(10)
                                XMLInData.Result = varTmp(11)
                                XMLInData.Reference = varTmp(12)
                                XMLInData.Remark = varTmp(13)
                                XMLInData.RsltDate = varTmp(14)
                                XMLInData.IOFlag = varTmp(15)
                                
                                
'''                                SQL = ""
'''                                SQL = SQL & "select equipno, equipcode, examname, examtype "
'''                                SQL = SQL & "  from equipexam "
'''                                SQL = SQL & " where examcode = '" & XMLInData.ExamID & "' "
'''                                Res = db_select_Col(gLocal, SQL)
'''                                If Res > 0 Then
'                                    PEquipno = gReadBuf(0)
'                                    PEquipCode = gReadBuf(1)
'                                    PExamname = gReadBuf(2)

                                    If strBarno <> XMLInData.ChartNo And strCommDate <> XMLInData.CommDate Then
                                        pEqipType = "C"

                                        pGrid_Point = SeqSearch_New(frmWorkList.spdWork, XMLInData.ChartNo, pEqipType, colPID)

                                        If pGrid_Point = 0 Then
                                            pGrid_Point = SeqNullSearch(frmWorkList.spdWork, XMLInData.ChartNo, colPID)
                                            If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
                                            .RowHeight(-1) = 12
                                        End If

                                        .SetText colCHECKBOX, pGrid_Point, "1"
                                        .SetText colHOSPDATE, pGrid_Point, XMLInData.CommDate
                                        '.SetText colBARCODE, pGrid_Point, pEqipType
                                        strBarNum = Mid(XMLInData.CommDate, 5, 4) & Format(XMLInData.ChartNo, "0000000000")
                                        .SetText colBARCODE, pGrid_Point, strBarNum
                                        .SetText colPID, pGrid_Point, XMLInData.ChartNo
                                        .SetText colPNAME, pGrid_Point, XMLInData.PatName
                                                    strPJumin = Left(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
                                                    Call CalAgeSex(strPJumin, Format(Date, "yyyy/mm/dd"))
                                        .SetText colPSEX, pGrid_Point, gPatGen.Sex
                                        .SetText colPAGE, pGrid_Point, gPatGen.Age
                                        .SetText colSTATE, pGrid_Point, "Order"
                                        
                                        
                                        strInsNo = getMaxTestNum(XMLInData.CommDate)

                                    End If
                                          SQL = "Select ChartNo from pat_res " & vbCr
                                    SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' " & vbCr
                                    SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' " & vbCr
                                    SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'" & vbCr
                                    SQL = SQL & "   and BarCode  = '" & strBarNum & "'" & vbCr
                                    SQL = SQL & "   and ExamType = '" & pEqipType & "'" & vbCr

                                    Set RS = AdoCn_Local.Execute(SQL, , 1)
                                    
                                    If RS.EOF = True And RS.BOF = True Then
                                              SQL = " insert into pat_res("
                                        SQL = SQL & "Company,HospCode,ChartNo, "
                                        SQL = SQL & "PatName,PatSex,PatAge,PatJumin,PatNo,"
                                        SQL = SQL & "CommDate,ExamNo,ExamID,ComExamID, "
                                        SQL = SQL & "Specimen,Result,Reference,Remark,RsltDate,IOFlag,BarCode,ExamType)"
                                        SQL = SQL & " values ("
                                        SQL = SQL & "'" & XMLInData.Company & "',"
                                        SQL = SQL & "'" & XMLInData.HospCode & "',"
                                        SQL = SQL & "'" & XMLInData.ChartNo & "',"
                                        SQL = SQL & "'" & XMLInData.PatName & "',"
                                        SQL = SQL & "'" & gPatGen.Sex & "',"
                                        SQL = SQL & "'" & gPatGen.Age & "',"
                                        SQL = SQL & "'" & XMLInData.PatJumin & "',"
                                        SQL = SQL & "'" & XMLInData.PatNo & "',"
                                        SQL = SQL & "'" & XMLInData.CommDate & "',"
                                        SQL = SQL & "'" & XMLInData.ExamNo & "',"
                                        SQL = SQL & "'" & XMLInData.ExamID & "',"
                                        SQL = SQL & "'" & XMLInData.ComExamID & "',"
                                        SQL = SQL & "'" & XMLInData.Specimen & "',"
                                        SQL = SQL & "'" & XMLInData.Result & "',"
                                        SQL = SQL & "'" & XMLInData.Reference & "',"
                                        'SQL = SQL & "'" & XMLInData.Remark & "',"
                                        SQL = SQL & "'" & strInsNo & "',"
                                        SQL = SQL & "'" & XMLInData.RsltDate & "',"
                                        SQL = SQL & "'" & XMLInData.IOFlag & "',"
                                        SQL = SQL & "'" & strBarNum & "',"
                                        SQL = SQL & "'" & pEqipType & "')"

                                        If Not DBExec(AdoCn_Local, SQL) Then
                                            'SaveQuery SQL
                                            'Exit Function
                                        End If
'
'                                    '-- �ӵ������ ���� ������ �����
                                    Else
                                              SQL = " Update pat_res Set "
                                        SQL = SQL & " PatName = '" & XMLInData.PatName & "', "
                                        SQL = SQL & " PatSex  = '" & gPatGen.Sex & "' "
                                        SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
                                        SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
                                        SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
                                        SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
                                        SQL = SQL & "   and ExamType = '" & pEqipType & "'"

                                        If Not DBExec(AdoCn_Local, SQL) Then
                                            'SaveQuery SQL
                                            'Exit Function
                                        End If
                                    End If

                                    strBarno = XMLInData.ChartNo
                                    strCommDate = XMLInData.CommDate
                                    .SetText colITEMS, pGrid_Point, GetSampleITEM(pGrid_Point)

'''                                End If
                            
                            End If
'                            Res = GetDBSelectColumn(gLocal, SQL)
                            XMLInData.ComExamID = ""
                            
                            '-- ���� ���� ���
'                            If Res > 0 Then
'
'
'                            End If
                            
                            XMLInData.ComExamID = ""
                        Next
                        
                    End With
                End If
                
                frmWorkList.chkAll.Value = "1"
            Else
                frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
                        
    End Select

     
    frmWorkList.spdWork.RowHeight(-1) = 12
    frmWorkList.spdWork.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal
    
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
                        If RS.Fields("REFJUDGE").Value & "" = "H" Or RS.Fields("REFJUDGE").Value & "" = "L" Then
                            frmMain.spdROrder.ForeColor = vbRed
                        Else
                            frmMain.spdROrder.ForeColor = vbBlack
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
    Dim RS1             As ADODB.Recordset
    Dim strBarcode      As String
    Dim strRegDate      As String
    Dim lngExamNo       As Long
    Dim strItems        As String
    Dim sqlRet          As Integer
    
    GetSampleITEM = ""
    
    strRegDate = Trim(GetText(frmWorkList.spdWork, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(frmWorkList.spdWork, asRow, colBARCODE))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Select Case gOCS

        Case "KOMAIN"
            SQL = "Exec AP_INF_Bar_Order_Coda '" & gHOSP.MACHCD & "', '" & strBarcode & "'"
            
            'Call SetSQLData("ITEM��ȸ", SQL)
            
            Set RS1 = New ADODB.Recordset
            RS1.Open AdoCn.Execute("Exec AP_INF_Bar_Order_Coda '" & gHOSP.MACHCD & "', '" & strBarcode & "'", sqlRet)
            If Not RS1.EOF = True And Not RS1.BOF = True Then
                Do Until RS1.EOF
                    With frmMain.spdOrder
                        .ReDraw = False
                        If strItems = "" Then
                            strItems = GetTestNm(Trim(RS1.Fields("Coda")) & "/" & Trim(RS1.Fields("SubCoda")), False)
                        Else
                            strItems = strItems & "," & GetTestNm(Trim(RS1.Fields("Coda")) & "/" & Trim(RS1.Fields("SubCoda")), False)
                        End If
                    End With
                    DoEvents
                        
                    RS1.MoveNext
                Loop
            End If
            
            GetSampleITEM = strItems
            
            RS1.Close
            
            Exit Function
         
         Case "EASYS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM " & vbCr
            SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c " & vbCr
            SQL = SQL & " WHERE a.ACC_YMD  = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   AND a.RECEPT_NO = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND a.STS_CD    = 'A'" & vbCr                                                'A:����, R:�������
            SQL = SQL & "   AND a.SUTAK_CD  = ''" & vbCr
            SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO " & vbCr
'            SQL = SQL & "   AND a.STS_CD    = b.STS_CD " & vbCr
            SQL = SQL & " ORDER BY ORD_CD "
                            
        Case "UBCARE"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT EXAMID AS ITEM " & vbCr
            SQL = SQL & "  FROM PAT_RES "
            SQL = SQL & " WHERE commdate = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   AND BARCODE = '" & strBarcode & "'" & vbCr
            SQL = SQL & " Order by EXAMID "
            Call SetSQLData("ITEM��ȸ", SQL)
            
            '-- Record Count ������
            AdoCn_Local.CursorLocation = adUseClient
            Set RS1 = AdoCn_Local.Execute(SQL, , 1)
            If Not RS1.EOF = True And Not RS1.BOF = True Then
                Do Until RS1.EOF
                    'With frmMain.spdOrder
                    '    .ReDraw = False
                        If strItems = "" Then
                            strItems = GetTestNm(Trim(RS1.Fields("ITEM")) & "", False)
                        Else
                            strItems = strItems & "/" & GetTestNm(Trim(RS1.Fields("ITEM")), False)
                        End If
                        
                    'End With
                    'DoEvents
                    
                    RS1.MoveNext
                Loop
            End If
            
            GetSampleITEM = strItems
            
            RS1.Close
                        
            Exit Function
            
    End Select
    
    Call SetSQLData("ITEM��ȸ", SQL)
    
    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS1 = AdoCn.Execute(SQL, , 1)
    If Not RS1.EOF = True And Not RS1.BOF = True Then
        Do Until RS1.EOF
            With frmMain.spdOrder
                .ReDraw = False
                If strItems = "" Then
                    strItems = GetTestNm(Trim(RS1.Fields("ITEM")) & "", False)
                Else
                    strItems = strItems & "/" & GetTestNm(Trim(RS1.Fields("ITEM")), False)
                End If
                
            End With
            DoEvents
            
            RS1.MoveNext
        Loop
    End If
    
    GetSampleITEM = strItems
    
    RS1.Close
    
End Function

Private Function f_subSet_XMLWorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String) As Variant
    Dim strPath   As String
    Dim strBuffer As String
    Dim i         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIdx As Integer
    
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    
    '-- �������ϸ�� ��θ� �����Ѵ�.
    strPath = "C:\UBCare\SINAI\IF\ExamIF_In.xml"

    
    '1���ξ� �������� MSDN����
    Dim TextLine
    Open strPath For Input As #1 ' ������ ���ϴ�.
    
    Do While Not EOF(1) ' ������ ���� ���� ������ �ݺ��մϴ�.
        Line Input #1, TextLine ' ������ ������ ���� �о���Դϴ�.
        strBuffer = strBuffer & TextLine
    Loop
    
    Close #1 ' ������ �ݽ��ϴ�
 
    intIdx = 0
    lngBufLen = Len(strBuffer)
        
    For i = 1 To lngBufLen
        If intIdx = 0 Then
            BufChar = Mid$(strBuffer, i, 4)
        Else
            BufChar = Mid$(strBuffer, i + 3)
        End If
        
        If BufChar = "<�˻�>" Then
            intIdx = 1
            strTmp = BufChar
        Else
            strTmp = strTmp & BufChar
            If intIdx = 1 Then Exit For
        End If
    
    Next
    
'    f_subSet_XMLWorkList = Split(strTmp, "</�˻�>")
    strTmp = Replace(strTmp, "<�˻�>", ""): strTmp = Replace(strTmp, "</�˻�>", "|")
    strTmp = Replace(strTmp, "<��ü>", ""): strTmp = Replace(strTmp, "</��ü>", ",")
    strTmp = Replace(strTmp, "<�������ȣ>", ""): strTmp = Replace(strTmp, "</�������ȣ>", ",")
    strTmp = Replace(strTmp, "<��Ʈ��ȣ>", ""): strTmp = Replace(strTmp, "</��Ʈ��ȣ>", ",")
    strTmp = Replace(strTmp, "<�����ڸ�>", ""): strTmp = Replace(strTmp, "</�����ڸ�>", ",")
    strTmp = Replace(strTmp, "<�ֹε�Ϲ�ȣ>", ""): strTmp = Replace(strTmp, "</�ֹε�Ϲ�ȣ>", ",")
    strTmp = Replace(strTmp, "<������ȣ>", ""): strTmp = Replace(strTmp, "</������ȣ>", ",")
    strTmp = Replace(strTmp, "<�Ƿ���>", ""): strTmp = Replace(strTmp, "</�Ƿ���>", ",")
    strTmp = Replace(strTmp, "<�˻��ȣ>", ""): strTmp = Replace(strTmp, "</�˻��ȣ>", ",")
    strTmp = Replace(strTmp, "<�˻�ID>", ""): strTmp = Replace(strTmp, "</�˻�ID>", ",")
    strTmp = Replace(strTmp, "<��ü�˻�ID>", ""): strTmp = Replace(strTmp, "</��ü�˻�ID>", ",")
    strTmp = Replace(strTmp, "<��ü>", ""): strTmp = Replace(strTmp, "</��ü>", ",")
    strTmp = Replace(strTmp, "<���ġ>", ""): strTmp = Replace(strTmp, "</���ġ>", ",")
    strTmp = Replace(strTmp, "<����ġ>", ""): strTmp = Replace(strTmp, "</����ġ>", ",")
    strTmp = Replace(strTmp, "<�Ұ�>", ""): strTmp = Replace(strTmp, "</�Ұ�>", ",")
    strTmp = Replace(strTmp, "<�����>", ""): strTmp = Replace(strTmp, "</�����>", ",")
    strTmp = Replace(strTmp, "<��ü>", ""): strTmp = Replace(strTmp, "</��ü>", ",")
    strTmp = Replace(strTmp, "<�Կ��ܷ�����>", ""): strTmp = Replace(strTmp, "</�Կ��ܷ�����>", ",")
    
    f_subSet_XMLWorkList = Split(strTmp, "|")
    blnSameRecord = True
    
    'Kill strPath
    
    Screen.MousePointer = 0

    
    Exit Function
        
ErrorTrap:
    
    blnSameRecord = False
    Screen.MousePointer = 0
    
    
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
    Dim strBarcode      As String
    Dim strRegDate      As String
    Dim lngExamNo       As Long
    Dim intTestCnt      As Integer
    
    Dim intCol     As Integer
    
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    Dim sqlRet  As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    
    '-- üũ�ڽ� ��üũ
    SetText SPD, "0", asRow, colCHECKBOX
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    Select Case gOCS
        Case "KOMAIN"
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = New ADODB.Recordset
            SQL = "Exec AP_INF_Bar_Order_Coda '" & gHOSP.MACHCD & "', '" & strBarcode & "'"
            
            RS.Open AdoCn.Execute("Exec AP_INF_Bar_Order_Coda '" & gHOSP.MACHCD & "', '" & strBarcode & "'", sqlRet)
            
            Call SetSQLData("ȯ����ȸ", SQL)
            
            '-- 2017.09.05
            SetText SPD, "0", asRow, colCHECKBOX
            
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    With SPD
                        .ReDraw = False
                        intTestCnt = intTestCnt + 1
                        SetText SPD, "1", asRow, colCHECKBOX
                        SetText SPD, Trim(RS.Fields("ORDERDATE")) & "", asRow, colHOSPDATE
                        SetText SPD, Trim(RS.Fields("BCID")), asRow, colBARCODE
                        SetText SPD, Trim(RS.Fields("Hcode")) & "", asRow, colPID
                        mOrder.PID = Trim(RS.Fields("Hcode")) & ""
                        SetText SPD, Trim(RS.Fields("PtName")) & "", asRow, colPNAME
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                        
                        '-- ȭ�鿡 ǥ��
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("Coda")) & "/" & Trim(RS.Fields("subCoda")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "��", asRow, intCol)
                                Exit For
                            End If
                        Next
                        gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("Coda")) & "/" & Trim(RS.Fields("subCoda")) & "',"
                        
                    End With
                    DoEvents
                    
                    RS.MoveNext
                Loop
            End If
            
            RS.Close
            
        Case "EASYS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & " a.ACC_YMD     AS HOSPDATE"
            SQL = SQL & ", a.RECEPT_NO  AS BARCODE"
            SQL = SQL & ", a.PTNT_NO    AS PID"
            SQL = SQL & ", c.PTNT_NM    AS PNAME"
            SQL = SQL & ", a.ORD_CD AS ITEM " & vbCr
            SQL = SQL & " ,a.SPC_CD AS SPCCD "
            SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c " & vbCr
            SQL = SQL & " WHERE a.RECEPT_NO = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND a.STS_CD    = 'A'" & vbCr                                                'A:����, R:�������
            SQL = SQL & "   AND a.SUTAK_CD  = ''" & vbCr
            SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO " & vbCr
            SQL = SQL & "   AND a.ptnt_no    = c.ptnt_no " & vbCr
            SQL = SQL & " ORDER BY a.ORD_CD "

            
            Call SetSQLData("���ڵ���ȸ", SQL)
            
            '-- Record Count ������
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    With SPD
                        .ReDraw = False
                        intTestCnt = intTestCnt + 1
                        SetText SPD, "1", asRow, colCHECKBOX
                        SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                        SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                        SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                        
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        
                        
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


Public Function getEASYSJudge(ByVal pOrdCD As String, ByVal pResult As String) As String
    Dim strLow      As String
    Dim strHigh     As String
    
    getEASYSJudge = ""
    
          SQL = "Select REFLOW, REFHIGH  "
    SQL = SQL & "  From EQPMASTER"
    SQL = SQL & " Where EQUIPCD = '" & gHOSP.MACHCD & "' "
    SQL = SQL & "   And TESTCODE =  '" & pOrdCD & "'"
    
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strLow = Trim(RS.Fields("REFLOW") & "")
        strHigh = Trim(RS.Fields("REFHIGH") & "")
        
        If strLow <> "" And strHigh <> "" And pResult <> "" And IsNumeric(strLow) And IsNumeric(strHigh) And IsNumeric(pResult) Then
            If Val(pResult) > Val(strHigh) Then
                getEASYSJudge = "H"
            ElseIf Val(pResult) < Val(strLow) Then
                getEASYSJudge = "L"
            Else
                getEASYSJudge = " "
            End If
        Else
            getEASYSJudge = " "
        End If
    Else
        getEASYSJudge = ""
    End If
        
    RS.Close
    
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
    Dim strPTID         As String
    
    If pTV = "" Then
        ActiveRow = SPDRST.ActiveRow
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
        'Case "C3750N1" 'Crea
        '        strCalTestCd = "C3750N1"   'Crea(24�ð���)
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
        strPTID = Trim(RS.Fields("PID")) & ""
    End If
    
    RS.Close
    
    Select Case pTestCd
        Case "URR"    '��Ұ��ҿ�
            'ȯ�ڹ�ȣ,������,���ؽð�,ó���ڵ��Դϴ�.
            '���������[C3730N2]�� �����´�.
            SQL = "SELECT [dbo].FUN_H7LIS_PRE_RESULT4('" & strPTID & "', '" & Format(Now, "yyyymmdd") & "', '" & Format(Now, "hhmm") & "', 'C3730N2')"
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
        Case "C3750"    'Creatine
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
            
'        Case "C3750N1" 'Crea
'            If IsNumeric(strResult) Then
'                strCalResult = strResult * 10 * CCur(pTV)
'                strCalResult = Format(strCalResult, "#,##0.0")
'            Else
'                strCalResult = ""
'            End If
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
                Res = SaveTransData_MCC(ActiveRow)
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

