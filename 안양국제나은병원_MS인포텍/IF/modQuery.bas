Attribute VB_Name = "modQuery"
Option Explicit

Public SQL          As String
Public RS           As ADODB.Recordset

'-- 검사마스터 조회
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
    
    '-- Record Count 가져옴
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
                '-- 소수점변환으로 사용
                'Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCTemp").Value & "", intRow, colLQCTemp)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("QCTemp").Value & "", intRow, colLUseResSpec)
               ' Call SetText(frmMain.spdTest, IIf(AdoRs_Local.Fields("QCTemp").Value & "" = "1", "사용", "미사용"), intRow, colLUseResSpec)

                
                
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

'-- 검사오더마스터 조회
Public Sub GetOrderMST()
    Dim intRow          As Long
    
    gAllOrdCd = ""
    intRow = 0
    
    SQL = ""
    SQL = SQL & "SELECT ORDERCODE FROM ORDMASTER " & vbCr
    SQL = SQL & " ORDER BY ORDERCODE "
    
    '-- Record Count 가져옴
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

'-- 검사코드로 검사명 조회
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
    
    '-- Record Count 가져옴
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

'-- 검사항목 조회
Public Function GetTest(ByVal pTestCd As String) As String
    
On Error GoTo RST
    GetTest = ""
    
    SQL = ""
    SQL = SQL & "Select ORD_NM "
    SQL = SQL & "  From LIS_ORD_LIST_V" & vbCr
    SQL = SQL & " Where ord_cd = '" & pTestCd & "'" & vbCr
  
    '-- Record Count 가져옴
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
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Function

'-- QC 물질 저장
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
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "_SetQCList_Header" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- QC 물질 저장
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
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "SetQCList_Detail" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- QC 물질 조회
Public Sub GetQCList_Header()
    Dim i   As Integer
    
On Error GoTo RST
    
    SQL = ""
    SQL = SQL & "Select LotID,MachID,InstrumentID "
    SQL = SQL & "  From QCHEADER " & vbCr
  
    '-- Record Count 가져옴
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
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "_GetQCList_Header" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- QC 물질 조회 -상세
Public Sub GetQCList_QCID(ByVal strInst As String)
    Dim i   As Integer
    
On Error GoTo RST
    frmQCMaster.spdQCID.MaxRows = 0
    
    SQL = ""
    SQL = SQL & "Select InstrumentID,QCLevel,ID "
    SQL = SQL & "  From QCDetail " & vbCr
    SQL = SQL & " Where InstrumentID = '" & strInst & "'" & vbCr
    SQL = SQL & " Order By  InstrumentID,QCLevel,ID "
    
    '-- Record Count 가져옴
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
     
                strErrMsg = "위    치 : " & "GetQCList_QCID" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


'-- QC 여부 조회 (덤프만으로 QC 인지 모를때..)
Public Function strQCFlag(ByVal strInst As String, ByVal strQCBarCd As String) As String
    
On Error GoTo RST
    
    strQCFlag = ""

    SQL = ""
    SQL = SQL & "Select Count(*) AS CNT  "
    SQL = SQL & "  From QCDetail " & vbCr
    SQL = SQL & " Where InstrumentID = '" & strInst & "'" & vbCr
    SQL = SQL & "   And ID = '" & strQCBarCd & "'" & vbCr
    
    '-- Record Count 가져옴
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
     
                strErrMsg = "위    치 : " & "strQCFlag" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Function

'-- QC 물질 리스트 조회(서버)
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
    
    '-- Record Count 가져옴
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
     
                strErrMsg = "위    치 : " & "GetQCList_Detail" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- QC 결과 리스트 조회
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
    '-- Record Count 가져옴
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
    
    '-- Record Count 가져옴
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
     
                strErrMsg = "위    치 : " & "GetQCResult_Detail" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Function


'-- QC 결과 리스트 조회
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
    '-- Record Count 가져옴
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
    
    '-- Record Count 가져옴
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
                strQCVal = strQCVal & pEqpCD & "|"          ' lab           // 447834(병원코드로 대체 가능?)
                strQCVal = strQCVal & strLotID & "|"        ' lot           // 159792(입력)
                strQCVal = strQCVal & pAnalyID & "|"        ' analyte       // 검사항목마다 세팅,  Cyfra 21-1 : pAnalyte = "222"
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
     
                strErrMsg = "위    치 : " & "GetQCResult_Detail" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
        'If strAge <= 7 Then
        '    SQL = "Select YMAX as MAX, YMIN as MIN "
        'Else
            If strSex = "M" Then
                     SQL = "Select MMAX as MAX, MMIN as MIN "
            Else
                     SQL = "Select WMAX as MAX, WMIN as MIN "
            End If
        'End If
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

'-- 워크리스트 조회
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
        Case "PHILL"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "  a.request_date AS HOSPDATE"
            SQL = SQL & ", a.exam_no AS PID"
            SQL = SQL & ", a.chart_no AS CHARTNO"
            SQL = SQL & ", a.personal_id AS JUMIN"
            SQL = SQL & ", a.person_name AS PNAME"
            SQL = SQL & ", a.person_sex AS SEX"
            SQL = SQL & ", a.person_age AS AGE"
            SQL = SQL & ", COUNT(b.exam_code) AS CNT " & vbCr
            SQL = SQL & "  FROM TRUST a, TRURES b " & vbCr
            SQL = SQL & " WHERE a.request_date Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
            SQL = SQL & "   AND b.pro_code IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND b.exam_code <> 'X999' " & vbCr
            SQL = SQL & "   AND b.exam_code IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND a.request_date = b.request_date " & vbCr
            SQL = SQL & "   AND a.exam_no = b.exam_no " & vbCr
            SQL = SQL & " GROUP BY a.request_date, a.exam_no, a.chart_no, a.personal_id, a.person_name, a.person_sex, a.person_age" & vbCr
            SQL = SQL & " ORDER BY a.request_date, a.exam_no "
        
            Call SetSQLData("워크조회", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count 가져옴
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                frmWorkList.spdWork.MaxRows = 0
                strItems = ""
                Do Until RS.EOF
                    iCnt = iCnt + 1
                    With frmWorkList.spdWork
                        .ReDraw = False
                        .MaxRows = .MaxRows + 1
                        intRow = .MaxRows
                        strBarcode = Trim(RS.Fields("HOSPDATE")) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0")
                            
                        SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                        SetText frmWorkList.spdWork, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                        SetText frmWorkList.spdWork, strBarcode, intRow, colBARCODE
                        SetText frmWorkList.spdWork, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                        SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
                        SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                        SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                        SetText frmWorkList.spdWork, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                        SetText frmWorkList.spdWork, Trim(RS.Fields("JUMIN")) & "", intRow, colPJUMIN
                        SetText frmWorkList.spdWork, frmWorkList.txtSeq.Text, intRow, colSEQNO
                        SetText frmWorkList.spdWork, RS.Fields("CNT"), intRow, colOCNT
                        SetText frmWorkList.spdWork, GetSampleITEM(intRow), intRow, colITEMS
                        
                        frmWorkList.txtSeq.Text = frmWorkList.txtSeq.Text + 1
                    End With
                    DoEvents
                    
                    RS.MoveNext
                Loop
                frmWorkList.chkAll.Value = "1"
            Else
                frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
        Case "JWINFO"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "  a.RECEIPTDATE AS HOSPDATE"
            SQL = SQL & ", a.SPECIMENNUM AS BARCODE"
            SQL = SQL & ", a.RECEIPTNO AS CHARTNO"
            SQL = SQL & ", a.IPDOPD AS INOUT "
            SQL = SQL & ", a.PTNO AS PID"
            SQL = SQL & ", a.SNAME AS PNAME"
            SQL = SQL & ", COUNT(a.LABCODE) AS CNT " & vbCr
            SQL = SQL & "   FROM SLA_LabMaster a, SLA_LabResult b" & vbCr
            SQL = SQL & " WHERE a.RECEIPTDATE between '" & Format(pFrom, "####-##-##") & "' and '" & Format(pTo, "####-##-##") & "'" & vbCr
            SQL = SQL & "   AND a.ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND b.LABCODE IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND a.RECEIPTNO = b.RECEIPTNO " & vbCr
            SQL = SQL & "   AND a.ORDERCODE = b.ORDERCODE " & vbCr
            SQL = SQL & "   and a.RECEIPTDATE = b.RECEIPTDATE" & vbCr
            SQL = SQL & "   AND a.SPECIMENNUM = b.SPECIMENNUM" & vbCr
            SQL = SQL & "   AND a.JSTATUS < '3'" & vbCr
            SQL = SQL & " GROUP BY a.RECEIPTDATE, a.SPECIMENNUM, a.RECEIPTNO, a.IPDOPD, a.PTNO, a.SNAME" & vbCr
            SQL = SQL & " ORDER BY a.RECEIPTDATE,a.SPECIMENNUM "
                
            Call SetSQLData("워크조회", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count 가져옴
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
                            SetText frmWorkList.spdWork, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
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
                frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
        Case "MCC"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & " READING_YMD AS HOSPDATE"
            SQL = SQL & ", BCODE_NO AS BARCODE"
            'SQL = SQL & ", RECEPT_NO AS CHARTNO"
            SQL = SQL & ", PTNT_NO AS PID"
            SQL = SQL & " ,PTNT_NM AS PNAME"
            SQL = SQL & " ,AGE AS AGE"
            SQL = SQL & " ,SEX AS SEX"
            SQL = SQL & " ,IO_GB AS INOUT "
            SQL = SQL & ", COUNT(ORD_CD) AS CNT " & vbCr
            SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
            SQL = SQL & " WHERE READING_YMD between '" & pFrom & "' AND '" & pTo & "'" & vbCr
            SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND STS_CD = '0'" & vbCr    '0 접수, 1:결과전송
            SQL = SQL & " GROUP BY READING_YMD,BCODE_NO,PTNT_NO,PTNT_NM,AGE,SEX,IO_GB " & vbCr
            SQL = SQL & " ORDER BY READING_YMD,PTNT_NO,BCODE_NO "
            
            Call SetSQLData("워크조회", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count 가져옴
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
'                            SetText frmWorkList.spdWork, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                            SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                            SetText frmWorkList.spdWork, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                            SetText frmWorkList.spdWork, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "입원", "외래"), intRow, colINOUT
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
                frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
        Case "GINUS"
                  SQL = "SELECT /*+ INDEX (coif scccoifm_ix1) INDEX (prex scrprexh_ix3) INDEX (ptbs pmcptbsm_ux1) INDEX (rslt scrrslth_ux1) INDEX (xpsl mosxpslh_ix2) */" & vbCr
            SQL = SQL & " prex.acp_dt AS HOSPDATE, "
            SQL = SQL & " prex.smp_no AS BARCODE, coif.exam_mach_cd, rslt.exam_stus, "
            SQL = SQL & " prex.pt_no AS PID, "
            SQL = SQL & " ptbs.pt_nm AS PNAME, "
            SQL = SQL & " ptbs.ssn_1, ptbs.ssn_2," & vbCr
            SQL = SQL & "       DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd) as gnl_add_typ_cd, xpsl.adms_ymd , xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd, Max(Trim(coif.lmt_trm_day))" & vbCr
            SQL = SQL & "  FROM scrprexh prex, pmcptbsm ptbs, scccoifm coif, mosxpslh xpsl, scrrslth rslt" & vbCr
            SQL = SQL & " WHERE prex.hos_org_no               = '" & gHOSP.HOSPCD & "'" & vbCr
            SQL = SQL & "   AND SUBSTR(prex.acp_dt,1,8) BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
            SQL = SQL & "   AND rslt.hos_org_no    = prex.hos_org_no" & vbCr
            SQL = SQL & "   AND rslt.smp_no        = prex.smp_no" & vbCr
            SQL = SQL & "   AND rslt.prcp_seq      = prex.prcp_seq" & vbCr
            SQL = SQL & "   AND rslt.exam_seq      = prex.exam_seq" & vbCr
            SQL = SQL & "   AND rslt.exam_stus    IN ('0')" & vbCr
            SQL = SQL & "   AND ptbs.hos_org_no    = prex.hos_org_no" & vbCr
            SQL = SQL & "   AND ptbs.pt_no         = prex.pt_no" & vbCr
            SQL = SQL & "   AND coif.hos_org_no    = prex.hos_org_no" & vbCr
            SQL = SQL & "   AND coif.exam_cd       = prex.cd" & vbCr
            SQL = SQL & "   AND coif.use_typ       = 'Y'" & vbCr
            SQL = SQL & "   AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt" & vbCr
            SQL = SQL & "   AND coif.exam_mach_cd LIKE '" & gHOSP.MACHCD & "%'" & vbCr
            SQL = SQL & "   AND xpsl.smp_no        = prex.smp_no" & vbCr
            SQL = SQL & "   AND xpsl.hos_org_no    = prex.hos_org_no" & vbCr
            SQL = SQL & "   AND xpsl.prcp_typ_cd  IN ('O','C')" & vbCr
            SQL = SQL & "   GROUP BY prex.acp_dt, prex.smp_no, coif.exam_mach_cd ,rslt.exam_stus, prex.pt_no, ptbs.pt_nm, ptbs.ssn_1, ptbs.ssn_2, " & vbCr
            SQL = SQL & "            DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd), xpsl.adms_ymd,xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd" & vbCr
            SQL = SQL & "   ORDER BY prex.acp_dt, prex.smp_no " & vbCr
                    
            
            Call SetSQLData("워크조회", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count 가져옴
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
                frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
        Case "PLIS"
            SQL = ""
            SQL = SQL & "select distinct m.workarea             " & vbCr
            SQL = SQL & "     , m.accdt AS HOSPDATE             " & vbCr
            SQL = SQL & "     , m.accseq                        " & vbCr
            SQL = SQL & "     , m.spcyy                         " & vbCr
            SQL = SQL & "     , m.spcno                         " & vbCr
            SQL = SQL & "     , m.ptid AS PID                   " & vbCr
            SQL = SQL & "     , p.수진자명 AS PNAME             " & vbCr
            SQL = SQL & "     , m.rcvdt                         " & vbCr
            SQL = SQL & "     , m.rcvtm                         " & vbCr
            SQL = SQL & "     , COUNT(r.testcd) AS CNT          " & vbCr
            SQL = SQL & "  from plis..s2lab201 m                 " & vbCr
            SQL = SQL & "     , medichart..TB_인적사항 p             " & vbCr
            SQL = SQL & "     , plis..s2lab302 r                 " & vbCr
            SQL = SQL & "     , plis..s2lab001 e                 " & vbCr
            SQL = SQL & " where SUBSTRING(m.accdt,1,8) BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
            SQL = SQL & "   and r.testcd IN (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & "   and (r.vfydt IS NULL OR r.vfydt='')   " & vbCr
            SQL = SQL & "   and m.ptid = p.챠트번호                 " & vbCr
            SQL = SQL & "   and m.workarea = r.workarea             " & vbCr
            SQL = SQL & "   and m.accdt = r.accdt                   " & vbCr
            SQL = SQL & "   and m.accseq = r.accseq                 " & vbCr
            SQL = SQL & "   and r.testcd = e.testcd                 " & vbCr
            SQL = SQL & "  Group by m.workarea, m.accdt, m.spcyy,m.spcno,m.accseq, m.ptid,p.수진자명,m.rcvdt, m.rcvtm "
            SQL = SQL & "  Order by m.rcvdt, m.rcvtm                "
            
            Call SetSQLData("워크조회", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count 가져옴
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                frmWorkList.spdWork.MaxRows = 0
                strItems = ""
                Do Until RS.EOF
                    iCnt = iCnt + 1
                    getBarcode = Trim(RS("SPCYY")) & Format$(Trim(RS("SPCNO")), String$(9, "0"))

                    With frmWorkList.spdWork
                        .ReDraw = False
                        
                        For i = 1 To frmWorkList.spdWork.DataRowCnt
                            strDate = GetText(frmWorkList.spdWork, i, colHOSPDATE)
                            strBarcode = GetText(frmWorkList.spdWork, i, colBARCODE)
                            
                            If Trim(RS("HOSPDATE")) = strDate And getBarcode = strBarcode Then
                                blnSame = True
                            End If
                        Next
                        
                        If blnSame = False Then
                            .MaxRows = .MaxRows + 1
                            intRow = .MaxRows
                                
                            SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                            SetText frmWorkList.spdWork, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                            SetText frmWorkList.spdWork, getBarcode, intRow, colBARCODE
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
                            SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                            SetText frmWorkList.spdWork, Trim(RS.Fields("workarea")) & "", intRow, colRACKNO
                            SetText frmWorkList.spdWork, Trim(RS.Fields("accseq")) & "", intRow, colPOSNO
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
                frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
        Case "AMIS"
            SQL = ""
            SQL = SQL & "SELECT R.SPCMNO as BARCODE, P.PATID as PID, P.PATNAME as PNAME, P.SEX as SEX, O.ACPTDATE as HOSPDATE " & vbCr
            SQL = SQL & ", O.ACPTSEQ, O.RSVACPTSTATE, O.RESULTSTATE, O.DEPTCODE, O.ORDERDATE, O.SLIPNO, O.IOFLAG, O.ORDERCODE, O.ORDERNAME, R.RESULTFLAG, R.RESULTNO" & vbCr
            SQL = SQL & ", R.RESULTITEMCODE as ITEM " & vbCr
            SQL = SQL & "  FROM registinfos O, resultofnum R, PATMST P " & vbCr
            SQL = SQL & " WHERE O.acptdate = R.acptdate " & vbCr
            SQL = SQL & "   AND O.acptdate between '" & pFrom & "' and '" & pTo & "'" & vbCr
            SQL = SQL & "   AND R.resultitemcode in (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND O.patid = R.patid " & vbCr
            SQL = SQL & "   AND O.acptseq = R.acptseq " & vbCr
            SQL = SQL & "   AND O.patid = P.patid " & vbCr
            SQL = SQL & "   AND R.ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND O.CLAS = 4 " & vbCr '임상병리
            If frmWorkList.chkSave.Value = "0" Then
                SQL = SQL & "   AND R.RESULTFLAG = 0 " & vbCr
            End If
            SQL = SQL & "  ORDER BY R.SPCMNO"
            
            Call SetSQLData("워크조회", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count 가져옴
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
                frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
        Case "MSINFOTEC"
            SQL = ""
            SQL = SQL & "Select DISTINCT a.ORDT as HOSPDATE"
            SQL = SQL & ", b.PANM as PNAME,a.SPNO as BARCODE,a.PAID as PID, a.OIFL as INOUT,b.SEXS as SEX,b.AGES as AGE,a.NWNO as CHARTNO" & vbCr
            SQL = SQL & "  From emr.LRESULT a, emr.APATINF b" & vbCr
            SQL = SQL & " Where a.ORDT between  '" & pFrom & "' and '" & pTo & "'" & vbCr
            SQL = SQL & "   And a.PAID = b.PAID " & vbCr
            SQL = SQL & "   And a.ORCD in (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   And a.OKFL = 'N' " & vbCr  '-- 결과확정유무
            SQL = SQL & "   AND (a.RSFL IS NULL OR a.RSFL = 'N' OR a.RSFL = '')" & vbCr
            SQL = SQL & " ORDER BY a.ORDT, a.SPNO "
            
            Call SetSQLData("워크조회", SQL)
            
            frmWorkList.txtQuery.Text = SQL
        
            '-- Record Count 가져옴
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
                            SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                            SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                            SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
'                            SetText SPD, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "입원", "외래"), intRow, colINOUT
                            SetText SPD, frmWorkList.txtSeq.Text, intRow, colSEQNO
                            SetText SPD, GetSampleITEM(intRow), intRow, colITEMS
                            
                            frmWorkList.txtSeq.Text = frmWorkList.txtSeq.Text + 1
                        
                        End If
                    End With
                    
                    blnSame = False
                
                    DoEvents
                    
                    RS.MoveNext
                Loop
                frmWorkList.chkAll.Value = "1"
            Else
                frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
                frmWorkList.chkAll.Value = "0"
            End If
            
            RS.Close
            
    End Select

     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


'-- 장비결과 조회
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
    '-- 검사결과
    SQL = SQL & ",SEQNO,EXAMNAME,RESULT,REFJUDGE" & vbCr
    
    SQL = SQL & "  FROM PATRESULT " & vbCr
    '-- 검사일자
    If pDateType = 0 Then
        SQL = SQL & " WHERE EXAMDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
    '-- 접수일자
    Else
        SQL = SQL & " WHERE HOSPDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
    End If
    
    '-- 전송
    If pOpt = 1 Then
        SQL = SQL & "   AND SENDFLAG = '2' " & vbCr
    '-- 미전송
    ElseIf pOpt = 2 Then
        SQL = SQL & "   AND SENDFLAG <> '2' " & vbCr
    End If
    
    SQL = SQL & "   AND EXAMCODE IN (" & gAllTestCd & ") " & vbCr
    
    If pDateType = 0 Then
        SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,BARCODE,SEQNO"
    Else
        SQL = SQL & " ORDER BY HOSPDATE,SAVESEQ,BARCODE,SEQNO "
    End If
    
    '-- Record Count 가져옴
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
                            SetText frmMain.spdROrder, "장비결과", intRow, colSTATE
                    Case "2"
                            SetText frmMain.spdROrder, "전송완료", intRow, colSTATE
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
        'frmMain.lblStatus.Caption = ">> 조회 대상자가 없습니다."
        frmMain.chkRAll.Value = "0"
    End If
    
    RS.Close
     
    frmMain.spdROrder.RowHeight(-1) = 12
    frmMain.spdROrder.ReDraw = True
    
    Call frmMain.GetPatTRestResult_Search(1)
    
    Screen.MousePointer = 0

End Sub

'-- 검사자 ITEM 가져오기
Function GetSampleITEM(ByVal asRow As Long) As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strRegDate      As String
    Dim lngExamNo       As Long
    Dim strItems        As String
    Dim strSpcYY        As String
    Dim strSpcNo        As String
    
    GetSampleITEM = ""
    
    strRegDate = Trim(GetText(frmWorkList.spdWork, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(frmWorkList.spdWork, asRow, colBARCODE))
    strPatID = Trim(GetText(frmWorkList.spdWork, asRow, colPID))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    If Len(strBarcode) >= 11 Then
        strSpcYY = Mid(strBarcode, 1, 2)
        strSpcNo = Mid(strBarcode, 3, 9)
    End If
    
    Select Case gOCS
        Case "PHILL"
            '-- 알러지일 경우 바코드번호(전송일자 + 환자번호)를 인터페이스에서 만들어서 전송하기 때문에
            '-- 전송일자,환자번호를 바코드번호에서 찾아와서 조회한다.
            strRegDate = Mid(strBarcode, 1, 8)
            lngExamNo = Val(Mid(strBarcode, 9))
            
            SQL = ""
            SQL = SQL & "SELECT DISTINCT b.exam_code AS ITEM " & vbCr
            SQL = SQL & "  FROM TRUST a, TRURES b " & vbCr
            SQL = SQL & " WHERE a.request_date = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   AND a.request_date = b.request_date " & vbCr
            SQL = SQL & "   AND a.exam_no = '" & lngExamNo & "'"
            SQL = SQL & "   AND b.pro_code IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND b.exam_code <> 'X999' " & vbCr
            SQL = SQL & "   AND b.exam_code IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND a.exam_no = b.exam_no " & vbCr
            SQL = SQL & " ORDER BY b.exam_code "
    
        Case "JWINFO"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT LABCODE AS ITEM " & vbCr
            SQL = SQL & "   FROM SLA_Labresult " & vbCr
            SQL = SQL & " WHERE ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND LABCODE IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND RECEIPTDATE = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   AND SPECIMENNUM = '" & strBarcode & "'" & vbCr
            'SQL = SQL & "   AND JSTATUS < '3'" & vbCr
            SQL = SQL & "  ORDER BY LABCODE "
        
        Case "MCC"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM " & vbCr
            SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
            SQL = SQL & " WHERE READING_YMD = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   AND BCODE_NO = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & " ORDER BY ORD_CD "
        
        Case "GINUS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT CD AS ITEM " & vbCr
            SQL = SQL & "  FROM scrrslth " & vbCr
            SQL = SQL & " WHERE hos_org_no = '" & gHOSP.HOSPCD & "'" & vbCr
            SQL = SQL & "   AND SMP_NO = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & " ORDER BY CD "
    
        Case "PLIS"
            SQL = ""
            SQL = SQL & "select distinct r.testcd AS ITEM                " & vbCr
            SQL = SQL & "  from plis..s2lab201 m                 " & vbCr
            SQL = SQL & "     , plis..s2lab302 r                 " & vbCr
            SQL = SQL & "     , plis..s2lab001 e                 " & vbCr
            SQL = SQL & " where m.spcyy = '" & strSpcYY & "'    " & vbCr
            SQL = SQL & "   and m.spcno = '" & strSpcNo & "'    " & vbCr
'            SQL = SQL & "   and m.workarea = '" & gHOSP.LABCD & "'  " & vbCr
            SQL = SQL & "   and r.testcd IN (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & "   and (r.vfydt IS NULL OR r.vfydt='')   " & vbCr
            SQL = SQL & "   and m.workarea = r.workarea             " & vbCr
            SQL = SQL & "   and m.accdt = r.accdt                   " & vbCr
            SQL = SQL & "   and m.accseq = r.accseq                 " & vbCr
            SQL = SQL & "   and r.testcd = e.testcd                 " & vbCr
            SQL = SQL & "  Order by r.testcd                "
    
        Case "AMIS"
            SQL = ""
            SQL = SQL & "SELECT R.RESULTITEMCODE as ITEM " & vbCr
            SQL = SQL & "  FROM registinfos O, resultofnum R " & vbCr
            SQL = SQL & " WHERE O.acptdate = R.acptdate " & vbCr
            SQL = SQL & "   AND R.SPCMNO = '" & strBarcode & "'" & vbCr
            'SQL = SQL & "   AND O.acptdate between '" & pFrom & "' and '" & pTo & "'"
            SQL = SQL & "   AND R.resultitemcode in (" & gAllTestCd & ")"
            SQL = SQL & "   AND O.patid = R.patid " & vbCr
            SQL = SQL & "   AND O.acptseq = R.acptseq " & vbCr
            SQL = SQL & "   AND O.CLAS = 4 " & vbCr '임상병리
            SQL = SQL & "   AND R.ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND R.RESULTFLAG = 0 " & vbCr
            SQL = SQL & "  ORDER BY R.RESULTITEMCODE"
        
        Case "MSINFOTEC"
            SQL = ""
            SQL = SQL & "Select DISTINCT a.ORCD as ITEM " & vbCr
            SQL = SQL & "  From emr.LRESULT a" & vbCr
            SQL = SQL & " Where a.PAID = '" & strPatID & "' " & vbCr
            SQL = SQL & "   And a.SPNO = '" & strBarcode & "' " & vbCr
            SQL = SQL & "   And a.ORCD in (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   And a.OKFL <> 'Y' "   '-- 결과확정유무
        
    End Select
            
    Call SetSQLData("ITEM조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With frmMain.spdOrder
                .ReDraw = False
                If strItems = "" Then
                    strItems = GetTestNm(Trim(RS.Fields("ITEM")) & "", False)
                Else
                    strItems = strItems & "/" & GetTestNm(Trim(RS.Fields("ITEM")), False)
                End If
                
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    GetSampleITEM = strItems
    
    RS.Close
    
End Function

'Public Function GetTimeFull() As String
''Server의 현재 시간을 가져온다
''Return = 10:00:00
'    SQL = "select convert(char(8),getdate(),108) "
'    db_select_Var gServer, SQL, GetTimeFull
'End Function
'
'Public Function GetTimeShort() As String
''Server의 현재 시간을 가져온다
''Return = 10:00
'    SQL = "select convert(char(5),getdate(),108) "
'    db_select_Var gServer, SQL, GetTimeShort
'End Function


'Public Function GetDateFull_ORCL() As String
'    Dim s As String
'    Dim t As String
'
'
''Oracle : Server의 현재 날짜를 가져온다
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


'-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기
Function GetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    
    GetEquipExamCode_AU480 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmMain.spdOrder, intRow, colBARCODE))    '2 샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    
    frmMain.vasTemp.MaxRows = 0
    
    
    '-- 가져온 검사코드의 채널 찾기
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

'-- 검사자 정보 가져오기
Function GetSampleInfo(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strRegDate      As String
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
    
    
On Error GoTo DBErr
    
    GetSampleInfo = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    Select Case gOCS
        Case "MCC"
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
            SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
            SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND STS_CD = '0'" & vbCr    '0 접수, 1:결과전송
            SQL = SQL & " ORDER BY ORD_CD "
            
            Call SetSQLData("바코드조회", SQL)
            
            '-- Record Count 가져옴
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            
            '-- 2017.09.05
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
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                        SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                        SetText SPD, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "입원", "외래"), asRow, colINOUT
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                        
                        '01   Serum (SST)
                        '02   EDTA
                        '03   S.citrate
                        '04   Urine
                        '05   CSF
                        '07   Stool
                        '11  Pleural fluid
                        '20  전용
                        '22  Biopsy
                        
                        If Trim(RS.Fields("SPCCD")) & "" = "01" Then 'Serum
                            mOrder.SPCCD = "1"
                        ElseIf Trim(RS.Fields("SPCCD")) & "" = "04" Then 'Urine
                            mOrder.SPCCD = "2"
                        Else
                            mOrder.SPCCD = "1"  'Default 를 Serum 으로 한다.
                        End If
                        
                        '-- 화면에 표시
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "◇", asRow, intCol)
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
            
        Case "GINUS"
                  SQL = "SELECT /*+ INDEX (coif scccoifm_ix1) INDEX (prex scrprexh_ix3) INDEX (ptbs pmcptbsm_ux1) INDEX (rslt scrrslth_ux1) INDEX (xpsl mosxpslh_ix2) */" & vbCr
            SQL = SQL & " prex.acp_dt AS HOSPDATE, "
            SQL = SQL & " prex.smp_no AS BARCODE, coif.exam_mach_cd, rslt.exam_stus, "
            SQL = SQL & " prex.pt_no AS PID, "
            SQL = SQL & " ptbs.pt_nm AS PNAME, "
            SQL = SQL & " rslt.cd AS ITEM, "
            SQL = SQL & " ptbs.ssn_1, ptbs.ssn_2," & vbCr
            SQL = SQL & "       DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd) as gnl_add_typ_cd, xpsl.adms_ymd , xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd, Max(Trim(coif.lmt_trm_day))" & vbCr
            SQL = SQL & "  FROM scrprexh prex, pmcptbsm ptbs, scccoifm coif, mosxpslh xpsl, scrrslth rslt" & vbCr
            SQL = SQL & " WHERE prex.hos_org_no               = '" & gHOSP.HOSPCD & "'" & vbCr
            SQL = SQL & "   AND SUBSTR(prex.acp_dt,1,8) BETWEEN '" & Format(frmMain.dtpFrDt.Value, "yyyymmdd") & "' AND '" & Format(frmMain.dtpToDt.Value, "yyyymmdd") & "'" & vbCr
            SQL = SQL & "   AND rslt.hos_org_no    = prex.hos_org_no" & vbCr
            SQL = SQL & "   AND rslt.smp_no        = prex.smp_no" & vbCr
            SQL = SQL & "   AND rslt.prcp_seq      = prex.prcp_seq" & vbCr
            SQL = SQL & "   AND rslt.exam_seq      = prex.exam_seq" & vbCr
            SQL = SQL & "   AND rslt.exam_stus    IN ('0')" & vbCr
            SQL = SQL & "   AND ptbs.hos_org_no    = prex.hos_org_no" & vbCr
            SQL = SQL & "   AND ptbs.pt_no         = prex.pt_no" & vbCr
            SQL = SQL & "   AND coif.hos_org_no    = prex.hos_org_no" & vbCr
            SQL = SQL & "   AND coif.exam_cd       = prex.cd" & vbCr
            SQL = SQL & "   AND coif.use_typ       = 'Y'" & vbCr
            SQL = SQL & "   AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt" & vbCr
            SQL = SQL & "   AND coif.exam_mach_cd LIKE '" & gHOSP.MACHCD & "%'" & vbCr
            SQL = SQL & "   AND xpsl.smp_no        = prex.smp_no" & vbCr
            SQL = SQL & "   AND xpsl.hos_org_no    = prex.hos_org_no" & vbCr
            SQL = SQL & "   AND xpsl.prcp_typ_cd  IN ('O','C')" & vbCr
            SQL = SQL & "   AND prex.smp_no = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   GROUP BY prex.acp_dt, prex.smp_no, coif.exam_mach_cd ,rslt.exam_stus, prex.pt_no, ptbs.pt_nm, rslt.cd, ptbs.ssn_1, ptbs.ssn_2 " & vbCr
            SQL = SQL & "            , DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd), xpsl.adms_ymd,xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd" & vbCr
            SQL = SQL & "   ORDER BY prex.acp_dt, prex.smp_no " & vbCr
            
            Call SetSQLData("바코드조회", SQL)
            
            '-- Record Count 가져옴
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            
            '-- 2017.09.05
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
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                        
                        '-- 화면에 표시
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "◇", asRow, intCol)
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
            
        Case "JWINFO"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "  a.RECEIPTDATE AS HOSPDATE"
            SQL = SQL & ", a.SPECIMENNUM AS BARCODE"
            SQL = SQL & ", a.RECEIPTNO AS CHARTNO"
            SQL = SQL & ", a.IPDOPD AS INOUT "
            SQL = SQL & ", a.PTNO AS PID"
            SQL = SQL & ", a.SNAME AS PNAME"
            SQL = SQL & ", COUNT(a.LABCODE) AS CNT " & vbCr
            SQL = SQL & "   FROM SLA_LabMaster a, SLA_LabResult b" & vbCr
            SQL = SQL & " WHERE a.SPECIMENNUM = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND a.ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND b.LABCODE IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND a.RECEIPTNO = b.RECEIPTNO " & vbCr
            SQL = SQL & "   AND a.ORDERCODE = b.ORDERCODE " & vbCr
            SQL = SQL & "   and a.RECEIPTDATE = b.RECEIPTDATE" & vbCr
            SQL = SQL & "   AND a.SPECIMENNUM = b.SPECIMENNUM" & vbCr
            SQL = SQL & "   AND a.JSTATUS < '3'" & vbCr
            SQL = SQL & " GROUP BY a.RECEIPTDATE, a.SPECIMENNUM, a.RECEIPTNO, a.IPDOPD, a.PTNO, a.SNAME" & vbCr
            SQL = SQL & " ORDER BY a.RECEIPTDATE,a.SPECIMENNUM "
                
            Call SetSQLData("바코드조회", SQL)
            
            '-- Record Count 가져옴
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
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                        SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                        SetText SPD, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "입원", "외래"), asRow, colINOUT
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                
                        '-- 화면에 표시
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "◇", asRow, intCol)
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
        
        Case "PLIS"
            If Len(strBarcode) = 11 Then
                strSpcYY = Mid(strBarcode, 1, 2)
                strSpcNo = Mid(strBarcode, 3, 9)
            Else
                Exit Function
            End If
            
            SQL = ""
            SQL = SQL & "select distinct m.workarea             " & vbCr
            SQL = SQL & "     , m.accdt AS HOSPDATE             " & vbCr
            SQL = SQL & "     , m.accseq                        " & vbCr
            SQL = SQL & "     , m.spcyy                         " & vbCr
            SQL = SQL & "     , m.spcno                         " & vbCr
            SQL = SQL & "     , m.deptcd                        " & vbCr
            SQL = SQL & "     , m.SEX                           " & vbCr
            SQL = SQL & "     , m.AgeDay                        " & vbCr
            SQL = SQL & "     , m.ptid AS PID                   " & vbCr
            SQL = SQL & "     , p.수진자명 AS PNAME             " & vbCr
            SQL = SQL & "     , m.rcvdt                         " & vbCr
            SQL = SQL & "     , m.rcvtm                         " & vbCr
            SQL = SQL & "     , r.testcd AS ITEM                " & vbCr
            SQL = SQL & "     , e.abbrnm10                      " & vbCr
            SQL = SQL & "     , m.QCFG                          " & vbCr
            SQL = SQL & "  from plis..s2lab201 m                " & vbCr
            SQL = SQL & "     , medichart..TB_인적사항 p        " & vbCr
            SQL = SQL & "     , plis..s2lab302 r                " & vbCr
            SQL = SQL & "     , plis..s2lab001 e                " & vbCr
            SQL = SQL & " where m.spcyy = '" & strSpcYY & "'    " & vbCr
            SQL = SQL & "   and m.spcno = '" & strSpcNo & "'    " & vbCr
'            SQL = SQL & "   and m.workarea = '" & gHOSP.LABCD & "'  " & vbCr
            SQL = SQL & "   and r.testcd IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   and (r.vfydt IS NULL OR r.vfydt='') " & vbCr
            SQL = SQL & "   and m.ptid = p.챠트번호             " & vbCr
            SQL = SQL & "   and m.workarea = r.workarea         " & vbCr
            SQL = SQL & "   and m.accdt = r.accdt               " & vbCr
            SQL = SQL & "   and m.accseq = r.accseq             " & vbCr
            SQL = SQL & "   and r.testcd = e.testcd             " & vbCr
            SQL = SQL & "  Order by m.rcvdt, m.rcvtm            "
        
            Call SetSQLData("바코드조회", SQL)
            
            '-- Record Count 가져옴
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
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        SetText SPD, Trim(RS.Fields("accseq")) & "", asRow, colKEY1
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                        
                        If Trim(RS.Fields("QCFG")) & "" = "1" Then
                            mResult.Kind = "QC"
                        Else
                            mResult.Kind = ""
                        End If
                        
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        mOrder.WA = Trim(RS.Fields("workarea")) & ""
                        mOrder.AccSeq = RS.Fields("accseq")
                        
                        '-- 화면에 표시
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "◇", asRow, intCol)
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
            
        Case "AMIS"
            SQL = ""
            SQL = SQL & "SELECT R.SPCMNO as BARCODE, P.PATID as PID, P.PATNAME as PNAME, P.SEX as SEX, O.ACPTDATE as HOSPDATE " & vbCr
            SQL = SQL & ", O.ACPTSEQ, O.RSVACPTSTATE, O.RESULTSTATE, O.DEPTCODE, O.ORDERDATE, O.SLIPNO, O.IOFLAG, O.ORDERCODE, O.ORDERNAME, R.RESULTFLAG, R.RESULTNO" & vbCr
            SQL = SQL & ", R.RESULTITEMCODE as ITEM " & vbCr
            SQL = SQL & "  FROM registinfos O, resultofnum R, PATMST P " & vbCr
            SQL = SQL & " WHERE O.acptdate  = R.acptdate " & vbCr
            SQL = SQL & "   AND R.SPCMNO    = '" & strBarcode & "'"
            SQL = SQL & "   AND O.patid     = R.patid " & vbCr
            SQL = SQL & "   AND O.acptseq   = R.acptseq " & vbCr
            SQL = SQL & "   AND O.patid     = P.patid " & vbCr
            SQL = SQL & "   AND R.resultitemcode IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND R.ORDERCODE      IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND O.CLAS = 4 " & vbCr '임상병리
            SQL = SQL & "   AND R.RESULTFLAG = 0 " & vbCr
            SQL = SQL & "  ORDER BY R.SPCMNO"
            
            Call SetSQLData("바코드조회", SQL)
            
            '-- Record Count 가져옴
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
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                
                        '-- 화면에 표시
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "◇", asRow, intCol)
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
        
    
        Case "MSINFOTEC"
            SQL = ""
            SQL = SQL & "Select DISTINCT a.ORDT as HOSPDATE"
            SQL = SQL & ", b.PANM as PNAME"
            SQL = SQL & ", a.SPNO as BARCODE"
            SQL = SQL & ", a.PAID as PID"
            SQL = SQL & ", a.OIFL as INOUT"
            SQL = SQL & ", b.SEXS as SEX"
            SQL = SQL & ", b.AGES as AGE"
            SQL = SQL & ", a.NWNO as CHARTNO" & vbCr
            SQL = SQL & ", a.ORCD as ITEM " & vbCr
            SQL = SQL & "  From emr.LRESULT a, emr.APATINF b" & vbCr
            SQL = SQL & " Where a.SPNO = '" & strBarcode & "'" & vbCr
            If strPatID <> "" Then
                SQL = SQL & "   And a.PAID = '" & strPatID & "'" & vbCr
            End If
            SQL = SQL & "   And a.PAID = b.PAID " & vbCr
            SQL = SQL & "   And a.ORCD in (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   And a.OKFL = 'N' "   '-- 결과확정유무
            SQL = SQL & "   AND (a.RSFL IS NULL OR a.RSFL = 'N' OR a.RSFL = '')" & vbCr
            SQL = SQL & " ORDER BY a.ORDT, a.SPNO "
            
            Call SetSQLData("바코드조회", SQL)
            
            '-- Record Count 가져옴
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
                        SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                        SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                        mOrder.PID = Trim(RS.Fields("PID")) & ""
                        SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                        mOrder.PName = Trim(RS.Fields("PNAME")) & ""
                        SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                        SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                        'SetText SPD, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "입원", "외래"), asRow, colINOUT
                        SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                
                        '-- 화면에 표시
                        For intCol = colSTATE + 1 To .MaxCols
                            If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                Call SetText(SPD, "◇", asRow, intCol)
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
            SQL = SQL & "SAVESEQ"                           '저장순번(날짜별)
            SQL = SQL & ", EXAMDATE"                        '검사일자"
            SQL = SQL & ", HOSPDATE"                        '병원접수일자"
            SQL = SQL & ", EQUIPNO"                         '장비코드"
            SQL = SQL & ", BARCODE" & vbCrLf                '검체번호
            SQL = SQL & ", EQUIPCODE"                       '검사채널"
            SQL = SQL & ", ORDERCODE"                       '병원처방코드"
            SQL = SQL & ", EXAMCODE"                        '병원검사코드"
            SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
            SQL = SQL & ", EXAMNAME"                        '검사명
            SQL = SQL & ", SEQNO" & vbCrLf                  '검사일련번호"
            SQL = SQL & ", SAMPLETYPE"                      '검체유형"
            SQL = SQL & ", INOUT"                           '입/외
            SQL = SQL & ", DISKNO"                          'Rack (VERSACELL 에서는 실제 검사장비코드를 저장한다..)
            SQL = SQL & ", POSNO"                           'Pos
            SQL = SQL & ", EQUIPRESULT"                     '장비결과"
            SQL = SQL & ", RESULT" & vbCrLf                 'LIS 결과"
            SQL = SQL & ", REFJUDGE"                        '판정
            SQL = SQL & ", REFFLAG"                         'flag
            SQL = SQL & ", REFVALUE"                        '참고치
            SQL = SQL & ", CHARTNO"                         '챠트번호
            SQL = SQL & ", PID"                             '병록번호(내원번호)"
            SQL = SQL & ", PNAME" & vbCrLf
            SQL = SQL & ", PSEX"
            SQL = SQL & ", PAGE"
            SQL = SQL & ", PJUMIN"
            SQL = SQL & ", PANICVALUE"
            SQL = SQL & ", DELTAVALUE" & vbCrLf
            SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
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
            SQL = SQL & ",''"                                                   '검체유형
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
            'SQL = SQL & ",'" & mOrder.AccSeq & "'"                              'panic (accseq 대체사용)
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colKEY1)) & "'"                              'panic (accseq 대체사용)
            SQL = SQL & ",''"                                                   'delta
            SQL = SQL & ",'0'"                                                  '전송구분(0:미전송,1:전송)
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
            SQL = SQL & "SAVESEQ"                           '저장순번(날짜별)
            SQL = SQL & ", EXAMDATE"                        '검사일자"
            SQL = SQL & ", HOSPDATE"                        '병원접수일자"
            SQL = SQL & ", EQUIPNO"                         '장비코드"
            SQL = SQL & ", BARCODE" & vbCrLf                '검체번호
            SQL = SQL & ", EQUIPCODE"                       '검사채널"
            SQL = SQL & ", ORDERCODE"                       '병원처방코드"
            SQL = SQL & ", EXAMCODE"                        '병원검사코드"
            SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
            SQL = SQL & ", EXAMNAME"                        '검사명
            SQL = SQL & ", SEQNO" & vbCrLf                  '검사일련번호"
            SQL = SQL & ", SAMPLETYPE"                      '검체유형"
            SQL = SQL & ", INOUT"                           '입/외
            SQL = SQL & ", DISKNO"                          'Rack (VERSACELL 에서는 실제 검사장비코드를 저장한다..)
            SQL = SQL & ", POSNO"                           'Pos
            SQL = SQL & ", EQUIPRESULT"                     '장비결과"
            SQL = SQL & ", RESULT" & vbCrLf                 'LIS 결과"
            SQL = SQL & ", REFJUDGE"                        '판정
            SQL = SQL & ", REFFLAG"                         'flag
            SQL = SQL & ", REFVALUE"                        '참고치
            SQL = SQL & ", CHARTNO"                         '챠트번호
            SQL = SQL & ", PID"                             '병록번호(내원번호)"
            SQL = SQL & ", PNAME" & vbCrLf
            SQL = SQL & ", PSEX"
            SQL = SQL & ", PAGE"
            SQL = SQL & ", PJUMIN"
            SQL = SQL & ", PANICVALUE"
            SQL = SQL & ", DELTAVALUE" & vbCrLf
            SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
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
            SQL = SQL & ",''"                                                   '검체유형
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
            SQL = SQL & ",'0'"                                                  '전송구분(0:미전송,1:전송)
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
        SQL = SQL & "SAVESEQ"                           '저장순번(날짜별)
        SQL = SQL & ", EXAMDATE"                        '검사일자"
        SQL = SQL & ", HOSPDATE"                        '병원접수일자"
        SQL = SQL & ", EQUIPNO"                         '장비코드"
        SQL = SQL & ", BARCODE" & vbCrLf                '검체번호
        SQL = SQL & ", EQUIPCODE"                       '검사채널"
        SQL = SQL & ", ORDERCODE"                       '병원처방코드"
        SQL = SQL & ", EXAMCODE"                        '병원검사코드"
        SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
        SQL = SQL & ", EXAMNAME" & vbCrLf               '검사명
        SQL = SQL & ", SEQNO"                           '검사일련번호"
        SQL = SQL & ", SAMPLETYPE"                      '검체유형"
        SQL = SQL & ", INOUT"                           '입/외
        SQL = SQL & ", DISKNO"                          'Rack (VERSACELL 에서는 실제 검사장비코드를 저장한다..)
        SQL = SQL & ", POSNO" & vbCrLf                  'Pos
        SQL = SQL & ", EQUIPRESULT"                     '장비결과"
        SQL = SQL & ", RESULT"                          'LIS 결과"
        SQL = SQL & ", REFJUDGE"                        '판정
        SQL = SQL & ", REFFLAG"                         'flag
        SQL = SQL & ", REFVALUE" & vbCrLf               '참고치
        SQL = SQL & ", CHARTNO"                         '챠트번호
        SQL = SQL & ", PID"                             '병록번호(내원번호)"
        SQL = SQL & ", PNAME"
        SQL = SQL & ", PSEX"
        SQL = SQL & ", PAGE" & vbCrLf
        SQL = SQL & ", PJUMIN"
        SQL = SQL & ", PANICVALUE"
        SQL = SQL & ", DELTAVALUE"
        SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
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
        SQL = SQL & ",''"                                                   '검체유형
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
        SQL = SQL & ",'0'"                                                  '전송구분(0:미전송,1:전송)
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
        SQL = SQL & " EQUIPRESULT = '" & asEquipResult & "'"                                            '장비결과
        SQL = SQL & ",RESULT      = '" & asEquipResult & "'" & vbCr                                     'LIS 결과
        SQL = SQL & " WHERE SAVESEQ  = " & Trim(GetText(.spdROrder, asRow1, colSAVESEQ)) & vbCr         '저장순번(날짜별)
        SQL = SQL & "   AND EXAMDATE = '" & sExamDate & "'" & vbCr                                      '검사일자
        SQL = SQL & "   AND HOSPDATE = '" & Trim(GetText(.spdROrder, asRow1, colHOSPDATE)) & "'" & vbCr '병원접수일자
        SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCr                                   '장비코드
        SQL = SQL & "   AND BARCODE  = '" & Trim(GetText(.spdROrder, asRow1, colBARCODE)) & "'" & vbCr  '검체번호
        SQL = SQL & "   AND EXAMCODE = '24HRS-V'"
        If Not DBExec(AdoCn_Local, SQL) Then
            'SaveQuery SQL
            'Exit Function
        End If
        
'        Call CalProcess(gRow)
        
    End With
    
End Function

'-- 계산값 처리
'01   Serum (SST)
'02   EDTA
'03   S.citrate
'04   Urine
'05   CSF
'07   Stool
'11  Pleural fluid
'20  전용
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
    
    '1. 계산 대상 검사항목을 찾는다.
    Select Case pTestCd
        Case "C3730N1"  ' : 투석후 = C3730N1
                strCalTestCd = "URR"        '요소감소율
        Case "C3750"    'Creatine
                strCalTestCd = "EGFR"       'MDRD eGFR
        Case "C3791" 'NA
                strCalTestCd = "C3791N1"    'NA(24시간뇨)
        Case "C3792" 'K
                strCalTestCd = "C3792N1"    'K(24시간뇨)
        Case "C3793" 'Cl
                strCalTestCd = "C3793N1"    'Cl(24시간뇨)
        Case "C2200-1" 'micro TP
                strCalTestCd = "C2200-2"    'UTP(24시간뇨)
        Case "C3730" 'BUN
                strCalTestCd = "C3730-2"    'BUN(24시간뇨)
        Case "C3750N1" 'Crea
                strCalTestCd = "C3750N1"   'Crea(24시간뇨)
        Case "C3750N3"  'Crea(단회뇨)
                strCalTestCd = "C7230"      'MicroALB retio
        Case "C2302N6"    'M.alb
                strCalTestCd = "C7230"      'MicroALB retio
        Case Else
                Exit Function
    End Select
    
    '1. 계산처방항목이 있는지 찾는다.
          SQL = ""
    SQL = SQL & "SELECT COUNT(*) AS CNT" & vbCr
    SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
    SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
    SQL = SQL & "   AND ORD_CD = '" & strCalTestCd & "'"
        
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        If IsNull(RS.Fields("CNT")) Or RS.Fields("CNT") = 0 Then
            Exit Function
        End If
    End If
    RS.Close
    
    
    '2. 저장된 환자정보와 결과값을 가져온다.
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
    'SQL = SQL & "   AND STS_CD = '0'" & vbCr    '0 접수, 1:결과전송
        
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strResult = Trim(RS.Fields("RESULT")) & ""
        strAge = Trim(RS.Fields("AGE")) & ""
        strSex = Trim(RS.Fields("SEX")) & ""
        strSpc = Trim(RS.Fields("SPCCD")) & ""
        strPTID = Trim(RS.Fields("PID")) & ""
    End If
    
    If strResult = "" Then
        '서버에서 못찾을 경우..(로컬)
        SQL = ""
        SQL = SQL & "SELECT RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE BARCODE = '" & strBarcode & "'" & vbCr
        SQL = SQL & "   AND EXAMCODE = '" & pTestCd & "'"

        '-- Record Count 가져옴
        AdoCn_Local.CursorLocation = adUseClient
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            strResult = Trim(RS_L.Fields("RESULT")) & ""
        End If

        RS_L.Close
    End If
    
    RS.Close
    
    Select Case pTestCd
        Case "URR"    '요소감소울
            '환자번호,기준일,기준시간,처방코드입니다.
            '투석전결과[C3730N2]를 가져온다.
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
        Case "C3750"    'Creatine   ==> eGFR 계산
            If IsNumeric(strResult) And CCur(strResult) > 0 And strSex <> "" And strAge <> "" Then
                '18세 이상만 적용
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

        Case "C3750N3"      'Crea(단회뇨) 이면 M.alb 결과를 가져온다.
            SQL = "SELECT [dbo].FUN_H7LIS_PRE_RESULT4('" & strPTID & "', '" & Format(Now, "yyyymmdd") & "', '" & Format(Now, "hhmm") & "', 'C2302N6')"
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                strPreResult = Trim(RS.Fields(0)) & ""
            End If
            
            If strPreResult = "" Then
                '같이 처방나는 코드여서 못찾을 경우..(서버)
                SQL = ""
                SQL = SQL & "SELECT RESULT_NM AS RESULT" & vbCr
                SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
                SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   AND ORD_CD = 'C2302N6'"

                '-- Record Count 가져옴
                AdoCn.CursorLocation = adUseClient
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    strPreResult = Trim(RS.Fields("RESULT")) & ""
                End If

                RS.Close
                
                If strPreResult = "" Then
                    '서버에서 못찾을 경우..(로컬)
                    SQL = ""
                    SQL = SQL & "SELECT RESULT " & vbCr
                    SQL = SQL & "  FROM PATRESULT " & vbCr
                    SQL = SQL & " WHERE BARCODE = '" & strBarcode & "'" & vbCr
                    SQL = SQL & "   AND EXAMCODE = 'C2302N6'"
    
                    '-- Record Count 가져옴
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
        
        Case "C2302N6"       'M.alb 이면  Crea(단회뇨)결과를 가져온다.
            SQL = "SELECT [dbo].FUN_H7LIS_PRE_RESULT4('" & strPTID & "', '" & Format(Now, "yyyymmdd") & "', '" & Format(Now, "hhmm") & "', 'C3750N3')"
            AdoCn.CursorLocation = adUseClient
            Set RS = AdoCn.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                strPreResult = Trim(RS.Fields(0)) & ""
            End If
            If strPreResult = "" Then
                '같이 처방나는 코드여서 못찾을 경우..(서버)
                SQL = ""
                SQL = SQL & "SELECT RESULT_NM AS RESULT" & vbCr
                SQL = SQL & "  FROM LIS_INTERFACE1_V " & vbCr
                SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   AND ORD_CD = 'C3750N3'"

                '-- Record Count 가져옴
                AdoCn.CursorLocation = adUseClient
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    strPreResult = Trim(RS.Fields("RESULT")) & ""
                End If

                RS.Close
                
                If strPreResult = "" Then
                    '서버에서 못찾을 경우..(로컬)
                    SQL = ""
                    SQL = SQL & "SELECT RESULT " & vbCr
                    SQL = SQL & "  FROM PATRESULT " & vbCr
                    SQL = SQL & " WHERE BARCODE = '" & strBarcode & "'" & vbCr
                    SQL = SQL & "   AND EXAMCODE = 'C3750N3'"
    
                    '-- Record Count 가져옴
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

            '-- 결과Row 추가
            lsRstRow = SPDRST.DataRowCnt + 1
            If SPDRST.MaxRows < lsRstRow Then
                SPDRST.MaxRows = lsRstRow
            End If
    
            '결과값 표시
            For intCol = colSTATE + 1 To SPDORD.MaxCols
                If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                    SetText SPDORD, strResult, ActiveRow, intCol
                    Exit For
                End If
            Next
    
            '-- 결과 List
            SetText SPDRST, lsSeqNo, lsRstRow, colRSEQNO                '순번
            SetText SPDRST, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
            SetText SPDRST, lsTestCode, lsRstRow, colRTESTCD            '검사코드
            SetText SPDRST, lsTestName, lsRstRow, colRTESTNM            '검사명
            SetText SPDRST, strIntBase, lsRstRow, colRCHANNEL           '장비채널
            SetText SPDRST, strCalResult, lsRstRow, colRMACHRESULT      '장비결과
            SetText SPDRST, strCalResult, lsRstRow, colRLISRESULT       'LIS결과
            SetText SPDRST, "", lsRstRow, colRJUDGE                     '판정
            SetText SPDRST, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
            
            '-- 로컬 저장
            If pTV = "" Then
                SetLocalDB ActiveRow, lsRstRow, "1", ""
            Else
                SetLocalDB_R ActiveRow, lsRstRow, "1", ""
            End If
            
            '-- 결과Count
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
                '-- 저장 실패
                SetForeColor SPDORD, ActiveRow, ActiveRow, 1, colSTATE, 255, 0, 0
                SetText SPDORD, "Failed", ActiveRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor SPDORD, ActiveRow, ActiveRow, 1, colSTATE, 202, 255, 112
                SetText SPDORD, "저장완료", ActiveRow, colSTATE
                SetText SPDORD, "0", ActiveRow, colCHECKBOX
                
                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.MACHCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(SPDORD, ActiveRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(SPDORD, ActiveRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(SPDORD, ActiveRow, colSAVESEQ)) & vbCrLf
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
        End If
    End If
    
End Function


'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
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

