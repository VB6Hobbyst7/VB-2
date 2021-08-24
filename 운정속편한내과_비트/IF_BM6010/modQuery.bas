Attribute VB_Name = "modQuery"
Option Explicit

Public SQL              As String
Public RS               As ADODB.Recordset
Public blnSameRecord    As Boolean


'-- 검사마스터 조회
Public Sub GetTestList()
    Dim intRow          As Long
    
    frmMain.spdTest.MaxRows = 0
    intRow = 0
    gAllTestCd = ""
    Erase gArrEQP
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT EQUIPCD,SEQNO,TESTCODE,SENDCHANNEL,RSLTCHANNEL " & vbCr
    SQL = SQL & " ,TESTNAME,ABBRNAME,RESPREC,REFLOW,REFHIGH" & vbCr
    SQL = SQL & " ,REFLOWF,REFHIGHF" & vbCr
    SQL = SQL & " ,RESULTTYPE,CUTOFFUSE,COLIN,COLCOMP,COLOUT" & vbCr
    SQL = SQL & " ,COHIN,COHCOMP,COHOUT,COMOUT" & vbCr
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
            
            '처방코드, SUB코드용 추가 16,17
            '여자상한,하한 추가  18,19
            ReDim Preserve gArrEQP(AdoRs_Local.RecordCount, 19)
            
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
                
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("REFLOWF").Value & "", intRow, colLLOWF)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("REFHIGHF").Value & "", intRow, colLHIGHF)
                
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
                gArrEQP(intRow, 16) = ""    '처방코드로 사용
                gArrEQP(intRow, 17) = ""    'SUB코드로 사용
                gArrEQP(intRow, 18) = AdoRs_Local.Fields("REFLOWF").Value & ""
                gArrEQP(intRow, 19) = AdoRs_Local.Fields("REFHIGHF").Value & ""
                
                If Trim(AdoRs_Local.Fields("TESTCODE").Value) <> "" Then
                    If intRow = 1 Or gAllTestCd = "" Then
                        gAllTestCd = "'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                    Else
                        gAllTestCd = gAllTestCd & ",'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                    End If
                End If
                
                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 15
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
    
    '-- Record Count 가져옴
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

'-- 검사명으로 결과채널 조회
Public Function GetRsltChannel(ByVal pItem As String) As String
    Dim RS2             As ADODB.Recordset
    Dim intRow          As Long
    
    GetRsltChannel = ""
    
    SQL = ""
    SQL = SQL & "SELECT RSLTCHANNEL "
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE ABBRNAME = '" & pItem & "'"
    
    Set RS2 = New ADODB.Recordset
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set RS2 = AdoCn_Local.Execute(SQL, , 1)
    If Not RS2.EOF = True And Not RS2.BOF = True Then
        Do Until RS2.EOF
            GetRsltChannel = RS2.Fields("RSLTCHANNEL").Value & ""
            RS2.MoveNext
        Loop
    End If
    
    RS2.Close
    
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
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_SetQCList_Header" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "SetQCList_Detail" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetQCList_Header" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
     
                strErrMSG = "위    치 : " & "GetQCList_QCID" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
     
                strErrMSG = "위    치 : " & "strQCFlag" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
     
                strErrMSG = "위    치 : " & "GetQCList_Detail" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
     
                strErrMSG = "위    치 : " & "GetQCResult_Detail" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
     
                strErrMSG = "위    치 : " & "GetQCResult_Detail" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Function

'-- 워크리스트 조회
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)

    Select Case gEMR
        Case "AMIS"                         '아미스
                Call GetWorkList_AMIS(pFrom, pTo, SPD)
        
        Case "BIGUBCARE"
                Call GetWorkList_BIGUBCARE(pFrom, pTo, SPD)
        
        Case "BIT"                          '비트
                Call GetWorkList_BIT(pFrom, pTo, SPD)

        Case "BITUCHART"                    '비트 U 챠트
                Call GetWorkList_BITUCHART(pFrom, pTo, SPD)

        Case "BIT70"                        '비트 HIB70
                Call GetWorkList_BIT70(pFrom, pTo, SPD)
        
        Case "EMEDI"                        '이메디
                Call GetWorkList_AMIS(pFrom, pTo, SPD)
        
        Case "EASYS"                        '이지스, MCC
                Call GetWorkList_EASYS(pFrom, pTo, SPD)
        
        Case "EONM"                         '이온엠
                Call GetWorkList_EONM(pFrom, pTo, SPD)
            
        Case "GINUS"                         '지누스
                Call GetWorkList_GINUS(pFrom, pTo, SPD)
        
        Case "GSEN"                         '지센커뮤니케이션즈(이챠트)
                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
        
        Case "HWASAN"                       '화산
                Call GetWorkList_HWASAN(pFrom, pTo, SPD)
        
        Case "JAINCOM"                      '자인컴
                Call GetWorkList_JAINCOM(pFrom, pTo, SPD)
        
        Case "JWINFO"                       '중외정보
                Call GetWorkList_JWINFO(pFrom, pTo, SPD)
            
        Case "KCHART"                       '다대소프트
                Call GetWorkList_KCHART(pFrom, pTo, SPD)
        
        Case "KCWH"                         '근로복지공단
                Call GetWorkList_KCWH(pFrom, pTo, SPD)
        
        Case "KOMAIN"                       '중외정보
                Call GetWorkList_KOMAIN(pFrom, pTo, SPD)
            
        Case "KYU"                          '건양대학교병원 - 워크리스트 기능없음
                'Call GetWorkList_KYU(pFrom, pTo, SPD)
        
        Case "MEDICHART"                    '메디챠트
                Call GetWorkList_MEDICHART(pFrom, pTo, SPD)
                     
        Case "MEDIIT"                       '메디IT(필의료재단)
                Call GetWorkList_MEDIIT(pFrom, pTo, SPD)
                     
        Case "MEDITOLISS"                   '아름누리
                Call GetWorkList_MEDITOLISS(pFrom, pTo, SPD)
        
        Case "MCC"                          'MCC SP버전
                Call GetWorkList_MCC(pFrom, pTo, SPD)
        
        Case "MOD"                          'MOD 시스템
                Call GetWorkList_MOD(pFrom, pTo, SPD)
        
        Case "MSINFOTEC"                    'MS인포텍
                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)

        Case "NEOSOFT"                      '네오소프트
                Call GetWorkList_NEOSOFT(pFrom, pTo, SPD)

        Case "ONITGUM"                      '온아티 검진
                Call GetWorkList_ONITGUM(pFrom, pTo, SPD)

        Case "ONITEMR"                      '온아티 EMR
                Call GetWorkList_ONITEMR(pFrom, pTo, SPD)

        Case "PLIS"                         '포미스 슈바이처
                Call GetWorkList_PLIS(pFrom, pTo, SPD)

        Case "SY"                           'SY
                Call GetWorkList_SY(Format(pFrom, "yyyy-mm-dd"), Format(pTo, "yyyy-mm-dd"), SPD)
        
        Case "TWIN"                         '투윈정보
                Call GetWorkList_TWIN(pFrom, pTo, SPD)

        Case "UBCARE"                       '의사랑
                Call GetWorkList_UBCARE(pFrom, pTo, SPD)

'        Case "WELL"                         '웰커머스
'                Call GetWorkList_WELL(pFrom, pTo, SPD)

'        Case "ONIT"
'            Call GetWorkList_onit(pFrom, pTo, SPD)

'        Case "PLIS"
'            Call GetWorkList_PLIS(pFrom, pTo, SPD)
        Case Else
        
        
    End Select

End Sub

'-- 워크리스트 조회
Public Sub GetWorkList_EONM(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       O.H141_ODRDAT        AS HOSPDATE " & vbCr
    SQL = SQL & "      ,O.H141_TSAMPLENO     AS BARCODE  " & vbCr
    SQL = SQL & "      ,O.H141_SEQNO         AS PID      " & vbCr
    SQL = SQL & "      ,P.A110_CHARTNO       AS CHARTNO  " & vbCr
    SQL = SQL & "      ,P.A110_PATNM         AS PNAME    " & vbCr
    SQL = SQL & "      ,P.A110_JUMIN1        AS AGE      " & vbCr
    SQL = SQL & "      ,P.A110_SEX           AS SEX      " & vbCr
    SQL = SQL & "      ,COUNT(O.H141_SUGACD) AS CNT      " & vbCr
    SQL = SQL & "  FROM TB_H141_LISTAKEBODY O, TB_A110_PATINFO P                " & vbCr
    SQL = SQL & " Where O.H141_ODRDAT BETWEEN '" & pFrom & "' AND '" & pTo & "' " & vbCr
    SQL = SQL & "   AND P.A110_ChartNo = O.H141_CHARTNO                         " & vbCr
    SQL = SQL & "   AND O.H141_NOTYYN  = 'N'                                    " & vbCr
    SQL = SQL & "   And O.H141_SUGACD IN (" & gAllTestCd & ")                   " & vbCr
    SQL = SQL & " Group By O.H141_ODRDAT,O.H141_TSAMPLENO,O.H141_SEQNO,P.A110_CHARTNO,P.A110_PATNM,P.A110_JUMIN1,P.A110_SEX " & vbCr
    SQL = SQL & " Order By O.H141_ODRDAT, O.H141_SEQNO"
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'-- 워크리스트 조회
Public Sub GetWorkList_GINUS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    SQL = ""
    SQL = SQL & "SELECT /*+ INDEX (coif scccoifm_ix1) INDEX (prex scrprexh_ix3) INDEX (ptbs pmcptbsm_ux1) INDEX (rslt scrrslth_ux1) INDEX (xpsl mosxpslh_ix2) */" & vbCr
    SQL = SQL & "       prex.acp_dt                                             AS HOSPDATE     " & vbCr
    SQL = SQL & "     , prex.smp_no                                             AS BARCODE      " & vbCr
    SQL = SQL & "     , coif.exam_mach_cd                                                       " & vbCr
    SQL = SQL & "     , rslt.exam_stus                                                          " & vbCr
    SQL = SQL & "     , prex.pt_no                                              AS PID          " & vbCr
    SQL = SQL & "     , ptbs.pt_nm                                              AS PNAME        " & vbCr
    SQL = SQL & "     , ptbs.ssn_1                                                              " & vbCr
    SQL = SQL & "     , ptbs.ssn_2                                                              " & vbCr
    SQL = SQL & "     , DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd)    AS INOUT        " & vbCr
    SQL = SQL & "     , xpsl.adms_ymd                                                           " & vbCr
    SQL = SQL & "     , xpsl.mn_sub_typ_cd                                                      " & vbCr
    SQL = SQL & "     , xpsl.med_dpt_cd                                                         " & vbCr
    SQL = SQL & "     , xpsl.med_ymd                                                            " & vbCr
    SQL = SQL & "     , Max(Trim(coif.lmt_trm_day))                                             " & vbCr
    SQL = SQL & "  FROM scrprexh prex                                                           " & vbCr
    SQL = SQL & "     , pmcptbsm ptbs                                                           " & vbCr
    SQL = SQL & "     , scccoifm coif                                                           " & vbCr
    SQL = SQL & "     , mosxpslh xpsl                                                           " & vbCr
    SQL = SQL & "     , scrrslth rslt                                                           " & vbCr
    SQL = SQL & " WHERE SUBSTR(prex.acp_dt,1,8) BETWEEN '" & pFrom & "' AND '" & pTo & "'       " & vbCr
    SQL = SQL & "   AND prex.hos_org_no    = '" & gHOSP.HOSPCD & "'                             " & vbCr
    SQL = SQL & "   AND coif.exam_mach_cd  = '" & gHOSP.MACHCD & "'                             " & vbCr
    SQL = SQL & "   AND rslt.hos_org_no    = prex.hos_org_no                                    " & vbCr
    SQL = SQL & "   AND rslt.smp_no        = prex.smp_no                                        " & vbCr
    SQL = SQL & "   AND rslt.prcp_seq      = prex.prcp_seq                                      " & vbCr
    SQL = SQL & "   AND rslt.exam_seq      = prex.exam_seq                                      " & vbCr
    SQL = SQL & "   AND ptbs.hos_org_no    = prex.hos_org_no                                    " & vbCr
    SQL = SQL & "   AND ptbs.pt_no         = prex.pt_no                                         " & vbCr
    SQL = SQL & "   AND coif.hos_org_no    = prex.hos_org_no                                    " & vbCr
    SQL = SQL & "   AND coif.exam_cd       = prex.cd                                            " & vbCr
    SQL = SQL & "   AND xpsl.smp_no        = prex.smp_no                                        " & vbCr
    SQL = SQL & "   AND xpsl.hos_org_no    = prex.hos_org_no                                    " & vbCr
    SQL = SQL & "   AND coif.use_typ       = 'Y'                                                " & vbCr
    SQL = SQL & "   AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt               " & vbCr
    SQL = SQL & "   AND xpsl.prcp_typ_cd  IN ('O','C')                                          " & vbCr
    SQL = SQL & "   AND rslt.exam_stus    IN ('0')                                              " & vbCr
    SQL = SQL & "   GROUP BY prex.acp_dt, prex.smp_no, coif.exam_mach_cd ,rslt.exam_stus,       "
    SQL = SQL & "            prex.pt_no, ptbs.pt_nm, ptbs.ssn_1, ptbs.ssn_2,                    "
    SQL = SQL & "            DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd),              "
    SQL = SQL & "            xpsl.adms_ymd,xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd    " & vbCr
    SQL = SQL & "   ORDER BY prex.acp_dt, prex.smp_no                                           " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        
        SPD.MaxRows = 0
        
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                iCnt = iCnt + 1
                
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
                    
                    Select Case Trim(RS.Fields("INOUT"))
                        Case "O": SetText SPD, "외래", intRow, colINOUT
                        Case "E": SetText SPD, "응급", intRow, colINOUT
                        Case "I": SetText SPD, "입원", intRow, colINOUT
                    End Select
                    SetText SPD, CStr(iCnt), intRow, colOCNT

                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_EASYS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.ACC_YMD       AS HOSPDATE     " & vbCr
    SQL = SQL & "     , a.RECEPT_NO     AS BARCODE      " & vbCr
    SQL = SQL & "     , a.PTNT_NO       AS PID          " & vbCr
    SQL = SQL & "     , c.PTNT_NM       AS PNAME        " & vbCr
    SQL = SQL & "     , c.BIRTH_YMD     AS AGE          " & vbCr
    SQL = SQL & "     , c.SEX           AS SEX          " & vbCr
    SQL = SQL & "     , COUNT(a.ORD_CD) AS CNT          " & vbCr
    SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c            " & vbCr
    SQL = SQL & " WHERE a.ACC_YMD between '" & pFrom & "' AND '" & pTo & "' " & vbCr
    SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ")                    " & vbCr
    SQL = SQL & "   AND a.STS_CD     = 'A'                                  " & vbCr 'A:접수, R:결과전송
    SQL = SQL & "   AND a.SUTAK_CD   = ''                                   " & vbCr
    SQL = SQL & "   AND a.RECEPT_NO  = b.RECEPT_NO                          " & vbCr
    SQL = SQL & "   AND a.PTNT_NO    = c.PTNT_NO                            " & vbCr
    SQL = SQL & " GROUP BY a.ACC_YMD, a.RECEPT_NO, a.PTNT_NO, c.PTNT_NM,c.BIRTH_YMD,c.SEX " & vbCr
    SQL = SQL & " ORDER BY a.ACC_YMD, a.PTNT_NO " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_JWINFO(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.RECEIPTDATE    AS HOSPDATE " & vbCr
    SQL = SQL & "     , a.SPECIMENNUM    AS BARCODE  " & vbCr
    SQL = SQL & "     , a.RECEIPTNO      AS CHARTNO  " & vbCr
    SQL = SQL & "     , a.IPDOPD         AS INOUT    " & vbCr
    SQL = SQL & "     , a.PTNO           AS PID      " & vbCr
    SQL = SQL & "     , a.SNAME          AS PNAME    " & vbCr
    SQL = SQL & "     , COUNT(a.LABCODE) AS CNT      " & vbCr
    SQL = SQL & "   FROM SLA_LabMaster a, SLA_LabResult b   " & vbCr
    SQL = SQL & " WHERE a.RECEIPTNO     = b.RECEIPTNO       " & vbCr
    SQL = SQL & "   AND a.ORDERCODE     = b.ORDERCODE       " & vbCr
    SQL = SQL & "   and a.RECEIPTDATE   = b.RECEIPTDATE     " & vbCr
    SQL = SQL & "   AND a.SPECIMENNUM   = b.SPECIMENNUM     " & vbCr
    SQL = SQL & "   AND a.RECEIPTDATE BETWEEN '" & Format(pFrom, "####-##-##") & "' and '" & Format(pTo, "####-##-##") & "'" & vbCr
    SQL = SQL & "   AND b.LABCODE IN (" & gAllTestCd & ")   " & vbCr
    SQL = SQL & "   AND a.JSTATUS < '3'                     " & vbCr
    SQL = SQL & "   AND (b.Result = '' OR LTRIM(RTRIM(b.Result)) IS NULL)" & vbCr
    SQL = SQL & " GROUP BY a.RECEIPTDATE, a.SPECIMENNUM, a.RECEIPTNO, a.IPDOPD, a.PTNO, a.SNAME " & vbCr
    SQL = SQL & " ORDER BY a.RECEIPTDATE,a.SPECIMENNUM "
        
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_JAINCOM(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    SQL = ""
    SQL = SQL & "SELECT DiSTINCT "
    SQL = SQL & "       b.SCP42JDATE            as HOSPDATE         " & vbCr
    SQL = SQL & "     , a.SCP41SPMNO2           as BARCODE          " & vbCr
    SQL = SQL & "     , b.SCP42IDNOA            as PID              " & vbCr
    SQL = SQL & "     , a.SCP41NAME             as PNAME            " & vbCr
    SQL = SQL & "     , a.SCP41SEX              as SEX              " & vbCr
    SQL = SQL & "     , a.SCP41BIRTH            as AGE              " & vbCr
    SQL = SQL & "     , COUNT(b.SCP42SUGACD)    as CNT              " & vbCr
    SQL = SQL & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b    " & vbCr
    SQL = SQL & " WHERE a.SCP41PCODE    = b.SCP42PCODE              " & vbCr
    SQL = SQL & "   AND a.SCP41JDATE    = b.SCP42JDATE              " & vbCr
    SQL = SQL & "   AND a.SCP41SID      = b.SCP42SID                " & vbCr
    SQL = SQL & "   AND a.SCP41SPMNO2   = b.SCP42SPMNO2             " & vbCr
    SQL = SQL & "   AND b.SCP42JDATE BETWEEN '" & pFrom & "' AND '" & pTo & "'                              " & vbCr
    SQL = SQL & "   AND b.SCP42SUGACD  IN (" & gAllTestCd & ")                                              " & vbCr
    SQL = SQL & "   AND (b.SCP42RESULT IS NULL OR b.SCP42RESULT = '')                                       " & vbCr
    SQL = SQL & " GROUP BY b.SCP42JDATE, a.SCP41SPMNO2, b.SCP42IDNOA, a.SCP41NAME, a.SCP41SEX, a.SCP41BIRTH " & vbCr
    SQL = SQL & " ORDER BY b.SCP42JDATE, a.SPECIMENNUM                                                      " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


Public Sub GetWorkList_KCHART(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    'SQL = SQL & "    AND L.검사종류 = '" & gHOSP.LABCD & "'" & vbCr
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       J.접수일자          AS HOSPDATE " & vbCr
    SQL = SQL & "     , L.검체번호          AS BARCODE  " & vbCr
    SQL = SQL & "     , A.챠트번호          AS CHARTNO  " & vbCr
    SQL = SQL & "     , J.접수번호          AS PID      " & vbCr
    SQL = SQL & "     , A.환자이름          AS PNAME    " & vbCr
    SQL = SQL & "     , A.환자성별          AS SEX      " & vbCr
    SQL = SQL & "     , A.환자나이          AS AGE      " & vbCr
    'SQL = SQL & "     , L.진료검사ID        AS R        " & vbCr
    'SQL = SQL & "     , L.진료지원ID        AS P        " & vbCr
    SQL = SQL & "     , COUNT(L.처방코드)   AS CNT      " & vbCr
    SQL = SQL & "  FROM             TB_진료검사 L                                    " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료지원 J ON (L.진료지원ID = J.진료지원ID)   " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료일반 A ON (J.진료일자   = A.진료일자      " & vbCr
    SQL = SQL & "                                AND  J.챠트번호   = A.챠트번호      " & vbCr
    SQL = SQL & "                                AND  J.진료번호   = A.진료번호)     " & vbCr
    SQL = SQL & " Where J.접수일자 BETWEEN '" & Format(pFrom, "####-##-##") & "' and '" & Format(pTo, "####-##-##") & "'" & vbCr
    SQL = SQL & "   AND L.검사상태 < 5                                     " & vbCr
    SQL = SQL & "   AND L.처방코드 + L.서브코드 IN (" & gAllTestCd & ")    " & vbCr
    SQL = SQL & " GROUP BY J.접수일자, L.검체번호, A.챠트번호, J.접수번호, A.환자이름, A.환자성별, A.환자나이 " & vbCr
    SQL = SQL & " ORDER BY J.접수일자, J.접수번호                          " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_KCWH(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestNmS  As String
    
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim Prm3 As New ADODB.Parameter
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    strTestNmS = ""
    
'    SQL = ""
'    SQL = SQL & "SELECT "
'    SQL = SQL & "  to_char(b.ACPTDT, 'yyyy-MM-dd hh24:mi')                  as HOSPDATE " & vbCr    ' 처방일자
'    SQL = SQL & ", fn_lab_get_prtbcno_from_bcno(a.BCNO)                     as BARCODE  " & vbCr    ' 검체번호[바코드 번호]
'    SQL = SQL & ", b.WKGRPCD||'-'||b.WKYMD||'-'||lpad(b.WKSEQ, 4, '0')      as CHARTNO  " & vbCr    ' LAB 번호
'    SQL = SQL & ", b.PT_NO                                                  as PID      " & vbCr    ' 환자번호
'    SQL = SQL & ", b.PATNM                                                  as PNAME    " & vbCr    ' 환자명
'    SQL = SQL & ", b.SEX                                                    as SEX      " & vbCr    ' 성별
'    SQL = SQL & ", b.AGE                                                    as AGE      " & vbCr    ' 나이
'    SQL = SQL & ", f.CITIZEN1||'-'||substr(f.CITIZEN2, 1, 1) || '******'    as PJUMIN   " & vbCr    ' 주민번호
'    SQL = SQL & ", b.IOCLS                                                  as INOUT    " & vbCr    ' 입외구분
'    SQL = SQL & ", a.SPCCD                                                  as SPCCD    " & vbCr    ' 검체코드
'    SQL = SQL & ", fn_lab_get_spcnmd(a.SPCCD)                               as SPCNM    " & vbCr    ' 검체명
'    SQL = SQL & ", c.OTCLSCD                                                as ORDCD    " & vbCr    ' 처방코드 ??
'    SQL = SQL & ", a.TCLSCD                                                 as ITEM     " & vbCr    ' 검사코드
'    SQL = SQL & ", d.TNMD                                                   as ITEMNM   " & vbCr    ' 검사명
'    'SQL = SQL & ", b.ORDDRNM "                                                                      ' 의뢰의사
'    'SQL = SQL & ", b.DPNM    "                                                                      ' 부서명
'    'SQL = SQL & ", b.WARDNM  "                                                                      ' 병동명
'    'SQL = SQL & ", b.DPCD    "                                                                      ' 처방처
'    SQL = SQL & "  FROM SLRTSTMT a"
'    SQL = SQL & "     , SLCINFMT b"
'    SQL = SQL & "     , SLCORDMT c"
'    SQL = SQL & "     , SLFITEMT d"
'    SQL = SQL & "     , SLFITEMT i"
'    SQL = SQL & "     , SLFTIDMT e"
'    SQL = SQL & "     , APPATBAT f" & vbCr
'    SQL = SQL & " WHERE a.BCNO      = b.BCNO                " & vbCr
'    SQL = SQL & "   AND b.BCNO      = c.BCNO                " & vbCr
'    SQL = SQL & "   AND a.OTCLSCD   = c.OTCLSCD             " & vbCr
'    SQL = SQL & "   AND a.SPCCD     = c.SPCCD               " & vbCr
'    SQL = SQL & "   AND b.SPCFLAG   = '" & gHOSP.LABCD & "' " & vbCr  'vc_Acpt
'    SQL = SQL & "   AND a.RSTFLAG   IN ( 'L', 'R')          " & vbCr
'    SQL = SQL & "   AND a.TCLSCD    = d.TCLSCD              " & vbCr
'    SQL = SQL & "   AND c.OTCLSCD   = i.TCLSCD              " & vbCr
'    SQL = SQL & "   AND a.TCLSCD    = e.TCLSCD              " & vbCr
'    SQL = SQL & "   AND a.SPCCD     = e.SPCCD               " & vbCr
'    SQL = SQL & "   AND e.USDT      <= a.RLCOLLDT           " & vbCr
'    SQL = SQL & "   AND e.UEDT      > a.RLCOLLDT            " & vbCr
'    SQL = SQL & "   AND b.PT_NO     = f.PT_NO               " & vbCr
'    SQL = SQL & "   AND a.WKGRPCD   = E.WKGRPCD             " & vbCr
'    SQL = SQL & "   AND e.WKGRPCD   = '" & gHOSP.PARTCD & "'" & vbCr  '-- in_GrpCode : 그룹코드
''    SQL = SQL & "   AND a.WKYMD BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
'    SQL = SQL & "   AND a.WKYMD >= to_date(" & pFrom & ", 'YYYYmmdd')       " & vbCr     '-- 처방 검색 시작일자"
'    SQL = SQL & "   AND a.WKYMD <= to_date(" & pFrom & ", 'YYYYmmdd') + 1.0 " & vbCr     '-- 처방 검색 종료일자"
'    SQL = SQL & "   AND a.TCLSCD    IN (" & gAllTestCd & ") " & vbCr
'    SQL = SQL & " ORDER BY HOSPDATE, CHARTNO "
'
'    Call SetSQLData("워크조회", SQL)
    
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    With AdoCmd
        .CommandTimeout = 15
        .CommandText = "PKG_SUP_LAB_INTERFACE.pc_DownLoad_Order"
        .CommandType = adCmdStoredProc

        Set Prm1 = .CreateParameter("in_SDate", adVarChar, adParamInput, 100, pFrom)
        .Parameters.Append Prm1
        Set Prm2 = .CreateParameter("in_EDate", adVarChar, adParamInput, 100, pTo)
        .Parameters.Append Prm2
        Set Prm3 = .CreateParameter("in_GrpCode", adVarChar, adParamInput, 100, gHOSP.PARTCD)
        .Parameters.Append Prm3
    End With

    '-- SP 사용
    Set RS = New ADODB.Recordset
    RS.Open AdoCmd.Execute

    If Not RS.EOF = True And Not RS.BOF = True Then
        
        SPD.MaxRows = 0
        
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("ORDDT")) = strHospDate And Trim(RS("BCNO")) = strBarcode Then
                        blnSame = True
                        iCnt = 0
                    End If
                Next
                
                strTestNmS = strTestNmS & Trim(RS.Fields("TNMD") & "") & "/"
                
                If intRow > 0 Then
                    SetText SPD, strTestNmS, intRow, colITEMS
                End If
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                    iCnt = iCnt + 1
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("ORDDT")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BCNO")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("LABNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PT_NO")) & "", intRow, colPID
                    'SetText SPD, Trim(RS.Fields("INOUT")) & "", intRow, colINOUT
                    SetText SPD, Trim(RS.Fields("PATNAME")) & "", intRow, colPNAME
                    'SetText SPD, Trim(RS.Fields("PJUMIN")) & "", intRow, colPJUMIN
                    'SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    'SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, iCnt, intRow, colOCNT
                    
                    If intRow > 1 Then
                        strTestNmS = ""
                    End If
                    
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_HWASAN(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    ' O.Sex         성별('M':남, 'F':여, 'A':모두, 'E':기타)
    ' O.StatFg      응급여부('0':아님, '1':응급)
    ' O.AllTestNm   모든 검사명
'    SQL = SQL & "     , O.SpcCd                         " & vbCr
'    SQL = SQL & "     , O.SpcNm                         " & vbCr
'    SQL = SQL & "     , O.AllTestNm                     " & vbCr
'    SQL = SQL & "     , O.StatFg                        " & vbCr
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       O.ORDDT             as HOSPDATE " & vbCr
    SQL = SQL & "     , O.SPCNO             as BARCODE  " & vbCr
    SQL = SQL & "     , O.PTID              as PID      " & vbCr
    SQL = SQL & "     , O.PTNM              as PNAME    " & vbCr
    SQL = SQL & "     , O.SEX               as SEX      " & vbCr
    SQL = SQL & "     , O.AGE               as AGE      " & vbCr
    SQL = SQL & "     , COUNT(T.TESTCD)     as CNT      " & vbCr
    SQL = SQL & "  FROM TC201 O, TC301 T                " & vbCr
    SQL = SQL & " WHERE O.SPCNO = T.SPCNO               " & vbCr
    SQL = SQL & "   AND O.OrdDt between  '" & pFrom & "' and '" & pTo & "'   " & vbCr
    SQL = SQL & "   And T.TESTCD in (" & gAllTestCd & ")" & vbCr
    SQL = SQL & " GROUP BY O.ORDDT,O.SPCNO,O.PTID,O.PTNM,O.SEX,O.AGE " & vbCr
    SQL = SQL & " Order By O.ORDDT,O.SPCNO              " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


Public Sub GetWorkList_KOMAIN(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim sqlRet      As Integer
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    iCnt = 0
    
    If gHOSP.BARUSE = "Y" Then
        '바코드 사용
        SQL = "EXEC AP_INF_BAR_ORDER '" & gHOSP.MACHCD & "','" & pFrom & "','" & pTo & "'"
        '반환값 [EMRLIS2] : r.BCID, r.Hcode, r.Serial, c.PtName, r.Orderdate, ErYn
    Else
        '바코드 미사용  yyyy-mm-dd
        SQL = "EXEC AP_INF_S_ORDER '" & gHOSP.MACHCD & "','0','" & pFrom & "','" & pTo & "'"
        '반환값 [EMRLIS2] : LID,Hcode,Serial,PtName, Orderdate, ROrder, ErYN, Age, Sex,DeptM
    End If
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = New ADODB.Recordset
    RS.Open AdoCn.Execute(SQL, sqlRet)
    
    Call SetSQLData("워크조회", SQL)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        SPD.MaxRows = 0
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                iCnt = iCnt + 1
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("ORDERDATE")) = strHospDate And Trim(RS("LID")) = strBarcode Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("ORDERDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("LID")) & "", intRow, colBARCODE
                    'SetText SPD, Trim(RS.Fields("HCODE")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("SERIAL")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("RORDER")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PTNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, CStr(iCnt), intRow, colOCNT
                    SetText SPD, GetSampleITEM_SP(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'   -- 워크리스트조회
'   @StatustIndex   tinyint,          // 필수입력: 0-전체, 1-미완료, 2-완료
'   @WorkListCode   varchar(50),      // 필수입력: 워크리스트코드
'   @BeginDate      smalldatetime,    // 필수입력: 조회일-시작
'   @EndDate        smalldatetime,    // 필수입력: 조회일-끝
'   @BeginNo        int,              // 선택입력: 접수번호 시작 (기본값 : 0)
'   @EndNo          int,              // 선택입력: 접수번호 종료 (기본값 : 0)
'   @TestCodes      varchar(200)      // 선택입력: 검사코드

Public Sub GetWorkList_SY(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim sqlRet      As Integer
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    iCnt = 0
    
    SQL = ""
    SQL = SQL = "EXEC Interface_GetPatientResultList02 "
    SQL = SQL & "   '0'"
    SQL = SQL & " , '" & gHOSP.PARTCD & "'"
    SQL = SQL & " , '" & pFrom & "'"
    SQL = SQL & " , '" & pTo & "'"
    SQL = SQL & " , 0"
    SQL = SQL & " , 0"
    SQL = SQL & " , ''"       'gAllTestCd
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = New ADODB.Recordset
    RS.Open AdoCn.Execute(SQL, sqlRet)
    
    Call SetSQLData("워크조회", SQL)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        SPD.MaxRows = 0
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                iCnt = iCnt + 1
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("LabRegDate")) = strHospDate And Format(Trim(RS.Fields("LabRegDate")), "yymmdd") & PedLeftStr(Trim(RS.Fields("LabRegNo")), 5, "0") = strBarcode Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("LabRegDate")) & "", intRow, colHOSPDATE
                    SetText SPD, Format(Trim(RS.Fields("LabRegDate")), "yymmdd") & PedLeftStr(Trim(RS.Fields("LabRegNo")), 5, "0"), intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PatientChartNo")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("CompanyCode")) & "", intRow, colINOUT
                    SetText SPD, Trim(RS.Fields("LabRegNo")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PatientName")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("PatientSex")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("PatientAge")) & "", intRow, colPSEX
                    SetText SPD, CStr(iCnt), intRow, colOCNT
                    'SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


Public Sub GetWorkList_MSINFOTEC(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    'SQL = SQL & "     , a.OIFL          AS IO"
    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       a.ORDT          AS HOSPDATE " & vbCr
    SQL = SQL & "     , a.SPNO          AS BARCODE  " & vbCr
    SQL = SQL & "     , a.PAID          AS PID      " & vbCr
    SQL = SQL & "     , a.NWNO          AS CHARTNO  " & vbCr
    SQL = SQL & "     , b.PANM          AS PNAME    " & vbCr
    SQL = SQL & "     , b.SEXS          AS SEX      " & vbCr
    SQL = SQL & "     , b.AGES          AS AGE      " & vbCr
    SQL = SQL & "     , COUNT(a.ORCD)   AS CNT      " & vbCr
    SQL = SQL & "  From LRESULT a, APATINF b        " & vbCr
    SQL = SQL & " Where a.ORDT between  '" & pFrom & "' and '" & pTo & "'   " & vbCr
    SQL = SQL & "   And a.PAID = b.PAID                                     " & vbCr
    SQL = SQL & "   And a.ORCD IN (" & gAllTestCd & ")                      " & vbCr
    SQL = SQL & "   And a.OKFL <> 'Y'                                       " & vbCr   '-- 결과확정유무
    SQL = SQL & " GROUP BY a.ORDT,a.SPNO,a.PAID,a.NWNO,b.PANM,b.SEXS,b.AGES " & vbCr
    SQL = SQL & " Order By a.ORDT,a.PAID,b.PANM                             " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_MOD(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    
    SQL = ""
    SQL = SQL & "Select Distinct "
    SQL = SQL & "     , a.REQDATE           AS HOSPDATE     " & vbCr
    SQL = SQL & "     , c.SPECIMENID        AS BARCODE      " & vbCr
    SQL = SQL & "       a.PID               AS PID          " & vbCr
    SQL = SQL & "     , a.IOFLAG            AS IO           " & vbCr
    SQL = SQL & "     , b.PAT_NM            AS PNAME        " & vbCr
    SQL = SQL & "     , COUNT(c.EXAMCODE)   AS CNT          " & vbCr
    SQL = SQL & "  From EXAMREQ a, TI_PAT b, EXAMRES c      " & vbCr
    SQL = SQL & " Where a.PID       = b.PAT_CHART           " & vbCr
    SQL = SQL & "   And a.PID       = c.PID                 " & vbCr
    SQL = SQL & "   And a.SEQNO     = c.SEQNO               " & vbCr
    SQL = SQL & "   And a.RECENO    = c.RECENO              " & vbCr
    SQL = SQL & "   And a.REQDATE Between '" & pFrom & "' And '" & pTo & "' " & vbCr
    SQL = SQL & "   And c.EXAMCODE in (" & gAllTestCd & ")                  " & vbCr
    SQL = SQL & "   And (c.EXAMEND  = '' Or c.EXAMEND IS NULL)              " & vbCr
    SQL = SQL & " GROUP BY a.REQDATE,c.SPECIMENID,a.PID,a.IOFLAG,b.PAT_NM   " & vbCr
    SQL = SQL & " Order By a.REQDATE,c.SPECIMENID                           " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    Select Case Trim(RS.Fields("IO"))
                        Case "1": SetText SPD, "외래", intRow, colINOUT
                        Case "2": SetText SPD, "입원", intRow, colINOUT
                    End Select
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_MEDITOLISS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strJumin    As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       A.REQUEST_DATE          AS HOSPDATE                     " & vbCr
    SQL = SQL & "     , A.EXAM_NO               AS BARCODE                      " & vbCr
    SQL = SQL & "     , A.CHART_NO              AS CHARTNO                      " & vbCr
    SQL = SQL & "     , A.PERSON_NAME           AS PNAME                        " & vbCr
    SQL = SQL & "     , A.PERSONAL_ID           AS JUMIN                        " & vbCr
    SQL = SQL & "     , COUNT(B.EXAM_CODE)      AS CNT                          " & vbCr
    SQL = SQL & "  FROM MEDITOLISS..TOTAL A, MEDITOLISS..TOTRES B               " & vbCr
    SQL = SQL + " WHERE A.REQUEST_DATE Between '" & pFrom & "' And '" & pTo & "'" & vbCr
    SQL = SQL & "   And B.EXAM_CODE     IN (" & gAllTestCd & ")                 " & vbCr
    SQL = SQL & "   AND B.EXAM_PART     = '" & gHOSP.PARTCD & "'                " & vbCr    'C:생화학
    SQL = SQL & "   AND B.RESULT_VALUE  = ''                                    " & vbCr
    SQL = SQL & "   AND A.REQUEST_DATE  = B.REQUEST_DATE                        " & vbCr
    SQL = SQL & "   AND A.EXAM_NO       = B.EXAM_NO                             " & vbCr
    SQL = SQL & " GROUP BY A.REQUEST_DATE, A.EXAM_NO, A.CHART_NO, A.PERSON_NAME, A.PERSONAL_ID" & vbCr
    SQL = SQL & " ORDER BY A.REQUEST_DATE, A.EXAM_NO                            " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("JUMIN")) & "", intRow, colPJUMIN
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    strJumin = Trim(RS.Fields("JUMIN")) & ""
                    Call CalAgeSex(strJumin, Format(Date, "yyyy/mm/dd"))
                    SetText SPD, mPatient.AGE, intRow, colPAGE
                    SetText SPD, mPatient.SEX, intRow, colPSEX
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_MEDICHART(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       (a.진료년 + a.진료월 + a.진료일)    AS HOSPDATE     " & vbCr
    SQL = SQL & "     , a.챠트번호                          AS CHARTNO      " & vbCr
    SQL = SQL & "     , c.진료상태                          AS STATE        " & vbCr
    SQL = SQL & "     , b.수진자명                          AS PNAME        " & vbCr
    SQL = SQL & "     , b.주민등록번호                      AS PJUMIN       " & vbCr
    SQL = SQL & "     , COUNT(a.처방코드)                   AS CNT          " & vbCr
    SQL = SQL & "  From TB_검사항목 a, TB_인적사항 b, TB_진료기본 c         " & vbCr
    SQL = SQL & " Where (a.진료년 + a.진료월 + a.진료일) >= '" & pFrom & "' " & vbCr
    SQL = SQL & "   And (a.진료년 + a.진료월 + a.진료일) <= '" & pTo & "'   " & vbCr
    SQL = SQL & "   And a.처방번호 > 0                                      " & vbCr
    SQL = SQL & "   And c.진료상태 IN ('1','5','6','7','8','9')             " & vbCr
    SQL = SQL & "   And (a.처방코드 + a.서브코드) IN (" & gAllTestCd & ")   " & vbCr
    SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCr
    SQL = SQL & "   And a.진료년    = c.진료년                              " & vbCr
    SQL = SQL & "   And a.진료월    = c.진료월                              " & vbCr
    SQL = SQL & "   And a.진료일    = c.진료일                              " & vbCr
    SQL = SQL & "   And a.챠트번호  = c.챠트번호                            " & vbCr
    SQL = SQL & "   And a.챠트번호  = b.챠트번호                            " & vbCr
    SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCr
    SQL = SQL & " GROUP BY HOSPDATE, a.챠트번호, c.진료상태, b.수진자명, b.주민등록번호" & vbCr
    SQL = SQL & " Order By a.진료년, a.진료월, a.진료일, b.수진자명         " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_MEDIIT(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strPID      As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       P.request_date      AS HOSPDATE " & vbCr
    SQL = SQL & "     , P.exam_no           AS PID      " & vbCr
    SQL = SQL & "     , P.company_code      AS INOUT    " & vbCr
    SQL = SQL & "     , P.chart_no          AS CHARTNO  " & vbCr
    SQL = SQL & "     , p.person_name       AS PNAME    " & vbCr
    SQL = SQL & "     , P.person_sex        AS SEX      " & vbCr
    SQL = SQL & "     , P.person_age        AS AGE      " & vbCr
    SQL = SQL & "     , COUNT(R.pro_code)   AS CNT      " & vbCr
    SQL = SQL & "  FROM trust P, trures R               " & vbCr
    SQL = SQL & " WHERE P.request_date BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
    SQL = SQL & "   AND R.pro_code      IN (" & gAllTestCd & ") " & vbCr
    SQL = SQL & "   AND R.exam_code     <> 'X999'               " & vbCr
    SQL = SQL & "   AND P.request_date  = R.request_date        " & vbCr
    SQL = SQL & "   AND P.exam_no       = R.exam_no             " & vbCr
    SQL = SQL & " GROUP BY P.request_date, P.exam_no, P.company_code, P.chart_no, p.person_name, P.person_sex, P.person_age" & vbCr
    SQL = SQL & " ORDER BY P.request_date, P.exam_no            " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    strPID = GetText(SPD, i, colPID)
                    If Trim(RS("HOSPDATE")) = strHospDate And Mid(Trim(RS.Fields("HOSPDATE")), 3, 6) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0") = strPID Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0"), intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("INOUT")) & "", intRow, colINOUT
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'1 차결과마스터 (M01007TB1)
'gul_date    검사일  datetime(8)
'gul_bun_no  검체번호    smallint(2)
'gul_gum_code    검사항목코드    nvarchar(4)
'gul_value 결과값
'
'2 차결과마스터 (M01007TB3)
'gul2_date 검사일
'gul2_bun_no 검체번호
'gul2_gum_code 검사항목코드
'gul2_value 결과값

Public Sub GetWorkList_WELL(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    

    SQL = ""
    SQL = SQL & "SELECT DSITINCT"
    SQL = SQL & "       a.per1_date"
    SQL = SQL & "     , a.per1_mem_id"
    SQL = SQL & "     , a.per1_bun_no"
    SQL = SQL & "     , a.per1_name "
    SQL = SQL & "  FROM M01002TB1 a, M01007TB1 b"
    SQL = SQL & " WHERE a.KEYFIELD = b.GUL_KEY"
    SQL = SQL & "   AND a.per1_date Between '" & pFrom & "' And '" & pTo & "'" & vbCr
    SQL = SQL & "   AND b.gul_gum_code = 'M006' " & vbCr
    SQL = SQL & "   AND a.per1_date = b.gul_date"
    SQL = SQL & " Order By a.per1_bun_no"
    
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_NEOSOFT(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.WORK_DATE         as HOSPDATE " & vbCr
    SQL = SQL & "     , a.CHAM_INDEX        as BARCODE  " & vbCr
    SQL = SQL & "     , a.MEDM_ID           as PID      " & vbCr
    SQL = SQL & "     , b.CHAM_NAME         as PNAME    " & vbCr
    SQL = SQL & "     , b.CHAM_SEX          as SEX      " & vbCr
    SQL = SQL & "     , b.CHAM_YY           as AGE      " & vbCr
    SQL = SQL & "     , '입원'              as IO       " & vbCr
    SQL = SQL & "     , COUNT(a.CODE)       as CNT      " & vbCr
    SQL = SQL & "  From E_ORDER..ORDER_IN" & Format(Now, "yyyy") & " a "
    SQL = SQL & "     , E_BASECODE..HP_CHAM                          b          " & vbCr
    SQL = SQL & " Where a.WORK_DATE between '" & pFrom & "' AND '" & pTo & "'   " & vbCr
    SQL = SQL & "   And a.CHAM_INDEX = b.CHAM_INDEX                             " & vbCr
    SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                          " & vbCr
    SQL = SQL & "   AND a.TRANS = '2'                                           " & vbCr
    SQL = SQL & " UNION ALL                                                     " & vbCr
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.WORK_DATE         as HOSPDATE " & vbCr
    SQL = SQL & "     , a.CHAM_INDEX        as BARCODE  " & vbCr
    SQL = SQL & "     , a.MEDM_ID           as PID      " & vbCr
    SQL = SQL & "     , b.CHAM_NAME         as PNAME    " & vbCr
    SQL = SQL & "     , b.CHAM_SEX          as SEX      " & vbCr
    SQL = SQL & "     , b.CHAM_YY           as AGE      " & vbCr
    SQL = SQL & "     , '외래'              as IO       " & vbCr
    SQL = SQL & "     , COUNT(a.CODE)       as CNT      " & vbCr
    SQL = SQL & "  From E_ORDER..ORDER_OUT" & Format(Now, "yyyy") & " a "
    SQL = SQL & "     , E_BASECODE..HP_CHAM                           b         " & vbCr
    SQL = SQL & " Where a.WORK_DATE between '" & pFrom & "' AND '" & pTo & "'   " & vbCr
    SQL = SQL & "   And a.CHAM_INDEX = b.CHAM_INDEX                             " & vbCr
    SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                          " & vbCr
    SQL = SQL & "   AND a.TRANS = '2'                                           " & vbCr
    SQL = SQL & " GROUP BY a.WORK_DATE, a.CHAM_INDEX, a.MEDM_ID, b.CHAM_NAME, b.CHAM_SEX, b.CHAM_YY, IO " & vbCr
    SQL = SQL & " ORDER BY a.WORK_DATE, IO, a.CHAM_INDEX                        " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("IO")) & "", intRow, colINOUT
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_PLIS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strGBarcode As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT"
    SQL = SQL & "       m.workarea                      " & vbCr
    SQL = SQL & "     , m.accdt AS HOSPDATE             " & vbCr
'    SQL = SQL & "     , m.accseq                        " & vbCr
'    SQL = SQL & "     , m.spcyy                         " & vbCr
'    SQL = SQL & "     , m.spcno                         " & vbCr
    SQL = SQL & "     , m.ptid AS PID                   " & vbCr
    SQL = SQL & "     , p.ptnm AS PNAME                 " & vbCr
    SQL = SQL & "     , m.rcvdt                         " & vbCr
    SQL = SQL & "     , m.rcvtm                         " & vbCr
    SQL = SQL & "     , COUNT(r.testcd) AS CNT          " & vbCr
    SQL = SQL & "  FROM plis..s2lab201 m                " & vbCr
    SQL = SQL & "     , his001_v p                      " & vbCr
    SQL = SQL & "     , plis..s2lab302 r                " & vbCr
    SQL = SQL & "     , plis..s2lab001 e                " & vbCr
    SQL = SQL & " WHERE SUBSTRING(m.accdt,1,8) BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
    SQL = SQL & "   AND m.ptid = p.ptid COLLATE Korean_Wansung_CS_AS                    " & vbCr
    SQL = SQL & "   AND r.testcd IN (" & gAllTestCd & ")    " & vbCr
    SQL = SQL & "   AND (r.vfydt IS NULL OR r.vfydt='')     " & vbCr
    SQL = SQL & "   AND m.workarea = r.workarea             " & vbCr
    SQL = SQL & "   AND m.accdt = r.accdt                   " & vbCr
    SQL = SQL & "   AND m.accseq = r.accseq                 " & vbCr
    SQL = SQL & "   AND r.testcd = e.testcd                 " & vbCr
'    SQL = SQL & "  Group by m.workarea, m.accdt, m.spcyy,m.spcno,m.accseq, m.ptid,p.ptnm,m.rcvdt, m.rcvtm "
    SQL = SQL & "  Group by m.workarea, m.accdt, m.ptid,p.ptnm,m.rcvdt, m.rcvtm " & vbCr
    SQL = SQL & "  Order by m.rcvdt, m.rcvtm                " & vbCr
        
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        
        SPD.MaxRows = 0
        
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                strGBarcode = Trim(RS("SPCYY")) & Format$(Trim(RS("SPCNO")), String$(9, "0"))

                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And strGBarcode = strBarcode Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, strGBarcode, intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_ONITGUM(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       PER_GUMJIN_DATE     AS HOSPDATE                             " & vbCr
    SQL = SQL & "     , PER_NAME            AS PNAME                                " & vbCr
    SQL = SQL & "     , PER_GUM_NUM         AS BARCODE                              " & vbCr
    SQL = SQL & "     , COUNT(EDPSCODE)     AS CNT                                  " & vbCr
    SQL = SQL & "  FROM ONIT..GUMJIN_INTERFACE                                      " & vbCr
    SQL = SQL & " WHERE PER_GUMJIN_DATE BETWEEN '" & pFrom & "' AND '" & pTo & "'   " & vbCr
    SQL = SQL & "   AND EDPSCODE IN (" & gAllTestCd & ")                            " & vbCr
    SQL = SQL & "   AND (RESULT = ''  OR RESULT IS NULL)                            " & vbCr
    SQL = SQL & " GROUP BY PER_GUMJIN_DATE, PER_NAME, PER_GUM_NUM AS BARCODE        " & vbCr
    SQL = SQL & " ORDER BY PER_GUMJIN_DATE, PER_GUM_NUM                             " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_ONITEMR(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.ENTERDATE         AS HOSPDATE     " & vbCr
    SQL = SQL & "     , b.WAITSEQNO         AS BARCODE      " & vbCr
    SQL = SQL & "     , a.CHARTNO           AS CHARTNO      " & vbCr
    SQL = SQL & "     , c.SUJINNAME         AS PNAME        " & vbCr
    SQL = SQL & "     , a.SUJINPART         AS INOUT        " & vbCr    '62:검진
    SQL = SQL & "     ,COUNT(b.MAP2SEQNO)   AS CNT          " & vbCr
    SQL = SQL & "  FROM " & gSQLDB.DB & "..WAITPRSNP a      " & vbCr
    SQL = SQL & "      ," & gSQLDB.DB & "..JUN370_RESULTTB b" & vbCr
    SQL = SQL & "      ," & gSQLDB.DB & "..PEWPRSNP c       " & vbCr
    SQL = SQL & "      ," & gSQLDB.DB & "..BAGMAP2PREF d    " & vbCr
    SQL = SQL & " WHERE a.ENTERDATE BETWEEN  '" & pFrom & "' AND '" & pTo & "' " & vbCr
    SQL = SQL & "    AND a.JUNDAL       = '" & gHOSP.HOSPCD & "'    " & vbCr        '370
    SQL = SQL & "    AND a.WAITSEQNO    = b.WAITSEQNO               " & vbCr
    SQL = SQL & "    AND a.CHARTNO      = c.CHARTNO                 " & vbCr
    SQL = SQL & "    AND d.LABNO        IN (" & gHOSP.LABCD & ")    " & vbCr   '4
    SQL = SQL & "    AND b.MAP2SEQNO    IN (" & gAllTestCd & ")     " & vbCr
    SQL = SQL & "    AND b.MAP2SEQNO    = d.MAP2SEQNO               " & vbCr
    SQL = SQL & "    AND (b.RESULT = '' OR b.RESULT IS NULL)        " & vbCr
    SQL = SQL & " GROUP BY a.ENTERDATE, b.WAITSEQNO, a.CHARTNO, c.SUJINNAME, a.SUJINPART" & vbCr
    SQL = SQL & " ORDER BY a.ENTERDATE, b.WAITSEQNO                 " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    If Trim(RS.Fields("INOUT")) & "" = "62" Then
                        SetText SPD, "검진", intRow, colINOUT
                    Else
                        SetText SPD, "진료", intRow, colINOUT
                    End If
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub LetEqpMaster(ByVal pEqpCD As String)

    SQL = ""
    SQL = SQL & "UPDATE EQPMASTER SET EQUIPCD = '" & pEqpCD & "'"
                          
    Call DBExec(AdoCn_Local, SQL)

End Sub

Public Sub GetWorkList_MCC(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestNmS  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       READING_YMD     AS HOSPDATE " & vbCr
    SQL = SQL & "     , BCODE_NO        AS BARCODE  " & vbCr
    SQL = SQL & "     , PTNT_NO         AS PID      " & vbCr
    SQL = SQL & "     , PTNT_NM         AS PNAME    " & vbCr
    SQL = SQL & "     , AGE             AS AGE      " & vbCr
    SQL = SQL & "     , SEX             AS SEX      " & vbCr
    SQL = SQL & "     , IO_GB           AS INOUT    " & vbCr
    SQL = SQL & "     , ORD_CD          AS ITEM      " & vbCr
    'SQL = SQL & "     , COUNT(ORD_CD)   AS CNT      " & vbCr
    SQL = SQL & "  FROM LIS_INTERFACE1_V            " & vbCr
    SQL = SQL & " WHERE READING_YMD between '" & pFrom & "' AND '" & pTo & "'   " & vbCr
    SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ")                          " & vbCr
    SQL = SQL & "   AND STS_CD = '0'                                            " & vbCr    '0 접수, 1:결과전송
    'SQL = SQL & " GROUP BY READING_YMD,BCODE_NO,PTNT_NO,PTNT_NM,AGE,SEX,IO_GB   " & vbCr
    SQL = SQL & " ORDER BY READING_YMD,PTNT_NO,BCODE_NO,ORD_CD                         " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                        strTestNmS = ""
                        iCnt = 0
                    End If
                Next
                
                If blnSame = False Then
                    strTestNmS = strTestNmS & GetTestNm(Trim(RS.Fields("ITEM")) & "", False) & "/"
                    iCnt = iCnt + 1
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    If Trim(RS.Fields("INOUT")) & "" = "10" Then
                        SetText SPD, "입원", intRow, colINOUT
                    Else
                        SetText SPD, "외래", intRow, colINOUT
                    End If
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    'SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, iCnt, intRow, colOCNT
                    'SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    SetText SPD, strTestNmS, intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_TWIN(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       B.JOBDATE                               AS HOSPDATE     " & vbCr
    SQL = SQL & "     , C.SPECNO                                AS BARCODE      " & vbCr
    SQL = SQL & "     , C.PTNO                                  AS CHARTNO      " & vbCr
    SQL = SQL & "     , C.JOBNO                                 AS PID          " & vbCr
    SQL = SQL & "     , DECODE(C.GBIO,'I','입원','O','외래')    AS IO           " & vbCr
    SQL = SQL & "     , C.SNAME                                 AS PNAME        " & vbCr
    SQL = SQL & "     , C.SEX                                   AS SEX          " & vbCr
    SQL = SQL & "     , C.AGE                                   AS AGE          " & vbCr
    'SQL = SQL & "     , COUNT(A.MASTERCODE)                     AS CNT          " & vbCr
    SQL = SQL & "     , COUNT(A.SUBCODE)                        AS CNT          " & vbCr
    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A                             " & vbCr
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_MASTER  B                             " & vbCr
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_SPECMST C                             " & vbCr
    SQL = SQL & " Where B.JOBDATE BETWEEN '" & pFrom & "' AND '" & pTo & "'     " & vbCr '작업일자
    SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'                     " & vbCr '장비코드
    SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")                    " & vbCr
    SQL = SQL & "   AND C.STATUS  <= '3'                                        " & vbCr '검사상태
    SQL = SQL & "   And C.SPECNO  = A.SPECNO                                    " & vbCr
    SQL = SQL & "   And A.MASTERCODE = B.MASTERCODE                             " & vbCr
    SQL = SQL & " GROUP By B.JOBDATE, C.SPECNO, C.PTNO, C.JOBNO, C.GBIO, C.SNAME, C.SEX, C.AGE " & vbCr
    SQL = SQL & " ORDER BY B.JOBDATE, C.SPECNO                                  " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("IO")) & "", intRow, colINOUT
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


Private Function GetWorkList_XML() As Variant
    Dim strPath     As String
    Dim strBuffer   As String
    Dim BufChar     As String
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim i           As Long
    Dim lngBufLen   As Long
    Dim TextLine
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    blnSameRecord = False
    
    GetWorkList_XML = ""
    
    '-- 오더파일명과 경로를 지정한다.
    strPath = "C:\UBCare\SINAI\IF\EXAMIF_IN.xml"
    
    Open strPath For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1)         ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine
    Loop
    
    Close #1 ' 파일을 닫습니다
 
    intIdx = 0
    lngBufLen = Len(strBuffer)
        
    For i = 1 To lngBufLen
        If intIdx = 0 Then
            BufChar = Mid$(strBuffer, i, 4)
        Else
            BufChar = Mid$(strBuffer, i + 3)
        End If
        
        If BufChar = "<검사>" Then
            intIdx = 1
            strTmp = BufChar
        Else
            strTmp = strTmp & BufChar
            If intIdx = 1 Then Exit For
        End If
    Next
    
    strTmp = Replace(strTmp, "<검사>", "")
    strTmp = Replace(strTmp, "</검사>", "|")
    
    strTmp = Replace(strTmp, "<업체>", "")
    strTmp = Replace(strTmp, "</업체>", ",")
    
    strTmp = Replace(strTmp, "<요양기관번호>", "")
    strTmp = Replace(strTmp, "</요양기관번호>", ",")
    
    strTmp = Replace(strTmp, "<차트번호>", "")
    strTmp = Replace(strTmp, "</차트번호>", ",")
    
    strTmp = Replace(strTmp, "<수진자명>", "")
    strTmp = Replace(strTmp, "</수진자명>", ",")
    
    strTmp = Replace(strTmp, "<주민등록번호>", "")
    strTmp = Replace(strTmp, "</주민등록번호>", ",")
    
    strTmp = Replace(strTmp, "<내원번호>", "")
    strTmp = Replace(strTmp, "</내원번호>", ",")
    
    strTmp = Replace(strTmp, "<의뢰일>", "")
    strTmp = Replace(strTmp, "</의뢰일>", ",")
    
    strTmp = Replace(strTmp, "<검사번호>", "")
    strTmp = Replace(strTmp, "</검사번호>", ",")
    
    strTmp = Replace(strTmp, "<검사ID>", "")
    strTmp = Replace(strTmp, "</검사ID>", ",")
    
    strTmp = Replace(strTmp, "<업체검사ID>", "")
    strTmp = Replace(strTmp, "</업체검사ID>", ",")
    
    strTmp = Replace(strTmp, "<검체>", "")
    strTmp = Replace(strTmp, "</검체>", ",")
    
    strTmp = Replace(strTmp, "<결과치>", "")
    strTmp = Replace(strTmp, "</결과치>", ",")
    
    strTmp = Replace(strTmp, "<참조치>", "")
    strTmp = Replace(strTmp, "</참조치>", ",")
    
    strTmp = Replace(strTmp, "<소견>", "")
    strTmp = Replace(strTmp, "</소견>", ",")
    
    strTmp = Replace(strTmp, "<결과일>", "")
    strTmp = Replace(strTmp, "</결과일>", ",")
    
    strTmp = Replace(strTmp, "<업체>", "")
    strTmp = Replace(strTmp, "</업체>", ",")
    
    strTmp = Replace(strTmp, "<입원외래구분>", "")
    strTmp = Replace(strTmp, "</입원외래구분>", ",")
    
    GetWorkList_XML = Split(strTmp, "|")
    
'    Call SetSQLData("오더저장", strTmp, "A")
    blnSameRecord = True
    
    Screen.MousePointer = 0
    
    Exit Function
        
ErrorTrap:
    
    blnSameRecord = False
    Screen.MousePointer = 0
    
    
End Function

Public Sub GetWorkList_UBCARE(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
    Dim RS_L        As ADODB.Recordset
    Dim RS_C        As ADODB.Recordset
    Dim varXML      As Variant
    Dim varTmp      As Variant
    Dim strBarNum   As String
    Dim strJumin    As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    
    blnSame = False
    
    '1. XML 파일을 읽는다.
    varXML = GetWorkList_XML
    
    If blnSameRecord = True Then
        If UBound(varXML) > 1 Then
            For i = 0 To UBound(varXML) - 1
                varTmp = Split(varXML(i), ",")
                
                '2. 해당검사코드의 채널, 검사명을 가져온다.
                SQL = ""
                SQL = SQL & "SELECT DISTINCT SENDCHANNEL, TESTNAME      " & vbCr
                SQL = SQL & "  FROM EQPMASTER                           " & vbCr
                SQL = SQL & " WHERE TESTCODE = '" & Trim(varTmp(8)) & "'" & vbCr
                    
                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                
                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                    XMLInData.ComExamID = Trim(RS_L.Fields("SENDCHANNEL").Value)
                    XMLInData.Company = varTmp(0)
                    XMLInData.HospCode = varTmp(1)
                    XMLInData.ChartNo = varTmp(2)
                    XMLInData.PatName = varTmp(3)
                    XMLInData.PatJumin = varTmp(4)
                    XMLInData.PatNo = varTmp(5)
                    XMLInData.CommDate = varTmp(6)
                    XMLInData.ExamNo = varTmp(7)
                    XMLInData.ExamID = varTmp(8)
                    'XMLInData.ComExamID = varTmp(9)
                    XMLInData.Specimen = varTmp(10)
                    XMLInData.Result = varTmp(11)
                    XMLInData.Reference = varTmp(12)
                    XMLInData.Remark = varTmp(13)
                    XMLInData.RsltDate = varTmp(14)
                    XMLInData.IOFlag = varTmp(15)
                    
                    strJumin = Left(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
                    Call CalAgeSex(strJumin, Format(Date, "yyyy/mm/dd"))
                    strBarNum = Mid(XMLInData.CommDate, 5, 4) & Format(XMLInData.ChartNo, "0000000000")
                    
                    '3. 새로운 환자인지 확인한다.
                    SQL = ""
                    SQL = SQL & "SELECT DISTINCT CHARTNO    " & vbCr
                    SQL = SQL & "  FROM UB_PATRESULT           " & vbCr
                    SQL = SQL & " WHERE CHARTNO  = '" & XMLInData.ChartNo & "'  " & vbCr
                    SQL = SQL & "   AND EXAMCODE = '" & XMLInData.ExamID & "'   " & vbCr
                    SQL = SQL & "   AND HOSPDATE = '" & XMLInData.CommDate & "' " & vbCr
                    SQL = SQL & "   AND BARCODE  = '" & strBarNum & "'          " & vbCr
                    SQL = SQL & "   AND EXAMTYPE = '" & gHOSP.PARTCD & "'       " & vbCr
                    
                    Set RS_C = AdoCn_Local.Execute(SQL, , 1)
                    
                    If Not RS_C.EOF = True And Not RS_C.BOF = True Then
                        '4-1. 기존 환자인 경우 이름,성별,나이만 업데이트 한다.
                        SQL = ""
                        SQL = SQL & "Update UB_PATRESULT Set "
                        SQL = SQL & "       PNAME    = '" & XMLInData.PatName & "'  " & vbCr
                        SQL = SQL & "     , PSEX     = '" & mPatient.SEX & "'       " & vbCr
                        SQL = SQL & "     , Page     = '" & mPatient.AGE & "'       " & vbCr
                        SQL = SQL & " Where EXAMTYPE = '" & gHOSP.PARTCD & "'       " & vbCr
                        SQL = SQL & "   and HOSPDATE = '" & XMLInData.CommDate & "' " & vbCr
                        SQL = SQL & "   and CHARTNO  = '" & XMLInData.ChartNo & "'  " & vbCr
                        SQL = SQL & "   and BARCODE  = '" & strBarNum & "'          " & vbCr
                        SQL = SQL & "   and EXAMCODE = '" & XMLInData.ExamID & "'   " & vbCr
                    Else
                        '4-2. 새로운 환자인 경우 레코드를 만든다
                        SQL = ""
                        SQL = SQL & "INSERT INTO UB_PATRESULT (" & vbCr
                        SQL = SQL & "  SAVESEQ"                         '저장순번(날짜별)
                        SQL = SQL & ", EXAMDATE"                        '검사일자"
                        SQL = SQL & ", HOSPDATE"                        '병원접수일자"
                        SQL = SQL & ", EQUIPNO"                         '장비코드"
                        SQL = SQL & ", BARCODE              " & vbCr    '검체번호
                        SQL = SQL & ", EQUIPCODE"                       '검사채널"
                        SQL = SQL & ", ORDERCODE"                       '병원처방코드"
                        SQL = SQL & ", EXAMCODE"                        '병원검사코드"
                        SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
                        SQL = SQL & ", EXAMNAME             " & vbCr    '검사명
                        SQL = SQL & ", SAMPLETYPE"                      '검체유형"
                        SQL = SQL & ", INOUT"                           '입/외
                        SQL = SQL & ", REFVALUE"                        '참고치
                        SQL = SQL & ", CHARTNO"                         '챠트번호
                        SQL = SQL & ", PID                  " & vbCr    '병록번호(내원번호)"
                        SQL = SQL & ", PNAME"
                        SQL = SQL & ", PSEX"
                        SQL = SQL & ", PAGE"
                        SQL = SQL & ", PJUMIN"
                        SQL = SQL & ", SENDFLAG             " & vbCr    '전송구분(0:미전송,1:전송)"
                        SQL = SQL & ", SENDDATE"
                        SQL = SQL & ", EXAMUID"
                        SQL = SQL & ", EXAMTYPE"
                        SQL = SQL & ", EXAMNO"
                        SQL = SQL & ", HOSPITAL)            " & vbCr
                        SQL = SQL & " VALUES (              " & vbCr
                        SQL = SQL & 0
                        SQL = SQL & ",'" & Format(Now, "yyyymmddhhmmss") & "'"
                        SQL = SQL & ",'" & XMLInData.CommDate & "'"
                        SQL = SQL & ",'" & gHOSP.MACHCD & "'"
                        SQL = SQL & ",'" & strBarNum & "'                           " & vbCr
                        SQL = SQL & ",'" & XMLInData.ComExamID & "'"
                        SQL = SQL & ",''"
                        SQL = SQL & ",'" & XMLInData.ExamID & "'"
                        SQL = SQL & ",''"
                        SQL = SQL & ",'" & Trim(RS_L.Fields("TESTNAME").Value) & "' " & vbCr
                        SQL = SQL & ",'" & XMLInData.Specimen & "'"                             '검체유형
                        SQL = SQL & ",'" & XMLInData.IOFlag & "'"
                        SQL = SQL & ",'" & XMLInData.Reference & "'"
                        SQL = SQL & ",'" & XMLInData.ChartNo & "'"
                        SQL = SQL & ",'" & XMLInData.PatNo & "'                     " & vbCr
                        SQL = SQL & ",'" & XMLInData.PatName & "'"
                        SQL = SQL & ",'" & mPatient.SEX & "'"
                        SQL = SQL & ",'" & mPatient.AGE & "'"
                        SQL = SQL & ",'" & strJumin & "'"
                        SQL = SQL & ",'0'                                           " & vbCr    '전송구분(0:미전송,1:전송)
                        SQL = SQL & ",''"
                        SQL = SQL & ",'" & gHOSP.USERID & "'"
                        SQL = SQL & ",'" & gHOSP.PARTCD & "'"
                        SQL = SQL & ",'" & XMLInData.ExamNo & "'"
                        SQL = SQL & ",'" & XMLInData.HospCode & "')                 " & vbCr
                    End If
                    
                    RS_C.Close
                    If Not DBExec(AdoCn_Local, SQL) Then
                        Call SetSQLData("저장에러", SQL, "A")
                    End If
                    
                End If
                RS_L.Close
            Next
        End If
    End If
    
    '5. 조회기간의 데이터를 불러온다.
    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       SAVESEQ                 " & vbCr
    SQL = SQL & "     , HOSPDATE                " & vbCr
    SQL = SQL & "     , INOUT                   " & vbCr
    SQL = SQL & "     , CHARTNO                 " & vbCr
    SQL = SQL & "     , BARCODE                 " & vbCr
    SQL = SQL & "     , PID                     " & vbCr
    SQL = SQL & "     , PNAME                   " & vbCr
    SQL = SQL & "     , PSEX                    " & vbCr
    SQL = SQL & "     , PAGE                    " & vbCr
    SQL = SQL & "     , PJUMIN                  " & vbCr
    SQL = SQL & "     , COUNT(EXAMCODE) AS CNT  " & vbCr
    SQL = SQL & "  From UB_PATRESULT                                            " & vbCr
    SQL = SQL & " Where HOSPDATE Between '" & pFrom & "' AND '" & pTo & "'      " & vbCr
    SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")                        " & vbCr
    SQL = SQL & "   And (RESULT = '' OR RESULT IS NULL)                         " & vbCr
    SQL = SQL & " Group By SAVESEQ,HOSPDATE,INOUT,CHARTNO,BARCODE,PID,PNAME,PSEX,PAGE,PJUMIN " & vbCr
    SQL = SQL & " Order by SAVESEQ,HOSPDATE,PNAME                               " & vbCr
            
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    Select Case Trim(Trim(RS.Fields("INOUT")) & "")
                        Case "0":   SetText SPD, "외래", intRow, colINOUT
                        Case "1":   SetText SPD, "입원", intRow, colINOUT
                        Case Else:  SetText SPD, Trim(Trim(RS.Fields("INOUT")) & ""), intRow, colINOUT
                    End Select
                    
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("PJUMIN")) & "", intRow, colPJUMIN
                    SetText SPD, Trim(RS.Fields("PSEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("PAGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub


Public Sub GetWorkList_AMIS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

'    SQL = SQL & "     , O.ACPTSEQ"
'    SQL = SQL & "     , O.RSVACPTSTATE"
'    SQL = SQL & "     , O.RESULTSTATE"
'    SQL = SQL & "     , O.DEPTCODE"
'    SQL = SQL & "     , O.ORDERDATE"
'    SQL = SQL & "     , O.SLIPNO"
'    SQL = SQL & "     , O.IOFLAG"
'    SQL = SQL & "     , O.ORDERCODE"
'    SQL = SQL & "     , O.ORDERNAME"
'    SQL = SQL & "     , R.RESULTFLAG"
'    SQL = SQL & "     , R.RESULTNO" & vbCr
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT"
    SQL = SQL & "       O.ACPTDATE              as HOSPDATE " & vbCr
    SQL = SQL & "     , R.SPCMNO                as BARCODE  " & vbCr
    SQL = SQL & "     , P.PATID                 as PID      " & vbCr
    SQL = SQL & "     , P.PATNAME               as PNAME    " & vbCr
    SQL = SQL & "     , P.SEX                   as SEX      " & vbCr
    SQL = SQL & "     , COUNT(R.RESULTITEMCODE) as CNT      " & vbCr
    SQL = SQL & "  FROM REGISTINFOS O, RESULTOFNUM R, PATMST P                  " & vbCr
    SQL = SQL & " WHERE O.ACPTDATE  = R.ACPTDATE                                " & vbCr
    SQL = SQL & "   AND O.PATID     = R.PATID                                   " & vbCr
    SQL = SQL & "   AND O.ACPTSEQ   = R.ACPTSEQ                                 " & vbCr
    SQL = SQL & "   AND O.PATID     = P.PATID                                   " & vbCr
    SQL = SQL & "   AND O.ACPTDATE BETWEEN '" & pFrom & "' and '" & pTo & "'    " & vbCr
    SQL = SQL & "   AND R.RESULTITEMCODE IN (" & gAllTestCd & ")                " & vbCr
    SQL = SQL & "   AND R.ORDERCODE      IN (" & gAllOrdCd & ")                 " & vbCr
    SQL = SQL & "   AND O.CLAS          = 4                                     " & vbCr '임상병리
    SQL = SQL & "   AND R.RESULTFLAG    = 0                                     " & vbCr
    SQL = SQL & " GROUP BY O.ACPTDATE,R.SPCMNO,P.PATID,P.PATNAME,P.SEX          " & vbCr
    SQL = SQL & " ORDER BY R.SPCMNO                                             " & vbCr

    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_BIT(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.RECEIPTDATE    AS HOSPDATE " & vbCr
    SQL = SQL & "     , a.SPECIMENNUM    AS BARCODE  " & vbCr
    SQL = SQL & "     , a.RECEIPTNO      AS CHARTNO  " & vbCr
    SQL = SQL & "     , a.IPDOPD         AS INOUT    " & vbCr
    SQL = SQL & "     , a.PTNO           AS PID      " & vbCr
    SQL = SQL & "     , a.SNAME          AS PNAME    " & vbCr
    SQL = SQL & "     , COUNT(a.LABCODE) AS CNT      " & vbCr
    SQL = SQL & "   FROM SLA_LabMaster a, SLA_LabResult b   " & vbCr
    SQL = SQL & " WHERE a.RECEIPTNO     = b.RECEIPTNO       " & vbCr
    SQL = SQL & "   AND a.ORDERCODE     = b.ORDERCODE       " & vbCr
    SQL = SQL & "   and a.RECEIPTDATE   = b.RECEIPTDATE     " & vbCr
    SQL = SQL & "   AND a.SPECIMENNUM   = b.SPECIMENNUM     " & vbCr
    SQL = SQL & "   AND a.RECEIPTDATE BETWEEN '" & Format(pFrom, "####-##-##") & "' and '" & Format(pTo, "####-##-##") & "'" & vbCr
    SQL = SQL & "   AND b.LABCODE IN (" & gAllTestCd & ")   " & vbCr
    SQL = SQL & "   AND a.JSTATUS < '3'                     " & vbCr
    SQL = SQL & " GROUP BY a.RECEIPTDATE, a.SPECIMENNUM, a.RECEIPTNO, a.IPDOPD, a.PTNO, a.SNAME " & vbCr
    SQL = SQL & " ORDER BY a.RECEIPTDATE,a.SPECIMENNUM "
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_BITUCHART(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Long
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strChartNo  As String
    Dim strTestNmS  As String
    Dim strItems    As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False

    
'    SQL = SQL & "     , LTRIM(RTRIM(R.RESOCMNUM))   AS BARCODE  " & vbCr
'    SQL = SQL & "     , LTRIM(RTRIM(O.OCMCHTNUM))   AS CHARTNO  " & vbCr
'    SQL = SQL & "     , LTRIM(RTRIM(R.RESOCMNUM))   AS PID      " & vbCr
'    SQL = SQL & "     , R.RESLABCOD                 AS ITEM     " & vbCr
'    SQL = SQL & "     , R.ResOdrSeq , R.ResSeq  , R.ResSubSeq   " & vbCr
    
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       SUBSTRING(O.OCMACPDTM,1,8)  AS HOSPDATE " & vbCr
    SQL = SQL & "     , R.RESOCMNUM                 AS BARCODE  " & vbCr
    SQL = SQL & "     , O.OCMCHTNUM                 AS CHARTNO  " & vbCr
    SQL = SQL & "     , R.RESOCMNUM                 AS PID      " & vbCr
    SQL = SQL & "     , P.PBSPATNAM                 AS PNAME    " & vbCr
    SQL = SQL & "     , P.PBSSEXTYP                 AS SEX      " & vbCr
    SQL = SQL & "     ,COUNT(R.RESLABCOD)           AS CNT      " & vbCr
    SQL = SQL & "   FROM DRBITPACK..RESINF AS R, DRBITPACK..OCMINF AS O, DRBITPACK..PBSINF AS P, DRBITPACK..LABMST AS E, DRBITPACK..ODRINF AS W" & vbCr
    SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & pFrom & "000000" & "' AND '" & pTo & "235959" & "'" & vbCr
    SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')       " & vbCr
    SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")         " & vbCr
    SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM                      " & vbCr
    SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM                   " & vbCr
    SQL = SQL & "   AND R.RESOCMNUM = W.ODROCMNUM                   " & vbCr
    'If UCase(gHOSP.MACHNM) <> "ABBOTTEMERALD" Then
    '    SQL = SQL & "   AND R.RESLABCOD = W.ODRCOD                      " & vbCr
    'End If
    SQL = SQL & "   AND R.RESLABCOD = E.LABCOD                      " & vbCr
    SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCr         '--  'I':중간 'F' 완료"
    SQL = SQL & "   AND W.ODRDELFLG = 'N'                           " & vbCr
    SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)  " & vbCr
    SQL = SQL & " GROUP BY O.OCMACPDTM,R.RESOCMNUM,O.OCMCHTNUM,R.RESOCMNUM,P.PBSPATNAM,P.PBSSEXTYP      " & vbCr
    SQL = SQL & " ORDER BY HOSPDATE, BARCODE, CHARTNO, PID"
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        
        SPD.MaxRows = 0
        
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                
                'Call SetSQLData("HospDate", "|" & strHospDate & ":" & Trim(RS("HOSPDATE")) & "|", "A")
                
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    strChartNo = Trim(GetText(SPD, i, colCHARTNO))
                    
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                        strTestNmS = ""
                        iCnt = 0
                    End If
                Next
                
                If blnSame = False Then
                    'strTestNmS = strTestNmS & GetTestNm(Trim(RS.Fields("ITEM")) & "", False) & "/"
                    'iCnt = iCnt + 1
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS 'gPatOrdCd 찾아옴
                    
                                                                
                    If gHOSP.MACHNM = "BM6010" Then
                        frmMain.txtNum = frmMain.txtNum + 1
                        If frmMain.txtNum = "85" Then
                            frmMain.txtRack = frmMain.txtRack + 1
                            frmMain.txtNum = 1
                        End If
                    
                        SetText SPD, Format(frmMain.txtRack, "00") & "-" & Format(frmMain.txtNum, "00"), intRow, colSEQNO
                        
                        strItems = ""
                        strItems = GetEquipExamCode_BM6010(gHOSP.MACHCD, strBarcode, intRow)
                        
                        SetText SPD, strItems, intRow, colKEY1
                        SetText SPD, Len(strItems) / 4, intRow, colOCNT

                        '-- 검사채널로 장비오더 만들기
                        If Trim(strItems) = "" Then
                            'mOrder.NoOrder = True
                            'mOrder.Order = ""
                            'If intRow > 0 Then
                                '-- 진행상태(Order) 표시
                                Call SetText(SPD, "오더없음", intRow, colSTATE)
                            'End If
                        Else
                            'mOrder.NoOrder = False
                            'mOrder.Order = strItems

                            'If intRow > 0 Then
                                '-- 진행상태(Order) 표시
                                Call SetText(SPD, "오더준비", intRow, colSTATE)
                            'End If
                        End If
                        
                    Else
                        SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    End If
                    
                    'SetText SPD, strTestNmS, intRow, colITEMS
                    
                    If gWORKPOS = "P" Then
                        'SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        'SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_BM6010(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String
    Dim strIntBase  As String
    
    GetEquipExamCode_BM6010 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strIntBase = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            '-- TIBC는 안보냄
            If strIntBase <> "99" Then
                strExamCode = strExamCode & strIntBase & Space(3 - Len(strIntBase)) & "M"
            End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_BM6010 = strExamCode
    
End Function

Public Sub GetWorkList_BIGUBCARE(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       i.IntOdrDte         AS HOSPDATE " & vbCr
    SQL = SQL & "     , i.IntLabNum         AS BARCODE  " & vbCr          ' 검사번호"
    SQL = SQL & "     , i.IntChtNum         AS CHARTNO  " & vbCr          ' 차트번호"
    SQL = SQL & "     , i.IntEmgYon         AS INOUT    " & vbCr          ' 응급여부"
    SQL = SQL & "     , i.IntPatNam         AS PNAME    " & vbCr          ' 환자명"
    SQL = SQL & "     , i.IntSexTyp         AS SEX      " & vbCr          ' 성별"
    SQL = SQL & "     , COUNT(i.IntLabCod)  AS CNT      " & vbCr
    SQL = SQL & "  FROM interfacedb..IntRst i, aphdb..rstinf r  " & vbCr
    SQL = SQL & " WHERE r.RstOdrStt not in ('OC')               " & vbCr
    SQL = SQL & "   AND (r.rstrstval = '' or rstrstval is null) " & vbCr
    'If gHOSP.MACHNM <> "HITACHI7080" Then
        SQL = SQL & "   AND i.intodrtyp = '" & gHOSP.PARTCD & "'" & vbCr  ''HEMO'
    'End If
    SQL = SQL & "   AND i.IntOdrDte BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
    SQL = SQL & "   AND i.IntLabCod + cast(IntLabseq as varchar(3)) IN (" & gAllTestCd & ")" & vbCr
    SQL = SQL & "   AND i.intlabnum = r.rstlabnum       " & vbCr
    SQL = SQL & "   AND i.intodrdte = r.rstodrdte       " & vbCr
    SQL = SQL & "   AND i.intlabseq = r.rstlabseq       " & vbCr
    SQL = SQL & "   AND i.intlabcod = r.rstodrcod       " & vbCr
    SQL = SQL & " GROUP BY i.IntOdrDte, i.IntLabNum, i.IntChtNum, i.IntEmgYon, i.IntPatNam, i.IntSexTyp" & vbCr
    SQL = SQL & " ORDER BY i.IntOdrDte, i.IntLabNum     " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("INOUT")) & "", intRow, colINOUT
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub GetWorkList_BIT70(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strChartNo  As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    pTo = pTo & "235959"
    
    'SQL = SQL & "     , L.LABINSNUM as 처방순서" & vbCr
    'SQL = SQL & "     , L.LABSMPNAM as 검체명" & vbCr
    'SQL = SQL & "     , L.LABODRSTP as SEQ " & vbCr
    'SQL = SQL & "     , M.MANRESNUM as JUMIN" & vbCr
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       L.LABODRDTE         as HOSPDATE " & vbCr
    SQL = SQL & "     , L.LABBARCOD         as BARCODE  " & vbCr
    SQL = SQL & "     , L.LABATTEND         as PID      " & vbCr
    SQL = SQL & "     , L.LABCHTNUM         as CHARTNO  " & vbCr
    SQL = SQL & "     , M.MANADMFOR         as IO       " & vbCr
    SQL = SQL & "     , M.MANPATNAM         as PNAME    " & vbCr
    SQL = SQL & "     , L.LABINSNUM         as SEQ      " & vbCr
    SQL = SQL & "     , COUNT(L.LABODRCOD)  as CNT      " & vbCr
    SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D, ME_MAN M " & vbCr
    SQL = SQL & " WHERE L.LABKEYNUM = D.DATKEYNUM       " & vbCr                    '-- 테이블연결키값
    SQL = SQL & "   AND L.LABATTEND = D.DATATTEND       " & vbCr                    '-- 내원번호
    SQL = SQL & "   AND L.LABATTEND = M.MANATTEND       " & vbCr                    '-- 내원번호
    SQL = SQL & "   AND L.LABCHTNUM = D.DATCHTNUM       " & vbCr                    '-- 챠트번호
    SQL = SQL & "   AND L.LABCHTNUM = M.MANCHTNUM       " & vbCr                    '-- 챠트번호
    SQL = SQL & "   AND L.LABODRDTE = D.DATODRDTE       " & vbCr                    '-- 처방일자
    SQL = SQL & "   AND L.LABODRDTE between  '" & pFrom & "' AND '" & pTo & "'" & vbCr
    SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCr
    SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCr    '-- 취소여부
    SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)  " & vbCr
    SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCr                        '-- 처리상태 (2:접수, 3:결과입력)
    SQL = SQL & " GROUP BY L.LABODRDTE, L.LABBARCOD,L.LABATTEND,L.LABCHTNUM,M.MANADMFOR,M.MANPATNAM " & vbCr
    SQL = SQL & " ORDER BY L.LABODRDTE, L.LABCHTNUM, L.LABBARCOD, L.LABINSNUM                       " & vbCr

    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
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
                    strChartNo = GetText(SPD, i, colCHARTNO)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("CHARTNO")) = strChartNo Then
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
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    Select Case Trim(Trim(RS.Fields("IO")) & "")
                        Case "A":   SetText SPD, "외래", intRow, colINOUT
                        Case "F":   SetText SPD, "입원", intRow, colINOUT
                        Case Else:  SetText SPD, "", intRow, colINOUT
                    End Select
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
                    Else
                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
                    End If
                
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
            frmWorkList.chkAll.Value = "0"
        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMSG = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMSG = strErrMSG & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMSG
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
                    
                    If gEMR <> "KOMAIN" Then
                        'SetText frmMain.spdROrder, GetSampleITEM(intRow, frmMain.spdROrder), intRow, colITEMS
                    End If
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
     
    frmMain.spdROrder.RowHeight(-1) = 15
    frmMain.spdROrder.ReDraw = True
    
    Call frmMain.GetPatTRestResult_Search(1)
    
    Screen.MousePointer = 0

End Sub

'-- 검사자 ITEM 가져오기
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
            SQL = SQL & "   AND O.CLAS = 4                                  " & vbCr '임상병리
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
            SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F')     " & vbCr         '--  'I':중간 'F' 완료"
            SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)      " & vbCr
            SQL = SQL & " Order By R.ResLabCod                                  " & vbCr
        
        Case "BITUCHART"
            SQL = ""
            SQL = SQL & " SELECT DISTINCT R.RESLABCOD AS ITEM                   " & vbCr
            SQL = SQL & "   FROM DRBITPACK..RESINF AS R                         " & vbCr
            SQL = SQL & " WHERE LTRIM(RTRIM(R.RESOCMNUM)) = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")             " & vbCr
            SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F')     " & vbCr         '--  'I':중간 'F' 완료"
            SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)      " & vbCr
            SQL = SQL & " Order By R.RESLABCOD                                  " & vbCr
        
        Case "BIT70"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT L.LABODRCOD as ITEM                " & vbCr
            'SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D, ME_MAN M" & vbCr
            SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D                       " & vbCr
            SQL = SQL & " WHERE L.LABCHTNUM  = '" & strChartNo & "'         " & vbCr
            SQL = SQL & "   AND L.LABODRDTE  = '" & strRegDate & "'         " & vbCr
            SQL = SQL & "   AND L.LABKEYNUM  = D.DATKEYNUM                  " & vbCr                    '-- 테이블연결키값
            SQL = SQL & "   AND L.LABATTEND  = D.DATATTEND                  " & vbCr                    '-- 내원번호
            'SQL = SQL & "   AND L.LABATTEND = M.MANATTEND                  " & vbCr                    '-- 내원번호
            SQL = SQL & "   AND L.LABCHTNUM  = D.DATCHTNUM                  " & vbCr                    '-- 챠트번호
            SQL = SQL & "   AND L.LABCHTNUM  = M.MANCHTNUM                  " & vbCr                    '-- 챠트번호
            SQL = SQL & "   AND L.LABODRDTE  = D.DATODRDTE                  " & vbCr                    '-- 처방일자
            SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCr
            SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCr    '-- 취소여부
            SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)  " & vbCr
            SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCr                            '-- 처리상태 (2:접수, 3:결과입력)
            SQL = SQL & " Order By L.LABODRCOD                              " & vbCr
        
        Case "EONM"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT O.H141_SUGACD AS ITEM              " & vbCr
            SQL = SQL & "  FROM TB_H141_LISTAKEBODY O, TB_A110_PATINFO P    " & vbCr
            SQL = SQL & " Where P.A110_ChartNo = O.H141_CHARTNO             " & vbCr
            SQL = SQL & "   AND O.H141_TSAMPLENO  = '" & strBarcode & "'    " & vbCr
            SQL = SQL & "   AND O.H141_NOTYYN = 'N'                         " & vbCr
            SQL = SQL & "   And O.H141_SUGACD in (" & gAllTestCd & ")       " & vbCr
            SQL = SQL & " ORDER BY O.H141_SUGACD                            " & vbCr
        
         Case "EASYS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM                     " & vbCr
            SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c    " & vbCr
            SQL = SQL & " WHERE a.ACC_YMD   = '" & strRegDate & "'          " & vbCr
            SQL = SQL & "   AND a.RECEPT_NO = '" & strBarcode & "'          " & vbCr
            SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ")            " & vbCr
            SQL = SQL & "   AND a.STS_CD    = 'A'                           " & vbCr 'A:접수, R:결과전송
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
'            SQL = SQL & "    AND L.검사종류 = '" & gHOSP.LABCD & "'" & vbCr
            SQL = ""
            SQL = SQL & "SELECT DISTINCT (L.처방코드 + L.서브코드) AS ITEM                  " & vbCr
            SQL = SQL & "  FROM             TB_진료검사 L                                   " & vbCr
            SQL = SQL & "       INNER JOIN  TB_진료지원 J ON  (L.진료지원ID = J.진료지원ID) " & vbCr
            SQL = SQL & "       INNER JOIN  TB_진료일반 A ON  (J.진료일자   = A.진료일자    " & vbCr
            SQL = SQL & "                                AND   J.챠트번호   = A.챠트번호    " & vbCr
            SQL = SQL & "                                AND   J.진료번호   = A.진료번호)   " & vbCr
            SQL = SQL & " Where L.검체번호= '" & strBarcode & "'                            " & vbCr
            SQL = SQL & "   AND L.검사상태 < 5                                              " & vbCr
            SQL = SQL & "   AND L.처방코드 + L.서브코드 IN (" & gAllTestCd & ")             " & vbCr
            SQL = SQL & " ORDER BY L.처방코드, L.서브코드                                   " & vbCr
        
        Case "KCWH"
            
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
            SQL = SQL & "Select DISTINCT (a.처방코드 + a.서브코드)      AS ITEM     " & vbCr
            SQL = SQL & "  From TB_검사항목 a, TB_진료기본 c                        " & vbCr
            SQL = SQL & " Where a.챠트번호 = '" & strChartNo & "'                   " & vbCr
            SQL = SQL & "   And a.처방번호 > 0                                      " & vbCr
            SQL = SQL & "   And c.진료상태 IN ('1','5','6','7','8','9')             " & vbCr
            SQL = SQL & "   And (a.처방코드 + a.서브코드) IN (" & gAllTestCd & ")   " & vbCr
            SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCr
            SQL = SQL & "   And a.진료년    = c.진료년                              " & vbCr
            SQL = SQL & "   And a.진료월    = c.진료월                              " & vbCr
            SQL = SQL & "   And a.진료일    = c.진료일                              " & vbCr
            SQL = SQL & "   And a.챠트번호  = c.챠트번호                            " & vbCr
            SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCr
            SQL = SQL & " Order By ITEM                                             " & vbCr

'            SQL = ""
'            SQL = SQL & "Select DISTINCT (a.처방코드 + a.서브코드)      AS ITEM     " & vbCr
'            SQL = SQL & "  from tb_검사항목 " & vbCr
'            SQL = SQL & " Where 챠트번호 = '" & argPID & "'" & vbCr
'            SQL = SQL & "   And 진료년   = '" & strYear & "'" & vbCr
'            SQL = SQL & "   And 진료월   = '" & strMonth & "'" & vbCr
'            SQL = SQL & "   And 진료일   = '" & strDay & "'" & vbCr
'            SQL = SQL & "   And 처방번호 > 0 " & vbCr
'            SQL = SQL & "   And (검사결과 is null or 검사결과 = '') " & vbCr
'            SQL = SQL & "   And 처방코드+서브코드 in (" & gAllExam & ")"
        
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
            SQL = SQL & "Select DISTINCT ORCD AS ITEM       " & vbCr
            SQL = SQL & "  From LRESULT                     " & vbCr
            SQL = SQL & " Where SPNO =  '" & strBarcode & "'" & vbCr
            SQL = SQL & "   And ORCD IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   And OKFL <> 'Y'                 " & vbCr   '-- 결과확정유무
            SQL = SQL & " Order By ORCD                     " & vbCr
        
        Case "NEOSOFT"
            If strInOut = "입원" Then
                SQL = ""
                SQL = SQL & "SELECT DISTINCT a.CODE as ITEM                         " & vbCr
                SQL = SQL & "  From E_ORDER..ORDER_IN" & Format(Now, "yyyy") & " a  " & vbCr
                SQL = SQL & " Where a.CHAM_INDEX =  '" & strBarcode & "'            " & vbCr
                SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                  " & vbCr
                SQL = SQL & "   AND a.TRANS = '2'                                   " & vbCr
                SQL = SQL & " ORDER BY a.CODE                                       " & vbCr
            ElseIf strInOut = "외래" Then
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
            SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'     " & vbCr '장비코드
            SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & "   AND C.STATUS   <= '3'                       " & vbCr '검사상태
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
        
            Call SetSQLData("ITEM조회", SQL)
    
            '-- Record Count 가져옴
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
            
                
    Call SetSQLData("ITEM조회", SQL)
    
    gPatOrdCd = ""
    
    If SQL <> "" Then
        '-- Record Count 가져옴
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                If strItems = "" Then
                    strItems = GetTestNm(Trim(RS.Fields("ITEM")) & "", False)
                Else
                    strItems = strItems & "/" & GetTestNm(Trim(RS.Fields("ITEM")), False)
                End If
                
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                
                RS.MoveNext
            Loop
        End If
        
        GetSampleITEM = strItems
        
        RS.Close
    Else
        GetSampleITEM = ""
    End If
    
End Function

'-- 검사자 ITEM 가져오기
Function GetSampleITEM_SP(ByVal asRow As Long, ByVal SPD As vaSpread) As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strRegDate      As String
    Dim strChartNo      As String
    Dim lngExamNo       As Long
    Dim strItems        As String
    Dim strSpcYY        As String
    Dim strSpcNo        As String
    Dim sqlRet          As Integer
    
    GetSampleITEM_SP = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
        
    Select Case gEMR
        Case "KOMAIN"
            If gHOSP.BARUSE = "Y" Then
                '바코드 사용
                SQL = "EXEC AP_INF_BAR_ORDER_CODA '" & gHOSP.MACHCD & "', '" & strBarcode & "'"
            Else
                '바코드 미사용
                SQL = "EXEC AP_INF_S_GETCODA '" & gHOSP.MACHCD & "', '" & strBarcode & "'"
            End If
        
            '-- Record Count 가져옴
            AdoCn.CursorLocation = adUseClient
            Set RS = New ADODB.Recordset
            RS.Open AdoCn.Execute(SQL, sqlRet)
            
            Call SetSQLData("ITEM조회", SQL)
            
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    If strItems = "" Then
                        strItems = GetTestNm(Trim(RS.Fields("CODA")) & "/" & Trim(RS.Fields("SUBCODA")), False)
                    Else
                        strItems = strItems & "," & GetTestNm(Trim(RS.Fields("CODA")) & "/" & Trim(RS.Fields("SUBCODA")), False)
                    End If
                    RS.MoveNext
                Loop
            End If

    End Select
    
    GetSampleITEM_SP = strItems
    
    RS.Close
    
End Function

'-- 검사자 ITEM 가져오기
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
        
    Select Case gEMR
        Case "NAVY"
            SQL = ""
            SQL = SQL & "Select DISTINCT EXAMCODE as ITEM " & vbCr
            SQL = SQL & "  From SLXWORKT " & vbCr
            SQL = SQL & " Where ORDDATE =  '" & strRegDate & "'" & vbCr
            SQL = SQL & "   And HOSPID = '" & gHOSP.HOSPCD & "'" & vbCr         ' 부대코드
            SQL = SQL & "   And ROOMCODE = '" & gHOSP.PARTCD & "'" & vbCr         ' 소변
            SQL = SQL & "   And SPCID    = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   And WORKCODE = '" & strChartNo & "'" & vbCr
            SQL = SQL & "   And PATNO    = '" & strPatID & "'" & vbCr
            SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")" & vbCr        ' 검사코드

    End Select
            
    '-- Record Count 가져옴
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

    Screen.MousePointer = 11
    
    GetSampleInfo = -1
    
    Select Case gEMR
        Case "AMIS"
                Call GetSampleInfo_AMIS(asRow, SPD)
        
        Case "BIGUBCARE"
                Call GetSampleInfo_BIGUBCARE(asRow, SPD)
        
        Case "BIT"
                Call GetSampleInfo_BIT(asRow, SPD)
        
        Case "BIT70"
                Call GetSampleInfo_BIT70(asRow, SPD)
        
        Case "BITUCHART"
                Call GetSampleInfo_BITUCHART(asRow, SPD)
                
        Case "EMEDI"
                Call GetSampleInfo_AMIS(asRow, SPD)
        
        Case "EASYS"
                Call GetSampleInfo_EASYS(asRow, SPD)
        
        Case "EONM"
                Call GetSampleInfo_EONM(asRow, SPD)
            
        Case "GINUS"
                Call GetSampleInfo_GINUS(asRow, SPD)
        
        Case "GSEN"
                Call GetSampleInfo_MSINFOTEC(asRow, SPD)
        
        Case "HWASAN"
                Call GetSampleInfo_HWASAN(asRow, SPD)
        
        Case "JAINCOM"
                Call GetSampleInfo_JAINCOM(asRow, SPD)
        
        Case "JWINFO"
                Call GetSampleInfo_JWINFO(asRow, SPD)

        Case "KCHART"
                Call GetSampleInfo_KCHART(asRow, SPD)
        
        Case "KCWH"
                Call GetSampleInfo_KCWH(asRow, SPD)
        
        Case "KOMAIN"
                Call GetSampleInfo_KOMAIN(asRow, SPD)
        
        Case "KYU"                  '건양대학교병원
                Call GetSampleInfo_KYU(asRow, SPD)

        Case "MCC"
                Call GetSampleInfo_MCC(asRow, SPD)
        
        Case "MEDICHART"
                Call GetSampleInfo_MEDICHART(asRow, SPD)
        
        Case "MEDIIT"
                Call GetSampleInfo_MEDIIT(asRow, SPD)
        
        Case "MEDITOLISS"                   '아름누리
                Call GetSampleInfo_MEDITOLISS(asRow, SPD)
        
        Case "MOD"
                Call GetSampleInfo_MOD(asRow, SPD)
        
        Case "MSINFOTEC"
                Call GetSampleInfo_MSINFOTEC(asRow, SPD)

        Case "NEOSOFT"
                Call GetSampleInfo_NEOSOFT(asRow, SPD)
        
        Case "ONITGUM"                      '온아티 검진
                Call GetSampleInfo_ONITGUM(asRow, SPD)

        Case "ONITEMR"                      '온아티 EMR
                Call GetSampleInfo_ONITEMR(asRow, SPD)
        
        Case "PLIS"                      '온아티 EMR
                Call GetSampleInfo_PLIS(asRow, SPD)
        
        Case "TWIN"
                Call GetSampleInfo_TWIN(asRow, SPD)

        Case "SY"
                Call GetSampleInfo_SY(asRow, SPD)

        Case "UBCARE"
                Call GetSampleInfo_UBCARE(asRow, SPD)
    
    End Select
            
    
    GetSampleInfo = 1
    
    Screen.MousePointer = 0
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_EONM(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo_EONM = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       O.H141_ODRDAT        AS HOSPDATE " & vbCr
    SQL = SQL & "      ,O.H141_TSAMPLENO     AS BARCODE  " & vbCr
    SQL = SQL & "      ,P.A110_CHARTNO       AS CHARTNO  " & vbCr
    SQL = SQL & "      ,P.A110_PATNM         AS PNAME    " & vbCr
    SQL = SQL & "      ,P.A110_JUMIN1        AS AGE      " & vbCr
    SQL = SQL & "      ,P.A110_SEX           AS SEX      " & vbCr
    SQL = SQL & "      ,O.H141_SUGACD        AS ITEM     " & vbCr
    SQL = SQL & "      ,O.H141_SEQNO         AS SUBITEM  " & vbCr
    SQL = SQL & "  FROM TB_H141_LISTAKEBODY O, TB_A110_PATINFO P  " & vbCr
    SQL = SQL & " Where P.A110_ChartNo      = O.H141_CHARTNO      " & vbCr
    SQL = SQL & "   AND O.H141_TSAMPLENO    = '" & strBarcode & "'" & vbCr
    SQL = SQL & "   AND O.H141_NOTYYN       = 'N'                 " & vbCr
    SQL = SQL & "   And O.H141_SUGACD in (" & gAllTestCd & ")     " & vbCr
    SQL = SQL & " Order By O.H141_SUGACD                          " & vbCr
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    '처방이 없을 수도 있으므로..
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("CHARTNO")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SUBITEM")) & ""
                        
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
    
    GetSampleInfo_EONM = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_EONM = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_GINUS(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo_GINUS = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    
    'SQL = SQL & "       rslt.smp_no, rslt.prcp_seq, rslt.exam_seq, rslt.rept_seq, rslt.cd, rslt.pt_no, rslt.exam_stus, rslt.mach_rslt, rslt.exam_rslt ," & vbCr
    
    
    SQL = ""
    SQL = SQL & "SELECT /*+ INDEX (coif scccoifm_ix1) INDEX (prex scrprexh_ix3) INDEX (ptbs pmcptbsm_ux1) INDEX (rslt scrrslth_ux1) INDEX (xpsl mosxpslh_ix2) */" & vbCr
    SQL = SQL & "       prex.acp_dt                                             AS HOSPDATE     " & vbCr
    SQL = SQL & "     , prex.smp_no                                             AS BARCODE      " & vbCr
    SQL = SQL & "     , coif.exam_mach_cd                                                       " & vbCr
    SQL = SQL & "     , rslt.exam_stus                                                          " & vbCr
    SQL = SQL & "     , prex.pt_no                                              AS PID          " & vbCr
    SQL = SQL & "     , ptbs.pt_nm                                              AS PNAME        " & vbCr
    SQL = SQL & "     , ptbs.ssn_1                                                              " & vbCr
    SQL = SQL & "     , ptbs.ssn_2                                                              " & vbCr
    SQL = SQL & "     , DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd)    AS INOUT        " & vbCr
    SQL = SQL & "     , xpsl.adms_ymd                                                           " & vbCr
    SQL = SQL & "     , xpsl.mn_sub_typ_cd                                                      " & vbCr
    SQL = SQL & "     , xpsl.med_dpt_cd                                                         " & vbCr
    SQL = SQL & "     , xpsl.med_ymd                                                            " & vbCr
    SQL = SQL & "     , rslt.prcp_seq, rslt.exam_seq, rslt.rept_seq                             " & vbCr
    SQL = SQL & "     , rslt.cd                                                 AS ITEM         " & vbCr
    SQL = SQL & "     , Max(Trim(coif.lmt_trm_day))                                             " & vbCr
    SQL = SQL & "  FROM scrprexh prex                                                           " & vbCr
    SQL = SQL & "     , pmcptbsm ptbs                                                           " & vbCr
    SQL = SQL & "     , scccoifm coif                                                           " & vbCr
    SQL = SQL & "     , mosxpslh xpsl                                                           " & vbCr
    SQL = SQL & "     , scrrslth rslt                                                           " & vbCr
    SQL = SQL & " WHERE prex.smp_no        = '" & strBarcode & "'                               " & vbCr
    SQL = SQL & "   AND prex.hos_org_no    = '" & gHOSP.HOSPCD & "'                             " & vbCr
    SQL = SQL & "   AND coif.exam_mach_cd  = '" & gHOSP.MACHCD & "'                             " & vbCr
    SQL = SQL & "   AND rslt.hos_org_no    = prex.hos_org_no                                    " & vbCr
    SQL = SQL & "   AND rslt.smp_no        = prex.smp_no                                        " & vbCr
    SQL = SQL & "   AND rslt.prcp_seq      = prex.prcp_seq                                      " & vbCr
    SQL = SQL & "   AND rslt.exam_seq      = prex.exam_seq                                      " & vbCr
    SQL = SQL & "   AND ptbs.hos_org_no    = prex.hos_org_no                                    " & vbCr
    SQL = SQL & "   AND ptbs.pt_no         = prex.pt_no                                         " & vbCr
    SQL = SQL & "   AND coif.hos_org_no    = prex.hos_org_no                                    " & vbCr
'    SQL = SQL & "   AND coif.exam_cd       = prex.cd                                            " & vbCr
    SQL = SQL & "   AND xpsl.smp_no        = prex.smp_no                                        " & vbCr
    SQL = SQL & "   AND xpsl.hos_org_no    = prex.hos_org_no                                    " & vbCr
    SQL = SQL & "   AND coif.use_typ       = 'Y'                                                " & vbCr
    SQL = SQL & "   AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt               " & vbCr
    SQL = SQL & "   AND xpsl.prcp_typ_cd  IN ('O','C')                                          " & vbCr
    SQL = SQL & "   AND rslt.exam_stus    IN ('0')                                              " & vbCr
    SQL = SQL & "   AND rslt.cd           IN (" & gAllTestCd & ")                               " & vbCr
    SQL = SQL & "   GROUP BY prex.acp_dt, prex.smp_no, coif.exam_mach_cd ,rslt.exam_stus,       "
    SQL = SQL & "            prex.pt_no, ptbs.pt_nm, ptbs.ssn_1, ptbs.ssn_2,                    "
    SQL = SQL & "            DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd),              "
    SQL = SQL & "            xpsl.adms_ymd,xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd    " & vbCr
    SQL = SQL & "            ,rslt.cd                                                           " & vbCr
    SQL = SQL & "   ORDER BY prex.acp_dt, prex.smp_no                                           " & vbCr
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    '처방이 없을 수도 있으므로..
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                    
                Select Case Trim(RS.Fields("INOUT"))
                    Case "O": SetText SPD, "외래", asRow, colINOUT
                    Case "E": SetText SPD, "응급", asRow, colINOUT
                    Case "I": SetText SPD, "입원", asRow, colINOUT
                End Select
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("prcp_seq")) & "|" & Trim(RS.Fields("exam_seq")) & "|" & Trim(RS.Fields("rept_seq")) & "|"
                        
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
    
    GetSampleInfo_GINUS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_GINUS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_EASYS(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo_EASYS = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.ACC_YMD   AS HOSPDATE " & vbCr
    SQL = SQL & "     , a.RECEPT_NO AS BARCODE  " & vbCr
    SQL = SQL & "     , a.PTNT_NO   AS PID      " & vbCr
    SQL = SQL & "     , c.PTNT_NM   AS PNAME    " & vbCr
    SQL = SQL & "     , c.BIRTH_YMD AS AGE      " & vbCr
    SQL = SQL & "     , c.SEX       AS SEX      " & vbCr
    SQL = SQL & "     , a.ORD_CD    AS ITEM     " & vbCr
    SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c " & vbCr
    SQL = SQL & " WHERE a.RECEPT_NO = '" & strBarcode & "'  " & vbCr
    SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ")    " & vbCr
    SQL = SQL & "   AND a.STS_CD    = 'A'                   " & vbCr    'A:접수, R:결과전송
    SQL = SQL & "   AND a.SUTAK_CD  = ''                    " & vbCr
    SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO           " & vbCr
    SQL = SQL & "   AND a.PTNT_NO    = c.PTNT_NO            " & vbCr
    SQL = SQL & " ORDER BY a.ORD_CD                         " & vbCr
    
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    '처방이 없을 수도 있으므로..
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_EASYS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_EASYS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_BIT(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo_BIT = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    SQL = ""
    SQL = SQL & " SELECT DISTINCT "
    SQL = SQL & "        SUBSTRING(O.OCMACPDTM,1,8) AS HOSPDATE " & vbCr
    SQL = SQL & "        ,R.RESOCMNUM               AS BARCODE  " & vbCr
    SQL = SQL & "        ,O.OCMCHTNUM               AS CHARTNO  " & vbCr
    SQL = SQL & "        ,R.RESOCMNUM               AS PID      " & vbCr
    SQL = SQL & "        ,P.PBSPATNAM               AS PNAME    " & vbCr
    SQL = SQL & "        ,P.PBSSEXTYP               AS SEX      " & vbCr
    SQL = SQL & "        ,R.ResLabCod               AS ITEM     " & vbCr
    SQL = SQL & "        ,R.ResOdrSeq, R.ResSeq, R.ResSubSeq    " & vbCr
    SQL = SQL & "   FROM RESINF AS R, OCMINF AS O, PBSINF AS P, LABMST AS E     " & vbCr
    SQL = SQL & " WHERE R.RESOCMNUM = O.OCMNUM                                  " & vbCr
    SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM                               " & vbCr
    SQL = SQL & "   AND R.RESLABCOD = E.LABCOD                                  " & vbCr
    SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')                   " & vbCr
    SQL = SQL & "   AND RTRIM(LTRIM(rtrim(R.RESOCMNUM))) = '" & strBarcode & "' " & vbCr
    SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")                     " & vbCr
    SQL = SQL & " Order By R.ResLabCod                                          " & vbCr
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    '처방이 없을 수도 있으므로..
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("ResOdrSeq")) & "|" & Trim(RS.Fields("ResSeq")) & "|" & Trim(RS.Fields("ResSubSeq"))
                        
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
    
    GetSampleInfo_BIT = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_BIT = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_BIGUBCARE(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo_BIGUBCARE = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       i.IntOdrDte                                 AS HOSPDATE " & vbCr
    SQL = SQL & "     , i.IntLabNum                                 AS BARCODE  " & vbCr           ' 검사번호"
    SQL = SQL & "     , i.IntChtNum                                 AS CHARTNO  " & vbCr           ' 차트번호"
    SQL = SQL & "     , i.IntPatNam                                 AS PNAME    " & vbCr             ' 환자명"
    SQL = SQL & "     , i.IntSexTyp                                 AS SEX      " & vbCr                ' 성별"
    SQL = SQL & "     , i.IntEmgYon                                 AS INOUT    " & vbCr              ' 응급여부"
    SQL = SQL & "     , i.IntLabCod + cast(IntLabseq as varchar(3)) AS ITEM     " & vbCr
    SQL = SQL & "     , i.IntLabSeq AS SUBITEM " & vbCr
    SQL = SQL & "  from interfacedb..IntRst i, aphdb..rstinf r " & vbCr
    SQL = SQL & " WHERE r.RstOdrStt not in ('OC') " & vbCr
    SQL = SQL & "   AND (r.rstrstval = '' or rstrstval is null)" & vbCr
    'If gHOSP.MACHNM <> "HITACHI7080" Then
        SQL = SQL & "   AND i.intodrtyp = '" & gHOSP.PARTCD & "'" & vbCr  ''HEMO'
    'End If
    SQL = SQL & "   AND i.IntLabNum = '" & strBarcode & "'" & vbCr
    SQL = SQL & "   AND i.IntLabCod + cast(IntLabseq as varchar(3)) IN (" & gAllTestCd & ")" & vbCr
    SQL = SQL & "   AND i.intlabnum = r.rstlabnum" & vbCr
    SQL = SQL & "   AND i.intodrdte = r.rstodrdte" & vbCr
    SQL = SQL & "   AND i.intlabseq = r.rstlabseq" & vbCr
    SQL = SQL & "   AND i.intlabcod = r.rstodrcod" & vbCr

    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    '처방이 없을 수도 있으므로..
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("INOUT")) & "", asRow, colINOUT
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    '.PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SUBITEM")) & ""

                        
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
    
    GetSampleInfo_BIGUBCARE = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_BIGUBCARE = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function


    
'-- 검사자 정보 가져오기
Function GetSampleInfo_BITUCHART(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim sFrDt           As String
    Dim sToDt           As String
    
On Error GoTo DBErr
    
    GetSampleInfo_BITUCHART = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    If strRegDate = "" Then
        sFrDt = Format(frmMain.dtpFrDt.Value, "yyyymmdd")
        sToDt = Format(frmMain.dtpToDt.Value, "yyyymmdd")
    End If
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       SUBSTRING(O.OCMACPDTM,1,8)  AS HOSPDATE " & vbCr
    SQL = SQL & "     , R.RESOCMNUM                 AS BARCODE  " & vbCr
    SQL = SQL & "     , O.OCMCHTNUM                 AS CHARTNO  " & vbCr
    SQL = SQL & "     , R.RESOCMNUM                 AS PID      " & vbCr
    SQL = SQL & "     , P.PBSPATNAM                 AS PNAME    " & vbCr
    SQL = SQL & "     , P.PBSSEXTYP                 AS SEX      " & vbCr
    SQL = SQL & "     , R.RESLABCOD                 AS ITEM     " & vbCr
    SQL = SQL & "     , R.ResOdrSeq , R.ResSeq  , R.ResSubSeq   " & vbCr
    SQL = SQL & "   FROM DRBITPACK..RESINF AS R, DRBITPACK..OCMINF AS O, DRBITPACK..PBSINF AS P, DRBITPACK..LABMST AS E, DRBITPACK..ODRINF AS W" & vbCr
    If strRegDate <> "" Then
        'SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & strRegDate & "000000" & "' AND '" & strRegDate & "235959" & "'" & vbCrLf
        SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & strRegDate & "' AND '" & strRegDate & "235959" & "'" & vbCrLf
    Else
        'SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & sFrDt & "000000" & "' AND '" & sToDt & "235959" & "'" & vbCrLf
        SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & sFrDt & "' AND '" & sToDt & "235959" & "'" & vbCrLf
    End If
    
    SQL = SQL & "   AND LTRIM(O.OCMCHTNUM) = '" & strChartNo & "'   " & vbCr   '& Space(10 - Len(lsPid)) & lsPid
    SQL = SQL & "   AND LTRIM(R.RESOCMNUM) = '" & strBarcode & "'   " & vbCr   '& Space(14 - Len(lsPid)) & lsPid
    SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')       " & vbCr
    SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")         " & vbCr
    SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM                      " & vbCr
    SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM                   " & vbCr
    SQL = SQL & "   AND R.RESOCMNUM = W.ODROCMNUM                   " & vbCr
    
    'If UCase(gHOSP.MACHNM) <> "ABBOTTEMERALD" Then
    '    SQL = SQL & "   AND R.RESLABCOD = W.ODRCOD                      " & vbCr
    'End If
    
    SQL = SQL & "   AND R.RESLABCOD = E.LABCOD                      " & vbCr
    SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCr         '--  'I':중간 'F' 완료"
    SQL = SQL & "   AND W.ODRDELFLG = 'N'                           " & vbCr
    SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)  " & vbCr
    SQL = SQL & " ORDER BY HOSPDATE, CHARTNO, PID"
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    '처방이 없을 수도 있으므로..
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
'                Select Case Trim(Trim(RS.Fields("IO")) & "")
'                    Case "A":   SetText SPD, "외래", asRow, colINOUT
'                    Case "F":   SetText SPD, "입원", asRow, colINOUT
'                    Case Else:  SetText SPD, Trim(Trim(RS.Fields("IO")) & ""), asRow, colINOUT
'                End Select
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("ResOdrSeq")) & "|" & Trim(RS.Fields("ResSeq")) & "|" & Trim(RS.Fields("ResSubSeq"))   '결과저장용 번호's
                        
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
    
    GetSampleInfo_BITUCHART = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_BITUCHART = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function
    
'-- 검사자 정보 가져오기
Function GetSampleInfo_BIT70(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo_BIT70 = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
'            SQL = SQL & ", L.LABINSNUM as 처방순서"
'            SQL = SQL & ", L.LABSMPNAM as 검체명"
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       L.LABSERIAL             " & vbCr
    SQL = SQL & "      ,L.LABODRDTE as HOSPDATE " & vbCr    '접수일자
    SQL = SQL & "      ,L.LABBARCOD as BARCODE  " & vbCr
    SQL = SQL & "      ,L.LABATTEND as PID      " & vbCr    '내원번호
    SQL = SQL & "      ,L.LABCHTNUM as CHARTNO  " & vbCr    '챠트번호
    SQL = SQL & "      ,M.MANADMFOR as IO       " & vbCr
    SQL = SQL & "      ,M.MANRESNUM as JUMIN    " & vbCr
    SQL = SQL & "      ,M.MANPATNAM as PNAME    " & vbCr
    SQL = SQL & "      ,L.LABODRCOD as ITEM     " & vbCr
    SQL = SQL & "      ,L.LABODRSTP as SEQ      " & vbCr    '검사일련번호
    SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D, ME_MAN M             " & vbCr
    SQL = SQL & " WHERE L.LABCHTNUM = '" & strChartNo & "'          " & vbCr
    SQL = SQL & "   AND L.LABODRDTE = '" & strRegDate & "'          " & vbCr
    SQL = SQL & "   AND L.LABKEYNUM = D.DATKEYNUM                   " & vbCr    '-- 테이블연결키값
    SQL = SQL & "   AND L.LABODRDTE = D.DATODRDTE                   " & vbCr    '-- 처방일자
    SQL = SQL & "   AND L.LABATTEND = D.DATATTEND                   " & vbCr    '-- 내원번호
    SQL = SQL & "   AND L.LABATTEND = M.MANATTEND                   " & vbCr    '-- 내원번호
    SQL = SQL & "   AND L.LABCHTNUM = D.DATCHTNUM                   " & vbCr    '-- 챠트번호
    SQL = SQL & "   AND L.LABCHTNUM = M.MANCHTNUM                   " & vbCr    '-- 챠트번호
    SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCr    '-- 취소여부
    SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)  " & vbCr
    SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCr
    SQL = SQL & "   AND L.LABENDDEP < '3' " & vbCrLf                            '-- 처리상태 (2:접수, 3:결과입력)
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    '처방이 없을 수도 있으므로..
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                Select Case Trim(Trim(RS.Fields("IO")) & "")
                    Case "A":   SetText SPD, "외래", asRow, colINOUT
                    Case "F":   SetText SPD, "입원", asRow, colINOUT
                    Case Else:  SetText SPD, Trim(Trim(RS.Fields("IO")) & ""), asRow, colINOUT
                End Select
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SEQ")) & ""
                        
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
    
    GetSampleInfo_BIT70 = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_BIT70 = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_JWINFO(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_JWINFO = -1
    
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
    SQL = SQL & "       a.RECEIPTDATE    AS HOSPDATE " & vbCr
    SQL = SQL & "     , a.SPECIMENNUM    AS BARCODE  " & vbCr
    SQL = SQL & "     , a.RECEIPTNO      AS CHARTNO  " & vbCr
    SQL = SQL & "     , a.IPDOPD         AS INOUT    " & vbCr
    SQL = SQL & "     , a.PTNO           AS PID      " & vbCr
    SQL = SQL & "     , a.SNAME          AS PNAME    " & vbCr
    SQL = SQL & "     , a.ORDERCODE      AS ORDCODE  " & vbCr
    SQL = SQL & "     , b.LABCODE        AS ITEM     " & vbCr
    SQL = SQL & "   FROM SLA_LabMaster a, SLA_LabResult b   " & vbCr
    SQL = SQL & " WHERE a.RECEIPTNO     = b.RECEIPTNO       " & vbCr
    SQL = SQL & "   AND a.ORDERCODE     = b.ORDERCODE       " & vbCr
    SQL = SQL & "   and a.RECEIPTDATE   = b.RECEIPTDATE     " & vbCr
    SQL = SQL & "   AND a.SPECIMENNUM   = b.SPECIMENNUM     " & vbCr
    SQL = SQL & "   AND a.SPECIMENNUM   = '" & strBarcode & "'" & vbCr
    SQL = SQL & "   AND b.LABCODE IN (" & gAllTestCd & ")   " & vbCr
    SQL = SQL & "   AND a.JSTATUS < '3'                     " & vbCr
    SQL = SQL & " ORDER BY b.LABCODE                        " & vbCr
    
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
                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCODE")) & ""
                        
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
    
    GetSampleInfo_JWINFO = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_JWINFO = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_JAINCOM(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_JAINCOM = -1
    
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
    SQL = SQL & "SELECT DiSTINCT "
    SQL = SQL & "       b.SCP42JDATE            as HOSPDATE         " & vbCr
    SQL = SQL & "     , a.SCP41SPMNO2           as BARCODE          " & vbCr
    SQL = SQL & "     , b.SCP42IDNOA            as PID              " & vbCr
    SQL = SQL & "     , a.SCP41NAME             as PNAME            " & vbCr
    SQL = SQL & "     , a.SCP41SEX              as SEX              " & vbCr
    SQL = SQL & "     , a.SCP41BIRTH            as AGE              " & vbCr
    SQL = SQL & "     , b.SCP42SUGACD           as ITEM             " & vbCr
    SQL = SQL & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b    " & vbCr
    SQL = SQL & " WHERE a.SCP41PCODE    = b.SCP42PCODE              " & vbCr
    SQL = SQL & "   AND a.SCP41JDATE    = b.SCP42JDATE              " & vbCr
    SQL = SQL & "   AND a.SCP41SID      = b.SCP42SID                " & vbCr
    SQL = SQL & "   AND a.SCP41SPMNO2   = b.SCP42SPMNO2             " & vbCr
    SQL = SQL & "   AND a.SCP41SPMNO2   =  '" & strBarcode & "'     " & vbCr
    SQL = SQL & "   AND b.SCP42SUGACD  IN (" & gAllTestCd & ")       " & vbCr
    SQL = SQL & "   AND (b.SCP42RESULT IS NULL OR b.SCP42RESULT = '')" & vbCr
    SQL = SQL & " ORDER BY b.SCP42JDATE, a.SPECIMENNUM               " & vbCr
    
    
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_JAINCOM = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_JAINCOM = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_KCHART(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_KCHART = -1
    
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
    
    'SQL = SQL & "    AND L.검사종류 = '" & gHOSP.LABCD & "'" & vbCr
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       J.접수일자                  AS HOSPDATE " & vbCr
    SQL = SQL & "     , L.검체번호                  AS BARCODE  " & vbCr
    SQL = SQL & "     , A.챠트번호                  AS CHARTNO  " & vbCr
    SQL = SQL & "     , J.접수번호                  AS PID      " & vbCr
    SQL = SQL & "     , A.환자이름                  AS PNAME    " & vbCr
    SQL = SQL & "     , A.환자성별                  AS SEX      " & vbCr
    SQL = SQL & "     , A.환자나이                  AS AGE      " & vbCr
    SQL = SQL & "     , L.진료검사ID                AS TESTID   " & vbCr
    SQL = SQL & "     , L.진료지원ID                AS SPRTID   " & vbCr
    SQL = SQL & "     , (L.처방코드+ L.서브코드)    AS ITEM     " & vbCr
    SQL = SQL & "  FROM             TB_진료검사 L                                    " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료지원 J ON (L.진료지원ID = J.진료지원ID)   " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료일반 A ON (J.진료일자   = A.진료일자      " & vbCr
    SQL = SQL & "                                AND  J.챠트번호   = A.챠트번호      " & vbCr
    SQL = SQL & "                                AND  J.진료번호   = A.진료번호)     " & vbCr
    SQL = SQL & " Where L.검체번호 = '" & strBarcode & "'                  " & vbCr
    SQL = SQL & "   AND L.검사상태 < 5                                     " & vbCr
    SQL = SQL & "   AND L.처방코드 + L.서브코드 IN (" & gAllTestCd & ")    " & vbCr
    SQL = SQL & " ORDER BY J.접수일자, J.접수번호                          " & vbCr
    
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 진료검사ID
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("TESTID")) & ""
                        
                        '-- 진료지원ID
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SPRTID")) & ""
                        
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
    
    GetSampleInfo_KCHART = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_KCHART = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

Function FN_LAB_GET_FLAGS(strMajorcd As String, strMinorcd As String)
    
    FN_LAB_GET_FLAGS = ""
    
    SQL = ""
    SQL = SQL & "SELECT DTLCDNM as FLAG             " & vbCr
    SQL = SQL & "  FROM SLFCDSMT                    " & vbCr
    SQL = SQL & " WHERE CMCD  = '" & strMajorcd & "'" & vbCr
    SQL = SQL & "   AND DTLCD = '" & strMinorcd & "'" & vbCr

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        FN_LAB_GET_FLAGS = Trim(RS.Fields("FLAG")) & ""
    End If
    
    RS.Close
    
End Function

Function FN_LAB_GET_FLAGS_DTLCD(strMajorcd As String, strMinorcd As String)
    
    FN_LAB_GET_FLAGS_DTLCD = ""
    
    SQL = ""
    SQL = SQL & "SELECT DTLCD as DTLCD"
    SQL = SQL & "  FROM SLFCDSMT"
    SQL = SQL & " WHERE CMCD    = vc_Majorcd"
    SQL = SQL & "   AND DTLCDNM = vc_Flag ;"

End Function
       
'-- 검사자 정보 가져오기
'''Function GetSampleInfo_KCWH(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
'''    Dim strRegDate      As String
'''    Dim strBarcode      As String
'''    Dim strPatID        As String
'''    Dim strChartNo      As String
'''    Dim intCol          As Integer
'''    Dim intTestCnt      As Integer
'''
'''
'''On Error GoTo DBErr
'''
'''    GetSampleInfo_KCWH = -1
'''
'''    intTestCnt = 0
'''    gPatOrdCd = ""
'''
'''    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
'''    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
'''    strPatID = Trim(GetText(SPD, asRow, colPID))
'''    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
'''
'''    If strBarcode = "" Then
'''        Exit Function
'''    End If
'''
'''    Screen.MousePointer = 11
'''
'''    If mResult.Kind = "QC" Then
'''        SQL = ""
'''        SQL = SQL & "select "
'''        SQL = SQL & "  ''               as HOSPDATE " & vbCr
'''        SQL = SQL & ", inf.bcd_no       as BARCODE  " & vbCr
'''        SQL = SQL & ", pat.pt_no        as PID      " & vbCr
'''        SQL = SQL & ", pat.pt_nm        as PNAME    " & vbCr
'''        SQL = SQL & ", rst.wrk_grp_cd   as INOUT    " & vbCr ' dpcd
'''        SQL = SQL & ", exm.clsp_cd      as CHARTNO  " & vbCr ' spccd
'''        SQL = SQL & ", rst.tst_cd       as ITEM     " & vbCr ' tclscd
'''        SQL = SQL & ", itm.tnmd         as ITEMNM   " & vbCr ' tnmd
'''        SQL = SQL & ", ''               as AGE      " & vbCr
'''        SQL = SQL & ", ''               as SEX      " & vbCr
'''        SQL = SQL & ", ''               as PJUMIN   " & vbCr
'''        SQL = SQL & ", ''               as SPCCD    " & vbCr
'''        SQL = SQL & ", ''               as SPCNM    " & vbCr
'''        SQL = SQL & ", ''               as ORDCD    " & vbCr
'''        SQL = SQL & "  FROM slqcptmt  pat"
'''        SQL = SQL & "     , slqexamt  exm"
'''        SQL = SQL & "     , slqordmt  ord"
'''        SQL = SQL & "     , slqinfmt  inf"
'''        SQL = SQL & "     , slqtstmt  rst"
'''        SQL = SQL & "     , slfitemt  itm" & vbCr
'''        SQL = SQL & " WHERE inf.bcd_no      = '" & strBarcode & "'  " & vbCr
'''        SQL = SQL & "   AND inf.bcd_no      = exm.bcd_no            " & vbCr
'''        SQL = SQL & "   AND inf.bcd_no      = ord.bCd_NO            " & vbCr
'''        SQL = SQL & "   AND inf.bcd_no      = rst.bcd_no            " & vbCr
'''        SQL = SQL & "   AND inf.pt_no       = pat.pt_no             " & vbCr
'''        SQL = SQL & "   AND ord.ord_seq_no  = exm.ord_seq_no        " & vbCr
'''        SQL = SQL & "   AND exm.ord_dt      = ord.ord_dt            " & vbCr
'''        SQL = SQL & "   AND rst.ord_cd      = exm.ord_cd            " & vbCr
'''        SQL = SQL & "   AND inf.tst_stat_cd IN ('L','R')            " & vbCr
'''        SQL = SQL & "   AND exm.ord_cd      = itm.tclscd            " & vbCr
'''
'''    Else
'''        SQL = ""
'''        SQL = SQL & "SELECT "
'''        SQL = SQL & "  to_char(b.ACPTDT, 'yyyy-MM-dd hh24:mi')                  as HOSPDATE " & vbCr    ' 처방일자
'''        SQL = SQL & ", fn_lab_get_prtbcno_from_bcno(a.BCNO)                     as BARCODE  " & vbCr    ' 검체번호[바코드 번호]
'''        SQL = SQL & ", b.WKGRPCD||'-'||b.WKYMD||'-'||lpad(b.WKSEQ, 4, '0')      as CHARTNO  " & vbCr    ' LAB 번호
'''        SQL = SQL & ", b.PT_NO                                                  as PID      " & vbCr    ' 환자번호
'''        SQL = SQL & ", b.PATNM                                                  as PNAME    " & vbCr    ' 환자명
'''        SQL = SQL & ", b.SEX                                                    as SEX      " & vbCr    ' 성별
'''        SQL = SQL & ", b.AGE                                                    as AGE      " & vbCr    ' 나이
'''        SQL = SQL & ", f.CITIZEN1||'-'||substr(f.CITIZEN2, 1, 1) || '******'    as PJUMIN   " & vbCr    ' 주민번호
'''        SQL = SQL & ", b.IOCLS                                                  as INOUT    " & vbCr    ' 입외구분
'''        SQL = SQL & ", a.SPCCD                                                  as SPCCD    " & vbCr    ' 검체코드
'''        SQL = SQL & ", fn_lab_get_spcnmd(a.SPCCD)                               as SPCNM    " & vbCr    ' 검체명
'''        SQL = SQL & ", c.OTCLSCD                                                as ORDCD    " & vbCr    ' 처방코드 ??
'''        SQL = SQL & ", a.TCLSCD                                                 as ITEM     " & vbCr    ' 검사코드
'''        SQL = SQL & ", d.TNMD                                                   as ITEMNM   " & vbCr    ' 검사명
'''        'SQL = SQL & ", b.ORDDRNM "                                                                      ' 의뢰의사
'''        'SQL = SQL & ", b.DPNM    "                                                                      ' 부서명
'''        'SQL = SQL & ", b.WARDNM  "                                                                      ' 병동명
'''        'SQL = SQL & ", b.DPCD    "                                                                      ' 처방처
'''        SQL = SQL & "  FROM SLRTSTMT a"
'''        SQL = SQL & "     , SLCINFMT b"
'''        SQL = SQL & "     , SLCORDMT c"
'''        SQL = SQL & "     , SLFITEMT d"
'''        SQL = SQL & "     , SLFITEMT i"
'''        SQL = SQL & "     , SLFTIDMT e"
'''        SQL = SQL & "     , APPATBAT f" & vbCr
'''        SQL = SQL & " WHERE a.BCNO      = '" & strBarcode & "'  " & vbCr
'''        SQL = SQL & "   AND a.BCNO      = b.BCNO                " & vbCr
'''        SQL = SQL & "   AND b.BCNO      = c.BCNO                " & vbCr
'''        SQL = SQL & "   AND a.OTCLSCD   = c.OTCLSCD             " & vbCr
'''        SQL = SQL & "   AND a.SPCCD     = c.SPCCD               " & vbCr
'''        SQL = SQL & "   AND b.SPCFLAG   = '" & gHOSP.LABCD & "' " & vbCr  'vc_Acpt
'''        SQL = SQL & "   AND a.RSTFLAG   IN ( 'L', 'R')          " & vbCr
'''        SQL = SQL & "   AND a.TCLSCD    = d.TCLSCD              " & vbCr
'''        SQL = SQL & "   AND c.OTCLSCD   = i.TCLSCD              " & vbCr
'''        SQL = SQL & "   AND a.TCLSCD    = e.TCLSCD              " & vbCr
'''        SQL = SQL & "   AND a.SPCCD     = e.SPCCD               " & vbCr
'''        SQL = SQL & "   AND e.USDT      <= a.RLCOLLDT           " & vbCr
'''        SQL = SQL & "   AND e.UEDT      > a.RLCOLLDT            " & vbCr
'''        SQL = SQL & "   AND b.PT_NO     = f.PT_NO               " & vbCr
'''    '''    SQL = SQL & "   AND a.WKGRPCD   = e.WKGRPCD             " & vbCr
'''    '''    SQL = SQL & "   AND e.WKGRPCD   = '" & gHOSP.PARTCD & "'" & vbCr  '-- in_GrpCode : 그룹코드
'''    ''''    SQL = SQL & "   AND a.WKYMD BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCr
'''    '''    SQL = SQL & "   AND a.WKYMD >= to_date(" & pFrom & ", 'YYYYmmdd')       " & vbCr     '-- 처방 검색 시작일자"
'''    '''    SQL = SQL & "   AND a.WKYMD <= to_date(" & pFrom & ", 'YYYYmmdd') + 1.0 " & vbCr     '-- 처방 검색 종료일자"
'''        SQL = SQL & "   AND a.TCLSCD    IN (" & gAllTestCd & ") " & vbCr
'''        SQL = SQL & " ORDER BY ITEM "
'''    End If
'''
'''    Call SetSQLData("바코드조회", SQL)
'''
'''    '-- Record Count 가져옴
'''    AdoCn.CursorLocation = adUseClient
'''    Set RS = AdoCn.Execute(SQL, , 1)
'''
'''    SetText SPD, "0", asRow, colCHECKBOX
'''
'''    If Not RS.EOF = True And Not RS.BOF = True Then
'''        Do Until RS.EOF
'''            With SPD
'''                .ReDraw = False
'''                intTestCnt = intTestCnt + 1
'''
'''                SetText SPD, "1", asRow, colCHECKBOX
'''                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
'''                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
'''                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colCHARTNO
'''                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
'''                SetText SPD, Trim(RS.Fields("INOUT")) & "", asRow, colINOUT
'''                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
'''                SetText SPD, Trim(RS.Fields("PJUMIN")) & "", asRow, colPJUMIN
'''                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
'''                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
'''
'''                '오더갯수
'''                SetText SPD, CStr(intTestCnt), asRow, colOCNT
'''
'''                '오더정보에 저장
'''                With mOrder
'''                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
'''                    .PID = Trim(RS.Fields("PID")) & ""
'''                    .PNAME = Trim(RS.Fields("PNAME")) & ""
'''                    .PSEX = Trim(RS.Fields("SEX")) & ""
'''                    .Count = CStr(intTestCnt)
'''                    .NoOrder = False
'''                End With
'''
'''                '-- 화면에 표시
'''                For intCol = colSTATE + 1 To .MaxCols
'''                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
'''                        .Row = asRow
'''                        .Col = intCol
'''                        .BackColor = vbYellow
'''                        Call SetText(SPD, "◇", asRow, intCol)
'''
'''                        '-- 처방코드
'''                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCD")) & ""
'''
'''                        '-- 검체코드
'''                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SPCCD")) & ""
'''
'''                        Exit For
'''                    End If
'''                Next
'''
'''                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
'''
'''            End With
'''            DoEvents
'''
'''            RS.MoveNext
'''        Loop
'''    End If
'''
'''    RS.Close
'''
'''    If gPatOrdCd <> "" Then
'''        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
'''    End If
'''
'''    GetSampleInfo_KCWH = 1
'''
'''    Screen.MousePointer = 0
'''
'''Exit Function
'''
'''DBErr:
'''    GetSampleInfo_KCWH = -1
'''    intTestCnt = 0
'''    Screen.MousePointer = 0
'''
'''
'''End Function

Function GetSampleInfo_KCWH(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim strOutData      As Variant
    
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    
'On Error GoTo DBErr
    
    GetSampleInfo_KCWH = -1
    
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
    
    If mResult.Kind = "QC" Then
        SQL = ""
        SQL = SQL & "select "
        SQL = SQL & "  ''               as HOSPDATE " & vbCr
        SQL = SQL & ", inf.bcd_no       as BARCODE  " & vbCr
        SQL = SQL & ", pat.pt_no        as PID      " & vbCr
        SQL = SQL & ", pat.pt_nm        as PNAME    " & vbCr
        SQL = SQL & ", rst.wrk_grp_cd   as INOUT    " & vbCr ' dpcd
        SQL = SQL & ", exm.clsp_cd      as CHARTNO  " & vbCr ' spccd
        SQL = SQL & ", rst.tst_cd       as ITEM     " & vbCr ' tclscd
        SQL = SQL & ", itm.tnmd         as ITEMNM   " & vbCr ' tnmd
        SQL = SQL & ", ''               as AGE      " & vbCr
        SQL = SQL & ", ''               as SEX      " & vbCr
        SQL = SQL & ", ''               as PJUMIN   " & vbCr
        SQL = SQL & ", ''               as SPCCD    " & vbCr
        SQL = SQL & ", ''               as SPCNM    " & vbCr
        SQL = SQL & ", ''               as ORDCD    " & vbCr
        SQL = SQL & "  FROM slqcptmt  pat"
        SQL = SQL & "     , slqexamt  exm"
        SQL = SQL & "     , slqordmt  ord"
        SQL = SQL & "     , slqinfmt  inf"
        SQL = SQL & "     , slqtstmt  rst"
        SQL = SQL & "     , slfitemt  itm" & vbCr
        SQL = SQL & " WHERE inf.bcd_no      = '" & strBarcode & "'  " & vbCr
        SQL = SQL & "   AND inf.bcd_no      = exm.bcd_no            " & vbCr
        SQL = SQL & "   AND inf.bcd_no      = ord.bCd_NO            " & vbCr
        SQL = SQL & "   AND inf.bcd_no      = rst.bcd_no            " & vbCr
        SQL = SQL & "   AND inf.pt_no       = pat.pt_no             " & vbCr
        SQL = SQL & "   AND ord.ord_seq_no  = exm.ord_seq_no        " & vbCr
        SQL = SQL & "   AND exm.ord_dt      = ord.ord_dt            " & vbCr
        SQL = SQL & "   AND rst.ord_cd      = exm.ord_cd            " & vbCr
        SQL = SQL & "   AND inf.tst_stat_cd IN ('L','R')            " & vbCr
        SQL = SQL & "   AND exm.ord_cd      = itm.tclscd            " & vbCr
    
    Else
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
        With AdoCmd
            .CommandTimeout = 15
            .CommandText = "pkg_sup_lab_interface.pc_DownLoad_Gen"
            .CommandType = adCmdStoredProc

            Set Prm1 = .CreateParameter("in_sBcno", adVarChar, adParamInput, 100, strBarcode)
            .Parameters.Append Prm1
            'Set Prm2 = .CreateParameter("out_rtnCs", adVarChar, adParamOutput, 5000, strOutData)
            '.Parameters.Append Prm2
        End With
    
        '-- SP 사용
        Set RS = New ADODB.Recordset
        RS.Open AdoCmd.Execute
        
    End If
    
    Call SetSQLData("바코드조회", strBarcode)
    
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("ORDDT")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BCNO")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("LABNO")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PT_NO")) & "", asRow, colPID
                'SetText SPD, Trim(RS.Fields("INOUT")) & "", asRow, colINOUT
                SetText SPD, Trim(RS.Fields("PATNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("JUMINO")) & "", asRow, colPJUMIN
                SetText SPD, mGetP(Trim(RS.Fields("SEXAGE")) & "", 1, "/"), asRow, colPSEX
                SetText SPD, mGetP(Trim(RS.Fields("SEXAGE")) & "", 2, "/"), asRow, colPAGE
'                SetText SPD, Trim(RS.Fields("SEXAGE")) & "", asRow, colPAGE
                    
'                SetText SPD, mGetP(RS.Fields(11).Value, 1, "/"), asRow, colPSEX
'                SetText SPD, mGetP(RS.Fields(11).Value, 2, "/"), asRow, colPAGE
                'SetText SPD, Trim(RS.Fields("SEXAGE")) & "", asRow, colPAGE
                    
                    
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BCNO")) & ""
                    .PID = Trim(RS.Fields("PT_NO")) & ""
                    .PNAME = Trim(RS.Fields("PATNAME")) & ""
                    .PSEX = mGetP(Trim(RS.Fields("SEXAGE")) & "", 1, "/")
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("TCLSCD")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        'gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCD")) & ""
                        
                        '-- 검체코드
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SPCCD")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("TCLSCD")) & "',"
                
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    'Call SetSQLData("gPatOrdCd", gPatOrdCd)
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_KCWH = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_KCWH = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function
'-- 검사자 정보 가져오기
Function GetSampleInfo_HWASAN(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_HWASAN = -1
    
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
        
    ' O.Sex         성별('M':남, 'F':여, 'A':모두, 'E':기타)
    ' O.StatFg      응급여부('0':아님, '1':응급)
    ' O.AllTestNm   모든 검사명
'    SQL = SQL & "     , O.SpcCd                         " & vbCr
'    SQL = SQL & "     , O.SpcNm                         " & vbCr
'    SQL = SQL & "     , O.AllTestNm                     " & vbCr
'    SQL = SQL & "     , O.StatFg                        " & vbCr
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       O.OrdDt     as HOSPDATE         " & vbCr
    SQL = SQL & "     , O.SPCNO     as BARCODE          " & vbCr
    SQL = SQL & "     , O.PtID      as PID              " & vbCr
    SQL = SQL & "     , O.PtNm      as PNAME            " & vbCr
    SQL = SQL & "     , O.Sex       as SEX              " & vbCr
    SQL = SQL & "     , O.Age       as AGE              " & vbCr
    SQL = SQL & "     , T.TESTCD as ITEM                " & vbCr
    SQL = SQL & "  FROM TC201 O, TC301 T                " & vbCr
    SQL = SQL & " WHERE O.SPCNO = T.SPCNO               " & vbCr
    SQL = SQL & "   AND O.SPCNO = '" & strBarcode & "'    " & vbCr
    SQL = SQL & "   And T.TESTCD in (" & gAllTestCd & ")" & vbCr
    SQL = SQL & " Order By T.TESTCD                     " & vbCr
    
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_HWASAN = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_HWASAN = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_KOMAIN(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim sqlRet          As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo_KOMAIN = -1
    
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
    
    If gHOSP.BARUSE = "Y" Then
        '바코드 사용
        SQL = "EXEC AP_INF_BAR_ORDER_CODA '" & gHOSP.MACHCD & "', '" & strBarcode & "'"
    Else
        '바코드 미사용
        SQL = "EXEC AP_INF_S_GETCODA '" & gHOSP.MACHCD & "', '" & strBarcode & "'"
        'Coda , SubCoda, Sys_Code, ROrder, Serial, Hcode, PtName, Orderdate, Lid

    End If
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = New ADODB.Recordset
    RS.Open AdoCn.Execute(SQL, sqlRet)
    
    Call SetSQLData("바코드조회", SQL)
    
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("ORDERDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("LID")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("SERIAL")) & "", asRow, colPID
                'SetText SPD, Trim(RS.Fields("RORDER")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PTNAME")) & "", asRow, colPNAME
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("LID")) & ""
                    .PID = Trim(RS.Fields("HCODE")) & ""
                    .PNAME = Trim(RS.Fields("PTNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("CODA")) & "/" & Trim(RS.Fields("SUBCODA")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("RORDER")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("CODA")) & "/" & Trim(RS.Fields("SUBCODA")) & "',"
                
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_KOMAIN = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_KOMAIN = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_SY(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim RS1             As ADODB.Recordset
    Dim strRegDate      As String
    Dim strOrgBarcode   As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim sqlRet          As Integer
    Dim lngRegNo        As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_SY = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strOrgBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strBarcode = strOrgBarcode
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    strRegDate = "20" & Format(Mid(strBarcode, 1, 6), "##-##-##")
    lngRegNo = Val(Mid(strBarcode, 7))
        
    Screen.MousePointer = 11
    
    SQL = ""
    SQL = SQL & "Exec Interface_GetPatientResult02 "
    SQL = SQL & "  '" & gHOSP.PARTCD & "'"
    SQL = SQL & " ,'" & strRegDate & "'"
    SQL = SQL & " ,'" & lngRegNo & "'"

    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = New ADODB.Recordset
    RS.Open AdoCn.Execute(SQL, sqlRet)
    
    Call SetSQLData("바코드조회", SQL)
    
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("LabRegDate")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("PatientChartNo")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("LabRegNo")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CompanyCode")) & "", asRow, colINOUT
                SetText SPD, Trim(RS.Fields("PatientName")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("PatientBirthDay")) & "", asRow, colPJUMIN
                SetText SPD, Trim(RS.Fields("PatientSex")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("PatientAge")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = strOrgBarcode
                    .PID = Trim(RS.Fields("LabRegNo")) & ""
                    .PNAME = Trim(RS.Fields("PatientName")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                SQL = ""
                SQL = SQL & " SELECT OrderCode, TestCode, TestSubCode " & vbCr
                SQL = SQL & "   FROM LC26_SQLAB..LabRegResult " & vbCr
                SQL = SQL & "  WHERE LABREGDATE = '" & strRegDate & "'" & vbCr
                SQL = SQL & "    AND LABREGNO   = " & lngRegNo & vbCr
                SQL = SQL & "    AND ORDERCODE  IN (" & gAllTestCd & ")"
                
                Set RS1 = AdoCn.Execute(SQL, , 1)
                If Not RS1.EOF = True And Not RS1.BOF = True Then
                    Do Until RS1.EOF
                        '-- 화면에 표시
                        For intCol = colSTATE + 1 To .vasID.MaxCols
                            If Trim(RS1.Fields("TestSubCode")) = gArrEQP(intCol - colSTATE, 2) Then
                                .Row = asRow
                                .Col = intCol
                                .BackColor = vbYellow
                                .BackColor = vbYellow
                                Call SetText(SPD, "◇", asRow, intCol)
                                
                                '-- 결과저장용 SEQ
                                gArrEQP(intCol - colSTATE, 17) = Trim(RS1.Fields("OrderCode")) & "|" & Trim(RS1.Fields("TestCode")) & "|" & Trim(RS1.Fields("TestSubCode"))
                                
                                gPatOrdCd = gPatOrdCd & "'" & Trim(RS1.Fields("TestSubCode")) & "',"
                                Exit For
                            End If
                        Next
                        
                        RS1.MoveNext
                    Loop
                End If
                RS1.Close
                                
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_SY = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_SY = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_MSINFOTEC(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_MSINFOTEC = -1
    
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
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       a.ORDT          AS HOSPDATE " & vbCr
    SQL = SQL & "     , a.SPNO          AS BARCODE  " & vbCr
    SQL = SQL & "     , a.PAID          AS PID      " & vbCr
    SQL = SQL & "     , a.NWNO          AS CHARTNO  " & vbCr
    SQL = SQL & "     , b.PANM          AS PNAME    " & vbCr
    SQL = SQL & "     , b.SEXS          AS SEX      " & vbCr
    SQL = SQL & "     , b.AGES          AS AGE      " & vbCr
    SQL = SQL & "     , a.ORCD          AS ITEM     " & vbCr
    SQL = SQL & "     , a.ORQN          AS SEQ      " & vbCr
    SQL = SQL & "  From LRESULT a, APATINF b        " & vbCr
    SQL = SQL & " Where a.SPNO = '" & strBarcode & "'   " & vbCr
    SQL = SQL & "   And a.PAID = b.PAID                 " & vbCr
    SQL = SQL & "   And a.ORCD IN (" & gAllTestCd & ")  " & vbCr
    SQL = SQL & "   And a.OKFL <> 'Y'                   " & vbCr   '-- 결과확정유무
    SQL = SQL & " Order By a.ORCD                       " & vbCr
    
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
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SEQ")) & ""
                        
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
    
    GetSampleInfo_MSINFOTEC = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MSINFOTEC = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_MCC(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
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
    SQL = SQL & "       READING_YMD AS HOSPDATE         " & vbCr
    SQL = SQL & "     , BCODE_NO    AS BARCODE          " & vbCr
    SQL = SQL & "     , PTNT_NO     AS PID              " & vbCr
    SQL = SQL & "     , PTNT_NM     AS PNAME            " & vbCr
    SQL = SQL & "     , AGE         AS AGE              " & vbCr
    SQL = SQL & "     , SEX         AS SEX              " & vbCr
    SQL = SQL & "     , IO_GB       AS INOUT            " & vbCr
    SQL = SQL & "     , ORD_CD      AS ITEM             " & vbCr
    SQL = SQL & "     , SP_CD       AS SPCCD            " & vbCr
    SQL = SQL & "  FROM LIS_INTERFACE1_V                " & vbCr
    SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "' " & vbCr
    SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ")  " & vbCr
    SQL = SQL & "   AND STS_CD = '0'                    " & vbCr    '0 접수, 1:결과전송
    SQL = SQL & " ORDER BY ORD_CD                       " & vbCr
        
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
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
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    '.PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_MEDICHART(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_MEDICHART = -1
    
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
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       (a.진료년 + a.진료월 + a.진료일)    AS HOSPDATE     " & vbCr
    SQL = SQL & "     , a.챠트번호                          AS CHARTNO      " & vbCr
    SQL = SQL & "     , c.진료상태                          AS STATE        " & vbCr
    SQL = SQL & "     , b.수진자명                          AS PNAME        " & vbCr
    SQL = SQL & "     , b.주민등록번호                      AS PJUMIN       " & vbCr
    SQL = SQL & "     , (a.처방코드 + a.서브코드)           AS ITEM         " & vbCr
    SQL = SQL & "  From TB_검사항목 a, TB_인적사항 b, TB_진료기본 c         " & vbCr
    SQL = SQL & " Where a.챠트번호 = '" & strChartNo & "'                   " & vbCr
    SQL = SQL & "   And a.처방번호 > 0                                      " & vbCr
    SQL = SQL & "   And c.진료상태 IN ('1','5','6','7','8','9')             " & vbCr
    SQL = SQL & "   And (a.처방코드 + a.서브코드) IN (" & gAllTestCd & ")   " & vbCr
    SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCr
    SQL = SQL & "   And a.진료년    = c.진료년                              " & vbCr
    SQL = SQL & "   And a.진료월    = c.진료월                              " & vbCr
    SQL = SQL & "   And a.진료일    = c.진료일                              " & vbCr
    SQL = SQL & "   And a.챠트번호  = c.챠트번호                            " & vbCr
    SQL = SQL & "   And a.챠트번호  = b.챠트번호                            " & vbCr
    SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCr
    SQL = SQL & " Order By a.진료년, a.진료월, a.진료일, b.수진자명         " & vbCr
        
        
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
                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("CHARTNO")) & ""
                    '.PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_MEDICHART = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MEDICHART = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_MEDIIT(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo        As Long
    
    
On Error GoTo DBErr
    
    GetSampleInfo_MEDIIT = -1
    
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
    SQL = SQL & "       P.request_date      AS HOSPDATE " & vbCr
    SQL = SQL & "     , P.exam_no           AS PID      " & vbCr
    SQL = SQL & "     , P.company_code      AS INOUT    " & vbCr
    SQL = SQL & "     , P.chart_no          AS CHARTNO  " & vbCr
    SQL = SQL & "     , p.person_name       AS PNAME    " & vbCr
    SQL = SQL & "     , P.person_sex        AS SEX      " & vbCr
    SQL = SQL & "     , P.person_age        AS AGE      " & vbCr
    SQL = SQL & "     , R.pro_code          AS ITEM     " & vbCr
    SQL = SQL & "  FROM trust P, trures R               " & vbCr
    SQL = SQL & " WHERE P.request_date  = '" & strRegDate & "'" & vbCr
    SQL = SQL & "   AND P.exam_no       = '" & lngRegNo & "'"
    SQL = SQL & "   AND R.pro_code      IN (" & gAllTestCd & ") " & vbCr
    SQL = SQL & "   AND R.exam_code     <> 'X999'               " & vbCr
    SQL = SQL & "   AND P.request_date  = R.request_date        " & vbCr
    SQL = SQL & "   AND P.exam_no       = R.exam_no             " & vbCr
    SQL = SQL & " ORDER BY P.request_date, P.exam_no            " & vbCr
                
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
                SetText SPD, Trim(RS.Fields("PID")), asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = strBarcode
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_MEDIIT = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MEDIIT = -1
    intTestCnt = 0
    Screen.MousePointer = 0
        
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_MEDITOLISS(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim strJumin        As String
    
On Error GoTo DBErr
    
    GetSampleInfo_MEDITOLISS = -1
    
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
    SQL = SQL & "       A.REQUEST_DATE          AS HOSPDATE                     " & vbCr
    SQL = SQL & "     , A.EXAM_NO               AS BARCODE                      " & vbCr
    SQL = SQL & "     , A.CHART_NO              AS CHARTNO                      " & vbCr
    SQL = SQL & "     , A.PERSON_NAME           AS PNAME                        " & vbCr
    SQL = SQL & "     , A.PERSONAL_ID           AS JUMIN                        " & vbCr
    SQL = SQL & "     , B.EXAM_CODE             AS ITEM                         " & vbCr
    SQL = SQL & "  FROM MEDITOLISS..TOTAL A, MEDITOLISS..TOTRES B               " & vbCr
    SQL = SQL + " WHERE A.REQUEST_DATE  = '" & strRegDate & "'                  " & vbCr
    SQL = SQL & "   And A.EXAM_NO       = '" & strBarcode & "'                  " & vbCr
    SQL = SQL & "   And B.EXAM_CODE     IN (" & gAllTestCd & ")                 " & vbCr
    SQL = SQL & "   AND B.EXAM_PART     = '" & gHOSP.PARTCD & "'                " & vbCr    'C:생화학
    SQL = SQL & "   AND B.RESULT_VALUE  = ''                                    " & vbCr
    SQL = SQL & "   AND A.REQUEST_DATE  = B.REQUEST_DATE                        " & vbCr
    SQL = SQL & "   AND A.EXAM_NO       = B.EXAM_NO                             " & vbCr
    SQL = SQL & " ORDER BY A.REQUEST_DATE, A.EXAM_NO                            " & vbCr
        
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("JUMIN")) & "", asRow, colPJUMIN
                strJumin = Trim(RS.Fields("JUMIN")) & ""
                Call CalAgeSex(strJumin, Format(Date, "yyyy/mm/dd"))
                SetText SPD, mPatient.AGE, asRow, colPAGE
                SetText SPD, mPatient.SEX, asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_MEDITOLISS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MEDITOLISS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_MOD(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_MOD = -1
    
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
    SQL = SQL & "Select Distinct "
    SQL = SQL & "     , a.REQDATE           AS HOSPDATE         " & vbCr
    SQL = SQL & "     , c.SPECIMENID        AS BARCODE          " & vbCr
    SQL = SQL & "       a.PID               AS PID              " & vbCr
    SQL = SQL & "     , a.IOFLAG            AS IO               " & vbCr
    SQL = SQL & "     , b.PAT_NM            AS PNAME            " & vbCr
    SQL = SQL & "     , a.RECENO            AS RECENO           " & vbCr
    SQL = SQL & "     , a.SEQNO             AS SEQ              " & vbCr
    SQL = SQL & "     , c.EXAMCODE          AS ITEM             " & vbCr
    SQL = SQL & "  From EXAMREQ a, TI_PAT b, EXAMRES c          " & vbCr
    SQL = SQL & " Where a.PID        = b.PAT_CHART              " & vbCr
    SQL = SQL & "   And a.PID        = c.PID                    " & vbCr
    SQL = SQL & "   And a.SEQNO      = c.SEQNO                  " & vbCr
    SQL = SQL & "   And a.RECENO     = c.RECENO                 " & vbCr
    SQL = SQL & "   And c.SPECIMENID = '" & strBarcode & "'     " & vbCr
    SQL = SQL & "   And c.EXAMCODE in (" & gAllTestCd & ")      " & vbCr
    SQL = SQL & "   And (c.EXAMEND = '' Or c.EXAMEND IS NULL)   " & vbCr
    SQL = SQL & " Order By a.REQDATE,c.SPECIMENID               " & vbCr
    
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                Select Case Trim(RS.Fields("IO"))
                    Case "1": SetText SPD, "외래", asRow, colINOUT
                    Case "2": SetText SPD, "입원", asRow, colINOUT
                End Select
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방번호
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("RECENO")) & ""
                                                
                        '-- 순번
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SEQ")) & ""
                                                
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
    
    GetSampleInfo_MOD = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MOD = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_NEOSOFT(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_NEOSOFT = -1
    
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
    SQL = SQL & "       a.WORK_DATE         as HOSPDATE " & vbCr
    SQL = SQL & "     , a.CHAM_INDEX        as BARCODE  " & vbCr
    SQL = SQL & "     , a.MEDM_ID           as PID      " & vbCr
    SQL = SQL & "     , b.CHAM_NAME         as PNAME    " & vbCr
    SQL = SQL & "     , b.CHAM_SEX          as SEX      " & vbCr
    SQL = SQL & "     , b.CHAM_YY           as AGE      " & vbCr
    SQL = SQL & "     , '입원'              as IO       " & vbCr
    SQL = SQL & "     , a.CODE              as ITEM     " & vbCr
    SQL = SQL & "  From E_ORDER..ORDER_IN" & Format(Now, "yyyy") & " a "
    SQL = SQL & "     , E_BASECODE..HP_CHAM                          b          " & vbCr
    SQL = SQL & " Where a.CHAM_INDEX = '" & strBarcode & "'                     " & vbCr
    SQL = SQL & "   And a.CHAM_INDEX = b.CHAM_INDEX                             " & vbCr
    SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                          " & vbCr
    SQL = SQL & "   AND a.TRANS = '2'                                           " & vbCr
    SQL = SQL & " UNION ALL                                                     " & vbCr
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.WORK_DATE         as HOSPDATE " & vbCr
    SQL = SQL & "     , a.CHAM_INDEX        as BARCODE  " & vbCr
    SQL = SQL & "     , a.MEDM_ID           as PID      " & vbCr
    SQL = SQL & "     , b.CHAM_NAME         as PNAME    " & vbCr
    SQL = SQL & "     , b.CHAM_SEX          as SEX      " & vbCr
    SQL = SQL & "     , b.CHAM_YY           as AGE      " & vbCr
    SQL = SQL & "     , '외래'              as IO       " & vbCr
    SQL = SQL & "     , a.CODE              as ITEM     " & vbCr
    SQL = SQL & "  From E_ORDER..ORDER_OUT" & Format(Now, "yyyy") & " a "
    SQL = SQL & "     , E_BASECODE..HP_CHAM                           b         " & vbCr
    SQL = SQL & " Where a.CHAM_INDEX = '" & strBarcode & "'                     " & vbCr
    SQL = SQL & "   And a.CHAM_INDEX = b.CHAM_INDEX                             " & vbCr
    SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                          " & vbCr
    SQL = SQL & "   AND a.TRANS = '2'                                           " & vbCr
    SQL = SQL & " ORDER BY a.WORK_DATE, IO, a.CHAM_INDEX                        " & vbCr
        
        
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
                SetText SPD, Trim(RS.Fields("IO")) & "", asRow, colINOUT
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_NEOSOFT = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_NEOSOFT = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_ONITGUM(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_ONITGUM = -1
    
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
    SQL = SQL & "       PER_GUMJIN_DATE     AS HOSPDATE                             " & vbCr
    SQL = SQL & "     , PER_NAME            AS PNAME                                " & vbCr
    SQL = SQL & "     , PER_GUM_NUM         AS BARCODE                              " & vbCr
    SQL = SQL & "     , EDPSCODE            AS ITEM                                 " & vbCr
    SQL = SQL & "  FROM ONIT..GUMJIN_INTERFACE                                      " & vbCr
    SQL = SQL & " WHERE PER_GUMJIN_DATE = '" & strRegDate & "'                     " & vbCr
    SQL = SQL & "   AND PER_GUM_NUM     = '" & strBarcode & "'                      " & vbCr
    SQL = SQL & "   AND EDPSCODE        IN (" & gAllTestCd & ")                     " & vbCr
    SQL = SQL & "   AND (RESULT = ''  OR RESULT IS NULL)                            " & vbCr
    SQL = SQL & " ORDER BY PER_GUMJIN_DATE, PER_GUM_NUM                             " & vbCr
        
        
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_ONITGUM = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_ONITGUM = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_ONITEMR(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_ONITEMR = -1
    
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
    SQL = SQL & "       a.ENTERDATE         AS HOSPDATE     " & vbCr
    SQL = SQL & "     , b.WAITSEQNO         AS BARCODE      " & vbCr
    SQL = SQL & "     , a.CHARTNO           AS CHARTNO      " & vbCr
    SQL = SQL & "     , c.SUJINNAME         AS PNAME        " & vbCr
    SQL = SQL & "     , a.SUJINPART         AS INOUT        " & vbCr    '62:검진
    SQL = SQL & "     , b.MAP2SEQNO         AS ITEM         " & vbCr
    SQL = SQL & "  FROM " & gSQLDB.DB & "..WAITPRSNP a      " & vbCr
    SQL = SQL & "      ," & gSQLDB.DB & "..JUN370_RESULTTB b" & vbCr
    SQL = SQL & "      ," & gSQLDB.DB & "..PEWPRSNP c       " & vbCr
    SQL = SQL & "      ," & gSQLDB.DB & "..BAGMAP2PREF d    " & vbCr
    SQL = SQL & " WHERE a.WAITSEQNO = '" & strBarcode & "'  " & vbCr
    SQL = SQL & "   AND a.JUNDAL    = '" & gHOSP.HOSPCD & "'    " & vbCr        '370
    SQL = SQL & "   AND a.WAITSEQNO = b.WAITSEQNO               " & vbCr
    SQL = SQL & "   AND a.CHARTNO   = c.CHARTNO                 " & vbCr
    SQL = SQL & "   AND d.LABNO     IN (" & gHOSP.LABCD & ")    " & vbCr   '4
    SQL = SQL & "   AND b.MAP2SEQNO IN (" & gAllTestCd & ")     " & vbCr
    SQL = SQL & "   AND b.MAP2SEQNO = d.MAP2SEQNO               " & vbCr
    SQL = SQL & "   AND (b.RESULT = '' OR b.RESULT IS NULL)     " & vbCr
    SQL = SQL & " ORDER BY a.ENTERDATE, b.WAITSEQNO             " & vbCr
                
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                If Trim(RS.Fields("INOUT")) & "" = "62" Then
                    SetText SPD, "검진", asRow, colINOUT
                Else
                    SetText SPD, "진료", asRow, colINOUT
                End If
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_ONITEMR = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_ONITEMR = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_PLIS(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim strSpcYY        As String
    Dim strSpcNo        As String

On Error GoTo DBErr
    
    GetSampleInfo_PLIS = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    If Len(strBarcode) = 11 Then
        strSpcYY = Mid(strBarcode, 1, 2)
        strSpcNo = Mid(strBarcode, 3, 9)
    End If
    
    Screen.MousePointer = 11
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       m.workarea      AS WA               " & vbCr
    SQL = SQL & "     , m.accdt         AS HOSPDATE         " & vbCr
    SQL = SQL & "     , m.accseq        AS ACCSEQ           " & vbCr
    SQL = SQL & "     , m.deptcd        AS INOUT            " & vbCr
    SQL = SQL & "     , m.SEX           AS SEX              " & vbCr
    SQL = SQL & "     , m.AgeDay        AS AGE              " & vbCr
    SQL = SQL & "     , m.ptid          AS PID              " & vbCr
    SQL = SQL & "     , p.ptnm          AS PNAME            " & vbCr
    SQL = SQL & "     , r.testcd AS ITEM                    " & vbCr
    SQL = SQL & "  FROM plis..s2lab201 m                    " & vbCr
    SQL = SQL & "     , his001_v p                          " & vbCr
    SQL = SQL & "     , plis..s2lab302 r                    " & vbCr
    SQL = SQL & "     , plis..s2lab001 e                    " & vbCr
    SQL = SQL & " WHERE m.spcyy     = '" & strSpcYY & "'    " & vbCr
    SQL = SQL & "   AND m.spcno     = '" & strSpcNo & "'    " & vbCr
    SQL = SQL & "   AND m.workarea  = '" & gHOSP.LABCD & "' " & vbCr
    SQL = SQL & "   AND m.workarea  = r.workarea            " & vbCr
    SQL = SQL & "   AND m.accdt     = r.accdt               " & vbCr
    SQL = SQL & "   AND m.accseq    = r.accseq              " & vbCr
    SQL = SQL & "   AND r.testcd    = e.testcd              " & vbCr
    SQL = SQL & "   AND r.testcd    IN (" & gAllTestCd & ") " & vbCr
    SQL = SQL & "   AND (r.vfydt IS NULL OR r.vfydt='')     " & vbCr
    SQL = SQL & "   AND m.ptid = p.ptid COLLATE Korean_Wansung_CS_AS " & vbCr
    SQL = SQL & " ORDER BY m.rcvdt, m.rcvtm                 " & vbCr
                
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("INOUT")) & "", asRow, colINOUT
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- WORKAREA
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("WA")) & ""
                                                
                        '-- ACCSEQ
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("ACCSEQ")) & ""
                        
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
    
    GetSampleInfo_PLIS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_PLIS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_TWIN(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_TWIN = -1
    
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
    SQL = SQL & "       B.JOBDATE                               AS HOSPDATE     " & vbCr
    SQL = SQL & "     , C.SPECNO                                AS BARCODE      " & vbCr
    SQL = SQL & "     , C.PTNO                                  AS CHARTNO      " & vbCr
    SQL = SQL & "     , C.JOBNO                                 AS PID          " & vbCr
    SQL = SQL & "     , DECODE(C.GBIO,'I','입원','O','외래')    AS IO           " & vbCr
    SQL = SQL & "     , C.SNAME                                 AS PNAME        " & vbCr
    SQL = SQL & "     , C.SEX                                   AS SEX          " & vbCr
    SQL = SQL & "     , C.AGE                                   AS AGE          " & vbCr
    SQL = SQL & "     , A.MASTERCODE                            AS ORDERCODE    " & vbCr
    SQL = SQL & "     , A.SUBCODE                               AS ITEM         " & vbCr
    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A                             " & vbCr
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_MASTER  B                             " & vbCr
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_SPECMST C                             " & vbCr
    SQL = SQL & " Where A.SPECNO = '" & strBarcode & "'                         " & vbCr
    SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'                     " & vbCr '장비코드
    SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")                    " & vbCr
    SQL = SQL & "   AND C.STATUS  <= '3'                                        " & vbCr '검사상태
    SQL = SQL & "   And C.SPECNO  = A.SPECNO                                    " & vbCr
    SQL = SQL & "   And A.MASTERCODE = B.MASTERCODE                             " & vbCr
    SQL = SQL & " ORDER BY B.JOBDATE, C.SPECNO                                  " & vbCr
        
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
                SetText SPD, Trim(RS.Fields("IO")) & "", asRow, colINOUT
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDERCODE")) & ""
                        
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
    
    GetSampleInfo_TWIN = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_TWIN = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_UBCARE(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_UBCARE = -1
    
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
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       SAVESEQ                 " & vbCr
    SQL = SQL & "     , HOSPDATE                " & vbCr
    SQL = SQL & "     , INOUT                   " & vbCr
    SQL = SQL & "     , CHARTNO                 " & vbCr
    SQL = SQL & "     , BARCODE                 " & vbCr
    SQL = SQL & "     , PID                     " & vbCr
    SQL = SQL & "     , PNAME                   " & vbCr
    SQL = SQL & "     , PSEX                    " & vbCr
    SQL = SQL & "     , PAGE                    " & vbCr
    SQL = SQL & "     , PJUMIN                  " & vbCr
    SQL = SQL & "     , EXAMCODE        AS ITEM " & vbCr
    SQL = SQL & "  From UB_PATRESULT                        " & vbCr
    SQL = SQL & " Where BARCODE = '" & strBarcode & "'      " & vbCr
    SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")    " & vbCr
    SQL = SQL & "   And (RESULT = '' OR RESULT IS NULL)     " & vbCr
    SQL = SQL & " Order by SAVESEQ,HOSPDATE,PNAME           " & vbCr
        
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
                Select Case Trim(Trim(RS.Fields("INOUT")) & "")
                    Case "0":   SetText SPD, "외래", asRow, colINOUT
                    Case "1":   SetText SPD, "입원", asRow, colINOUT
                    Case Else:  SetText SPD, Trim(Trim(RS.Fields("INOUT")) & ""), asRow, colINOUT
                End Select
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("PJUMIN")) & "", asRow, colPJUMIN
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
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
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_UBCARE = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_UBCARE = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function



'-- 검사자 정보 가져오기
Function GetSampleInfo_AMIS(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_AMIS = -1
    
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
    SQL = SQL & "SELECT DISTINCT"
    SQL = SQL & "       O.ACPTDATE              as HOSPDATE " & vbCr
    SQL = SQL & "     , R.SPCMNO                as BARCODE  " & vbCr
    SQL = SQL & "     , P.PATID                 as PID      " & vbCr
    SQL = SQL & "     , P.PATNAME               as PNAME    " & vbCr
    SQL = SQL & "     , P.SEX                   as SEX      " & vbCr
    SQL = SQL & "     , R.ORDERCODE             as ORDCODE  " & vbCr
    SQL = SQL & "     , R.RESULTITEMCODE        as ITEM     " & vbCr
    SQL = SQL & "  FROM REGISTINFOS O, RESULTOFNUM R, PATMST P      " & vbCr
    SQL = SQL & " WHERE O.ACPTDATE  = R.ACPTDATE                    " & vbCr
    SQL = SQL & "   AND O.PATID     = R.PATID                       " & vbCr
    SQL = SQL & "   AND O.ACPTSEQ   = R.ACPTSEQ                     " & vbCr
    SQL = SQL & "   AND O.PATID     = P.PATID                       " & vbCr
    SQL = SQL & "   AND R.SPCMNO = '" & strBarcode & "'             " & vbCr
    SQL = SQL & "   AND R.RESULTITEMCODE IN (" & gAllTestCd & ")    " & vbCr
    SQL = SQL & "   AND R.ORDERCODE      IN (" & gAllOrdCd & ")     " & vbCr
    SQL = SQL & "   AND O.CLAS          = 4                         " & vbCr '임상병리
    SQL = SQL & "   AND R.RESULTFLAG    = 0                         " & vbCr
    SQL = SQL & " ORDER BY R.SPCMNO                                 " & vbCr
        
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
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCODE")) & ""
                        
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
    
    GetSampleInfo_AMIS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_AMIS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_KYU(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
    Dim strDate     As String
    Dim intBcNow    As Integer
    Dim intBcFive   As Integer
    Dim intBcAdd    As Integer
    Dim strADT      As String
    Dim strSlip1    As String
    Dim strSlip2    As String
    
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim Prm3 As New ADODB.Parameter
    
On Error GoTo DBErr
    
    GetSampleInfo_KYU = -1
    
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
    
    strDate = Format(Now, "yyyy-mm-dd")
    intBcNow = DateDiff("d", "1999-01-01", strDate)
    intBcFive = Mid(strBarcode, 1, 5) '06351
    intBcAdd = intBcFive - intBcNow
    strADT = Format(Now + intBcAdd, "yyyy-mm-dd")
    strSlip1 = Mid(strBarcode, 6, 2)  '바코드번호가 10으로 시작하면 EXAM_TLA_INTERFACE_S,나머지 EXAM_INTERFACE_S
    strSlip2 = Mid(strBarcode, 8, 5)  '00001
   
    '-- SP 사용
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
        Set Prm1 = AdoCmd.CreateParameter("I_JEOBSUDT", adDate, adParamInput, 10, strADT)
        AdoCmd.Parameters.Append Prm1
        Set Prm2 = AdoCmd.CreateParameter("I_BARCODE", adDouble, adParamInput, 12, strBarcode)
        AdoCmd.Parameters.Append Prm2
    Else
        Set Prm1 = AdoCmd.CreateParameter("I_JEOBSUDT", adDate, adParamInput, 10, strADT)
        AdoCmd.Parameters.Append Prm1
        Set Prm2 = AdoCmd.CreateParameter("I_SLIPNO1", adInteger, adParamInput, 2, strSlip1)
        AdoCmd.Parameters.Append Prm2
        Set Prm3 = AdoCmd.CreateParameter("I_SLIPNO2", adInteger, adParamInput, 5, strSlip2)
        AdoCmd.Parameters.Append Prm3
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open AdoCmd.Execute
    
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, strADT, asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("PTNO")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("SLIPNO1")) & "", asRow, colRACKNO
                SetText SPD, Trim(RS.Fields("SLIPNO2")) & "", asRow, colPOSNO
                SetText SPD, Trim(RS.Fields("SNAME")) & "", asRow, colPNAME

                'frmMain.txtRcv.Text = frmMain.txtRcv.Text & Trim(RS.Fields("itemcd")) & vbCr
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = strBarcode
                    .PID = Trim(RS.Fields("PTNO")) & ""
                    .PNAME = Trim(RS.Fields("SNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEMCD")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        Exit For
                    End If
                Next
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEMCD")) & "',"
                
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_KYU = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_KYU = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    
End Function


Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    Dim strTestCd   As String
    
    With frmMain
        Select Case gEMR
            Case "UBCARE"
                sExamDate = Format(.dtpToday, "yyyymmdd")
                If Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) = "" Then
                    Exit Function
                End If
                
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET " & vbCr
                SQL = SQL & "   SAVESEQ       = " & Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) & vbCr
                SQL = SQL & "  ,EXAMDATE      = '" & sExamDate & "' " & vbCr
                SQL = SQL & "  ,DISKNO        = '" & Trim(GetText(.spdOrder, asRow1, colRACKNO)) & "'" & vbCr
                SQL = SQL & "  ,POSNO         = '" & Trim(GetText(.spdOrder, asRow1, colPOSNO)) & "'" & vbCr
                SQL = SQL & "  ,SEQNO         = " & Trim(GetText(.spdResult, asRow2, colRSEQNO)) & vbCr
                SQL = SQL & "  ,EQUIPRESULT   = '" & Trim(GetText(.spdResult, asRow2, colRMACHRESULT)) & "'" & vbCr
                SQL = SQL & "  ,RESULT        = '" & Trim(GetText(.spdResult, asRow2, colRLISRESULT)) & "'" & vbCr
                SQL = SQL & " WHERE EQUIPNO   = '" & gHOSP.MACHCD & "' " & vbCr
                SQL = SQL & "   AND HOSPDATE  = '" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "' " & vbCr
                SQL = SQL & "   AND BARCODE   = '" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "' " & vbCr
                SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'" & vbCr
                SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'"
                
                If Not DBExec(AdoCn_Local, SQL) Then
                    Exit Function
                End If
            
            Case Else
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
                    SQL = SQL & ",'" & Mid(Trim(GetText(.spdOrder, asRow1, colHOSPDATE)), 1, 10) & "'"
                    SQL = SQL & ",'" & gHOSP.MACHCD & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'"     '검사채널
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRORDERCD)) & "'"     '병원처방코드
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
                    SQL = SQL & ",''"                                                   'panic
                    SQL = SQL & ",''"                                                   'delta
                    SQL = SQL & ",'0'"                                                  '전송구분(0:미전송,1:전송)
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & gHOSP.USERID & "'"
                    SQL = SQL & ",'" & gHOSP.HOSPNM & "')"
                    
                    Call SetSQLData("로컬저장", SQL)

                    If Not DBExec(AdoCn_Local, SQL) Then
                        Exit Function
                    End If
                    
                    If UCase(gHOSP.MACHNM) = "ABBOTTEMERALD" And Trim(GetText(.spdResult, asRow2, colRCHANNEL)) = "HGB" Then
                        'B1010,L24
                        If Trim(GetText(.spdResult, asRow2, colRTESTCD)) = "B1010" Then
                            strTestCd = "L24"
                        ElseIf Trim(GetText(.spdResult, asRow2, colRTESTCD)) = "L24" Then
                            strTestCd = "B1010"
                        Else
                            Exit Function
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
                        SQL = SQL & ",'" & Mid(Trim(GetText(.spdOrder, asRow1, colHOSPDATE)), 1, 10) & "'"
                        SQL = SQL & ",'" & gHOSP.MACHCD & "'"
                        SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "'"
                        SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'"     '검사채널
                        SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRORDERCD)) & "'"     '병원처방코드
                        SQL = SQL & ",'" & strTestCd & "'"
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
                        SQL = SQL & ",''"                                                   'panic
                        SQL = SQL & ",''"                                                   'delta
                        SQL = SQL & ",'0'"                                                  '전송구분(0:미전송,1:전송)
                        SQL = SQL & ",''"
                        SQL = SQL & ",'" & gHOSP.USERID & "'"
                        SQL = SQL & ",'" & gHOSP.HOSPNM & "')"
                        
                        Call SetSQLData("로컬저장", SQL)
    
                        If Not DBExec(AdoCn_Local, SQL) Then
                            Exit Function
                        End If
                    End If

                End If
        End Select
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
    Dim strOrderCode     As String
    Dim strTestCode      As String
    Dim strTestName      As String
    Dim strSeqNo         As String
    Dim strRstRow        As Integer
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
        strPtId = Trim(RS.Fields("PID")) & ""
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
            SQL = "SELECT [dbo].FUN_H7LIS_PRE_RESULT4('" & strPtId & "', '" & Format(Now, "yyyymmdd") & "', '" & Format(Now, "hhmm") & "', 'C2302N6')"
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
            SQL = "SELECT [dbo].FUN_H7LIS_PRE_RESULT4('" & strPtId & "', '" & Format(Now, "yyyymmdd") & "', '" & Format(Now, "hhmm") & "', 'C3750N3')"
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
            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
            strSeqNo = Trim(RS_L.Fields("SEQNO")) & ""

            '-- 결과Row 추가
            strRstRow = SPDRST.DataRowCnt + 1
            If SPDRST.MaxRows < strRstRow Then
                SPDRST.MaxRows = strRstRow
            End If
    
            '결과값 표시
            For intCol = colSTATE + 1 To SPDORD.MaxCols
                If strTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
                    SetText SPDORD, strResult, ActiveRow, intCol
                    Exit For
                End If
            Next
    
            '-- 결과 List
            SetText SPDRST, strSeqNo, strRstRow, colRSEQNO                '순번
            SetText SPDRST, strOrderCode, strRstRow, colRORDERCD          '처방코드
            SetText SPDRST, strTestCode, strRstRow, colRTESTCD            '검사코드
            SetText SPDRST, strTestName, strRstRow, colRTESTNM            '검사명
            SetText SPDRST, strIntBase, strRstRow, colRCHANNEL           '장비채널
            SetText SPDRST, strCalResult, strRstRow, colRMACHRESULT      '장비결과
            SetText SPDRST, strCalResult, strRstRow, colRLISRESULT       'LIS결과
            SetText SPDRST, "", strRstRow, colRJUDGE                     '판정
            SetText SPDRST, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
            
            '-- 로컬 저장
            If pTV = "" Then
                SetLocalDB ActiveRow, strRstRow, "1", ""
            Else
                SetLocalDB_R ActiveRow, strRstRow, "1", ""
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
                SetText SPDORD, "저장실패", ActiveRow, colSTATE
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

