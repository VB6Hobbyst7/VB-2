Attribute VB_Name = "modQuery"
Option Explicit

Public SQL              As String
Public RS               As ADODB.Recordset
Public blnSameRecord    As Boolean


'-- 검사마스터 조회
Public Sub GetTestList()
    Dim intRow          As Long
    
    frmMain.spdTest.maxrows = 0
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
            .maxrows = AdoRs_Local.RecordCount
            
            '처방코드, SUB코드용 추가 16,17
            'ReDim Preserve gArrEQP(AdoRs_Local.RecordCount, 18)
            ReDim Preserve gArrEQP(AdoRs_Local.RecordCount + 2, 20)
            
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
                gArrEQP(intRow, 16) = ""    '처방코드로 사용(ACK : SPECIMENCD)
                gArrEQP(intRow, 17) = ""    'SUB코드로 사용 (ACK : LABSEQ)
                gArrEQP(intRow, 18) = ""    'SUB코드로 사용 (ACK : PARTGBN)
                
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
            .maxrows = AdoRs_Local.RecordCount
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


''-- 검사오더마스터 조회
'Public Sub DeleteLocalDB()
'    Dim sDate   As String
'
'    '====================로컬 DB지우기 - 30일 보관======================
'    sDate = Format(DateAdd("y", CDate(dtpToday), -gLocalExpDate), "yyyymmdd")
'
'    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
'    Res = SendQuery(gLocal, SQL)
'    '===================================================================
'
'End Sub

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
        For i = 1 To .maxrows
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
        For i = 1 To .maxrows
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
        frmQCMaster.spdHeader.maxrows = RS.RecordCount
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
    frmQCMaster.spdQCID.maxrows = 0
    
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
        frmQCMaster.spdQCID.maxrows = RS.RecordCount
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
    frmQCMaster.spdDetail.maxrows = 0
    
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
        frmQCMaster.spdDetail.maxrows = RS.RecordCount
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
        .maxrows = .maxrows + 1
        
        Call SetText(frmMain.spdQcResult, "Point", .maxrows, 1)
        Call SetText(frmMain.spdQcResult, strDtTM, .maxrows, 2)
        Call SetText(frmMain.spdQcResult, strRun, .maxrows, 3)
        Call SetText(frmMain.spdQcResult, strQCLevel, .maxrows, 4)
        Call SetText(frmMain.spdQcResult, pEqpCD, .maxrows, 5)
        Call SetText(frmMain.spdQcResult, strLotID, .maxrows, 6)
        Call SetText(frmMain.spdQcResult, pAnalyID, .maxrows, 7)
        Call SetText(frmMain.spdQcResult, strMethodID, .maxrows, 8)
        Call SetText(frmMain.spdQcResult, strMachID, .maxrows, 9)
        Call SetText(frmMain.spdQcResult, strReagentID, .maxrows, 10)
        Call SetText(frmMain.spdQcResult, strUnitID, .maxrows, 11)
        Call SetText(frmMain.spdQcResult, strTemperatureID, .maxrows, 12)
        Call SetText(frmMain.spdQcResult, "sa", .maxrows, 13)
        Call SetText(frmMain.spdQcResult, "", .maxrows, 14)
        Call SetText(frmMain.spdQcResult, "", .maxrows, 15)
        Call SetText(frmMain.spdQcResult, pResult, .maxrows, 16)
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

'-- 워크리스트 조회
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)

    Select Case gEMR

        Case "ONITGUM"                      '온아티 검진
                Call GetWorkList_ONITGUM(pFrom, pTo, SPD)

        Case "ONITEMR"                      '온아티 EMR
                Call GetWorkList_ONITEMR(pFrom, pTo, SPD)
        
        
    End Select

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
        
        SPD.maxrows = 0
        
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
                    .maxrows = .maxrows + 1
                    intRow = .maxrows
                        
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
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
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
    Dim strChartNo  As String
    
    Dim intBcNow    As Integer
    Dim strDate     As String
    Dim strYDate    As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.ENTERDATE         AS HOSPDATE     " & vbCr
    'SQL = SQL & "     , b.WAITSEQNO         AS BARCODE      " & vbCr
    SQL = SQL & "     , b.WAITSEQNO         AS PID          " & vbCr
    SQL = SQL & "     , a.CHARTNO           AS CHARTNO      " & vbCr
    SQL = SQL & "     , c.SUJINNAME         AS PNAME        " & vbCr
    SQL = SQL & "     , a.SUJINPART         AS INOUT        " & vbCr    '62:검진
    SQL = SQL & "     , c.PASSNO            AS JUMIN        " & vbCr
    SQL = SQL & "     ,COUNT(b.MAP2SEQNO)   AS CNT          " & vbCr
    SQL = SQL & "  FROM " & gSQLDB.db & "..WAITPRSNP a      " & vbCr
    SQL = SQL & "      ," & gSQLDB.db & "..JUN370_RESULTTB b" & vbCr
    SQL = SQL & "      ," & gSQLDB.db & "..PEWPRSNP c       " & vbCr
    'SQL = SQL & "      ," & gSQLDB.DB & "..BAGMAP2PREF d    " & vbCr
    SQL = SQL & " WHERE a.ENTERDATE BETWEEN  '" & pFrom & "' AND '" & pTo & "' " & vbCr
    SQL = SQL & "    AND a.JUNDAL       = '" & gHOSP.HOSPCD & "'    " & vbCr        '370
    SQL = SQL & "    AND a.WAITSEQNO    = b.WAITSEQNO               " & vbCr
    SQL = SQL & "    AND a.CHARTNO      = c.CHARTNO                 " & vbCr
    'SQL = SQL & "    AND d.LABNO        IN (" & gHOSP.LABCD & ")    " & vbCr   '4
    SQL = SQL & "    AND b.MAP2SEQNO    IN (" & gAllTestCd & ")     " & vbCr
    'SQL = SQL & "    AND b.MAP2SEQNO    = d.MAP2SEQNO               " & vbCr
    SQL = SQL & "    AND (b.RESULT = '' OR b.RESULT IS NULL)        " & vbCr
    SQL = SQL & " GROUP BY a.ENTERDATE, b.WAITSEQNO, a.CHARTNO, c.SUJINNAME, a.SUJINPART,c.PASSNO" & vbCr
    SQL = SQL & " ORDER BY a.ENTERDATE, b.WAITSEQNO                 " & vbCr
    
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        
        SPD.maxrows = 0
        
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                
                strDate = Format(Now, "yyyy-mm-dd")
                strYDate = Format(Now, "yyyy-01-01")
                
                intBcNow = DateDiff("d", strYDate, strDate)
                intBcNow = intBcNow + 1
                'intBcNow = Format(intBcNow, "000")
                
                strBarcode = Mid(strDate, 4, 1) & Format(CStr(intBcNow), "000") & Trim(RS.Fields("CHARTNO")) & ""
                
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strChartNo = GetText(SPD, i, colCHARTNO)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("CHARTNO")) = strChartNo Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    .maxrows = .maxrows + 1
                    intRow = .maxrows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, strBarcode, intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                                        
                    Call CalAgeSex(Trim(RS.Fields("JUMIN")) & "", Now)
                    
                    SetText SPD, mPatient.sex, intRow, colPSEX
                    SetText SPD, mPatient.age, intRow, colPAGE
                    
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
     
                strErrMsg = "위    치 : " & gHOSP.MACHNM & "_GetWorkList" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0
    
End Sub

Public Sub LetEqpMaster(ByVal pEqpCD As String)

    SQL = ""
    SQL = SQL & "UPDATE EQPMASTER SET EQUIPCD = '" & pEqpCD & "'"
                          
    Call DBExec(AdoCn_Local, SQL)

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
    Dim intTestCnt  As Integer
    
    Screen.MousePointer = 11
    iCnt = 0
    intTestCnt = 0
    
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
        frmMain.spdROrder.maxrows = 0
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
                
                intTestCnt = intTestCnt + 1
                
                If strSaveSeq <> Trim(RS.Fields("SAVESEQ")) & "" Or strExamDate <> Trim(RS.Fields("EXAMDATE")) & "" Then
                    .maxrows = .maxrows + 1
                    intRow = .maxrows
                    intTestCnt = 1
                    
                    SetText frmMain.spdROrder, "1", intRow, colCHECKBOX
                    SetText frmMain.spdROrder, Trim(RS.Fields("SAVESEQ")) & "", intRow, colSAVESEQ
                    SetText frmMain.spdROrder, Trim(RS.Fields("EXAMDATE")) & "", intRow, colEXAMDATE
                    SetText frmMain.spdROrder, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText frmMain.spdROrder, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText frmMain.spdROrder, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText frmMain.spdROrder, Trim(RS.Fields("DISKNO")) & "", intRow, colRACKNO
                    SetText frmMain.spdROrder, Trim(RS.Fields("POSNO")) & "", intRow, colPOSNO
                    SetText frmMain.spdROrder, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText frmMain.spdROrder, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText frmMain.spdROrder, Trim(RS.Fields("PSEX")) & "", intRow, colPSEX
                    SetText frmMain.spdROrder, Trim(RS.Fields("PAGE")) & "", intRow, colPAGE
                    SetText frmMain.spdROrder, Trim(RS.Fields("PJUMIN")) & "", intRow, colPJUMIN
                    SetText frmMain.spdROrder, Trim(RS.Fields("INOUT")) & "", intRow, colINOUT
                    SetText frmMain.spdROrder, Trim(RS.Fields("EQUIPNO")) & "", intRow, colKEY1
                    SetText frmMain.spdROrder, CStr(intTestCnt), intRow, colRCNT
                    
                    
                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                    Case "0"
                            SetText frmMain.spdROrder, "장비결과", intRow, colSTATE
                    Case "2"
                            SetText frmMain.spdROrder, "전송완료", intRow, colSTATE
                    End Select
                    
'                    If intRow <> 1 Then
'                        intTestCnt = 0
'                    End If
'                    If gEMR <> "KOMAIN" Then
'                        SetText frmMain.spdROrder, GetSampleITEM(intRow, frmMain.spdROrder), intRow, colITEMS
'                    End If
                End If
                    
                SetText frmMain.spdROrder, CStr(intTestCnt), intRow, colRCNT
                
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
            SQL = SQL & "  FROM " & gSQLDB.db & "..WAITPRSNP a      " & vbCr
            SQL = SQL & "      ," & gSQLDB.db & "..JUN370_RESULTTB b" & vbCr
            SQL = SQL & "      ," & gSQLDB.db & "..PEWPRSNP c       " & vbCr
'            SQL = SQL & "      ," & gSQLDB.DB & "..BAGMAP2PREF d    " & vbCr
            SQL = SQL & " WHERE a.WAITSEQNO = '" & strPatID & "'  " & vbCr
            SQL = SQL & "   AND a.JUNDAL    = '" & gHOSP.HOSPCD & "'" & vbCr        '370
            SQL = SQL & "   AND a.WAITSEQNO = b.WAITSEQNO           " & vbCr
            SQL = SQL & "   AND a.CHARTNO   = c.CHARTNO             " & vbCr
            'SQL = SQL & "   AND d.LABNO     IN (" & gHOSP.LABCD & ")" & vbCr   '4
            SQL = SQL & "   AND b.MAP2SEQNO IN (" & gAllTestCd & ") " & vbCr
            'SQL = SQL & "   AND b.MAP2SEQNO = d.MAP2SEQNO           " & vbCr
            SQL = SQL & "   AND (b.RESULT = '' OR b.RESULT IS NULL) " & vbCr
        

    End Select
            
                
    Call SetSQLData("ITEM조회", SQL)
    
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
                RS.MoveNext
            Loop
        End If
        
        GetSampleITEM = strItems
        
        RS.Close
    Else
        GetSampleITEM = ""
    End If
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer

    Screen.MousePointer = 11
    
    GetSampleInfo = -1
    
    Select Case gEMR
        Case "ONITEMR"                      '온아티 EMR
                Call GetSampleInfo_ONITEMR(asRow, SPD)
        Case "DONGGUK"
                Call GetSampleInfo_DONGGUK(asRow, SPD)

    End Select
            
    
    GetSampleInfo = 1
    
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
    Dim strDate         As String
    Dim strYDate        As String
    Dim intBcNow        As Integer
    
On Error GoTo DBErr
    
    GetSampleInfo_ONITEMR = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strPatID = "" Then
        strPatID = Mid(strBarcode, 2, 3)
        strPatID = Val(strPatID)
                
        strDate = Format(Now, "yyyy-mm-dd")
        strYDate = Format(Now, "yyyy-01-01")
        intBcNow = DateDiff("d", strYDate, strDate)
        
        strRegDate = DateAdd("d", CInt(strPatID) - 1, strYDate)
        strRegDate = Format(strRegDate, "yyyymmdd")
    
    End If
    
    If strChartNo = "" Then
        strChartNo = Mid(strBarcode, 5)

    End If

    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.ENTERDATE         AS HOSPDATE     " & vbCr
'    SQL = SQL & "     , b.WAITSEQNO         AS BARCODE      " & vbCr
    SQL = SQL & "     , b.WAITSEQNO         AS PID          " & vbCr
    SQL = SQL & "     , a.CHARTNO           AS CHARTNO      " & vbCr
    SQL = SQL & "     , c.SUJINNAME         AS PNAME        " & vbCr
    SQL = SQL & "     , a.SUJINPART         AS INOUT        " & vbCr    '62:검진
    SQL = SQL & "     , c.PASSNO            AS JUMIN        " & vbCr
    SQL = SQL & "     , b.MAP2SEQNO         AS ITEM         " & vbCr
    
    SQL = SQL & "  FROM " & gSQLDB.db & "..WAITPRSNP a      " & vbCr
    SQL = SQL & "      ," & gSQLDB.db & "..JUN370_RESULTTB b" & vbCr
    SQL = SQL & "      ," & gSQLDB.db & "..PEWPRSNP c       " & vbCr
'    SQL = SQL & "      ," & gSQLDB.DB & "..BAGMAP2PREF d    " & vbCr
    SQL = SQL & " WHERE a.CHARTNO = '" & strChartNo & "'  " & vbCr
    SQL = SQL & "   AND a.ENTERDATE = '" & strRegDate & "'    " & vbCr
    SQL = SQL & "   AND a.JUNDAL    = '" & gHOSP.HOSPCD & "'    " & vbCr        '370
    
    SQL = SQL & "   AND a.WAITSEQNO = b.WAITSEQNO               " & vbCr
    SQL = SQL & "   AND a.CHARTNO   = c.CHARTNO                 " & vbCr
'    SQL = SQL & "   AND d.LABNO     IN (" & gHOSP.LABCD & ")    " & vbCr   '4
    SQL = SQL & "   AND b.MAP2SEQNO IN (" & gAllTestCd & ")     " & vbCr
'    SQL = SQL & "   AND b.MAP2SEQNO = d.MAP2SEQNO               " & vbCr
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
                
                strDate = Format(Now, "yyyy-mm-dd")
                strYDate = Format(Now, "yyyy-01-01")
                intBcNow = DateDiff("d", strYDate, strDate)
                
                If strBarcode = "" Then
                    strBarcode = Mid(strDate, 4, 1) & CStr(intBcNow) & Trim(RS.Fields("CHARTNO")) & ""
                End If
                
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, strBarcode, asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                
                If Trim(RS.Fields("INOUT")) & "" = "62" Then
                    SetText SPD, "검진", asRow, colINOUT
                Else
                    SetText SPD, "진료", asRow, colINOUT
                End If
                
                Call CalAgeSex(Trim(RS.Fields("JUMIN")) & "", Now)
                
                SetText SPD, mPatient.sex, asRow, colPSEX
                SetText SPD, mPatient.age, asRow, colPAGE
                
                
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
    
    GetSampleInfo_ONITEMR = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_ONITEMR = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

Function GetSampleInfo_DONGGUK(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim strDate         As String
    Dim strYDate        As String
    Dim intBcNow        As Integer
    
    Dim rv              As Integer
    Dim idates1$, idates2$, iexamcode$
    Dim pt_no$(), patname$(), sex$(), age$()
    Dim spc_no$(), gnl_item_cd$(), bl_gth_dte$()
    Dim dept$(), wd_no$(), tst_cd$()
    Dim ispcno$
    Dim intTstCnt       As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_DONGGUK = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    ispcno$ = strBarcode
    rv = sl_d_60_sel_spcno&(strBarcode, pt_no$(), patname$(), sex$(), age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
                
    Call SetSQLData("바코드조회", ispcno$ & ",rv:" & rv, "A")
    
    If rv >= 1 Then
        SetText SPD, "0", asRow, colCHECKBOX
        With SPD
            .ReDraw = False
            intTestCnt = intTestCnt + 1
            
            SetText SPD, "1", asRow, colCHECKBOX
            SetText SPD, bl_gth_dte(0), asRow, colHOSPDATE
            SetText SPD, pt_no(0), asRow, colPID
            SetText SPD, patname(0), asRow, colPNAME
            SetText SPD, sex(0), asRow, colPSEX
            SetText SPD, age(0), asRow, colPAGE
            
            '오더갯수
            SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                             
            '오더정보에 저장
            With mOrder
                .BarNo = strBarcode
                .PID = pt_no(0)
                .PNAME = patname(0)
                .Count = CStr(intTestCnt)
                .NoOrder = False
            End With
            
            '-- 검사갯수만큼
            For intTstCnt = 0 To UBound(tst_cd)
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If tst_cd(intTstCnt) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        Exit For
                    End If
                Next
                gPatOrdCd = gPatOrdCd & "'" & Trim(tst_cd(intTstCnt)) & "',"
            Next
            
        End With
        DoEvents
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_DONGGUK = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_DONGGUK = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
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
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "'"
                    SQL = SQL & ",'" & gHOSP.MACHCD & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRORDERCD)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRSUBCD)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTNM)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRSEQNO)) & "'"
                    If gEMR = "ACK" Then
                        SQL = SQL & ",'" & mResult.SPECIMENCD & "'"                                                   '검체유형
                        SQL = SQL & ",'" & mResult.PARTGBN & "'"
                    Else
                        SQL = SQL & ",''"                                                   '검체유형
                        SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colINOUT)) & "'"
                    End If
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
                    SQL = SQL & ",'" & mResult.Panic & "'"                                                   'panic
                    SQL = SQL & ",'" & mResult.Delta & "'"                                                   'delta
                    SQL = SQL & ",'0'"                                                  '전송구분(0:미전송,1:전송)
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & gHOSP.USERID & "'"
                    SQL = SQL & ",'" & gHOSP.HOSPNM & "')"
                    
'                    Call SetSQLData("로컬저장", SQL, "A")
                    
                    If Not DBExec(AdoCn_Local, SQL) Then
                        Exit Function
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

