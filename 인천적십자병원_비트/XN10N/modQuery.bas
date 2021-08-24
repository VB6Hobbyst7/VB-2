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
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("CUTOFFUSE").Value & "", intRow, colLCUTUSE)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COLIN").Value & "", intRow, colLCOLIN)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COLCOMP").Value & "", intRow, colLCOLCOMP)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COLOUT").Value & "", intRow, colLCOLOUT)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COMOUT").Value & "", intRow, colLCOMOUT)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COHIN").Value & "", intRow, colLCOHIN)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COHCOMP").Value & "", intRow, colLCOHCOMP)
                Call SetText(frmMain.spdTest, AdoRs_Local.Fields("COHOUT").Value & "", intRow, colLCOHOUT)
                
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

'-- 워크리스트 조회 (바코드 번호)
Public Sub GetWorkList_Barcode(ByVal pBarNo As String)
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
    
    Screen.MousePointer = 11
    
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
            SQL = SQL & " WHERE a.request_date + a.chart_no) = '" & pBarNo & "'" & vbCr
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
        Case "MCC"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "  ORD_YMD AS HOSPDATE"
            SQL = SQL & ", BCODE_NO AS BARCODE"
            SQL = SQL & ", PTNT_NO AS PID"
            SQL = SQL & ", RECEPT_NO AS CHARTNO"
            SQL = SQL & ", '' AS JUMIN"
            SQL = SQL & ", PTNT_NM AS PNAME"
            SQL = SQL & ", SEX AS SEX"
            SQL = SQL & ", AGE AS AGE"
            SQL = SQL & ", COUNT(ORD_CD) AS CNT " & vbCr
            SQL = SQL & "  FROM H7LIS_BCODE_ORD " & vbCr
            SQL = SQL & " WHERE BCODE_NO = '" & pBarNo & "'" & vbCr
            SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND RESULT_TYPE = '20'" & vbCr
            SQL = SQL & " GROUP BY ORD_YMD,BCODE_NO,PTNT_NO,RECEPT_NO,PTNT_NM,SEX,AGE " & vbCr
            SQL = SQL & " ORDER BY ORD_YMD,RECEPT_NO,BCODE_NO "
        
            Call SetSQLData("워크조회(바코드)", SQL)
            
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
                            
                        SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                        SetText frmWorkList.spdWork, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                        SetText frmWorkList.spdWork, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
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
        
    End Select
    
    RS.Close
     
    frmWorkList.spdWork.RowHeight(-1) = 12
    frmWorkList.spdWork.ReDraw = True
    
    Screen.MousePointer = 0

End Sub


'-- 워크리스트 조회 (조회기간 별)
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
    
    Screen.MousePointer = 11
    
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
                    
            Exit Sub
            
        Case "MCC"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "  ORD_YMD AS HOSPDATE"
            SQL = SQL & ", BCODE_NO AS BARCODE"
            SQL = SQL & ", PTNT_NO AS PID"
            SQL = SQL & ", RECEPT_NO AS CHARTNO"
            SQL = SQL & ", '' AS JUMIN"
            SQL = SQL & ", PTNT_NM AS PNAME"
            SQL = SQL & ", SEX AS SEX"
            SQL = SQL & ", AGE AS AGE"
            SQL = SQL & ", COUNT(ORD_CD) AS CNT " & vbCr
            SQL = SQL & "  FROM H7LIS_BCODE_ORD " & vbCr
            SQL = SQL & " WHERE ORD_YMD Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
            SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND RESULT_TYPE = '20'" & vbCr
            SQL = SQL & " GROUP BY ORD_YMD,BCODE_NO,PTNT_NO,RECEPT_NO,PTNT_NM,SEX,AGE " & vbCr
            SQL = SQL & " ORDER BY ORD_YMD,RECEPT_NO,BCODE_NO "
        
        
        Case "BIT"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "  OcmAcpDtm AS HOSPDATE"
            SQL = SQL & ", ResSpmNum AS BARCODE"
            SQL = SQL & ", OcmNum AS PID"
            SQL = SQL & ", PbsChtNum AS CHARTNO"
            SQL = SQL & ", '' AS JUMIN"
            SQL = SQL & ", PbsPatNam AS PNAME"
            SQL = SQL & ", '외래' As INOUT"
            SQL = SQL & ", PbsPatSex AS SEX"
            SQL = SQL & ", PbsPatAge AS AGE"
            SQL = SQL & ", COUNT(ORD_CD) AS CNT " & vbCr
            SQL = SQL & "  FROM OcmInf, PbsInf, RsbInf, ResInf " & vbCrLf
            SQL = SQL & " WHERE PbsChtNum = OcmChtNum " & vbCrLf
            SQL = SQL & "   AND RsbOcmNum = OcmNum " & vbCrLf
            SQL = SQL & "   AND ResOcmNum = OcmNum " & vbCrLf
            SQL = SQL & "   AND RsbAcpCod = ResAcpCod " & vbCrLf
            SQL = SQL & "   AND RsbSpmNum = ResSpmNum " & vbCrLf
            SQL = SQL & "   AND RsbAcpCod = '" & gHOSP.PARTCD & "' " & vbCrLf
            SQL = SQL & "   AND ResSpmNum = '" & sBarcode & "' " & vbCrLf
            SQL = SQL & " UNION ALL " & vbCrLf
            SQL = SQL & "SELECT DISTINCT PbsChtNum, PbsPatNam, PbsResNum, IcmAcpDtm As OcmAcpDtm, 'I' As IOGBN, IcmOcmNum As OcmNum " & vbCrLf
            SQL = SQL & "  FROM IcmInf, PbsInf, RsbInf, ResInf " & vbCrLf
            SQL = SQL & " WHERE PbsChtNum = IcmChtNum " & vbCrLf
            SQL = SQL & "   AND ResOcmNum = IcmOcmNum " & vbCrLf
            SQL = SQL & "   AND RsbOcmNum = IcmOcmNum " & vbCrLf
            SQL = SQL & "   AND RsbAcpCod = ResAcpCod " & vbCrLf
            SQL = SQL & "   AND RsbSpmNum = ResSpmNum " & vbCrLf
            SQL = SQL & "   AND RsbAcpCod = '" & gGumPart & "' " & vbCrLf
            SQL = SQL & "   AND ResSpmNum = '" & sBarcode & "' " & vbCrLf
        
    End Select
    
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
                    
                SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                SetText frmWorkList.spdWork, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                SetText frmWorkList.spdWork, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
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
        
        frmWorkList.spdWork.Row = 1
        frmWorkList.spdWork.Action = ActionActiveCell
    Else
        frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
        frmWorkList.chkAll.Value = "0"
    End If
    
    RS.Close
     
    frmWorkList.spdWork.RowHeight(-1) = 12
    frmWorkList.spdWork.ReDraw = True
    
    Screen.MousePointer = 0

End Sub

'-- 검사결과 조회
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
    
    Screen.MousePointer = 11
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT SAVESEQ,EXAMDATE,HOSPDATE,EQUIPNO,BARCODE,SAMPLETYPE,DISKNO,POSNO" & vbCr
    SQL = SQL & ",CHARTNO,INOUT,PID,PNAME,PSEX,PAGE,PJUMIN,SENDFLAG,SENDDATE,EXAMUID,HOSPITAL " & vbCr
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
        SQL = SQL & "   AND SENDFLAG = '1' " & vbCr
    '-- 미전송
    ElseIf pOpt = 2 Then
        SQL = SQL & "   AND SENDFLAG = '2' " & vbCr
    End If
    
    SQL = SQL & "   AND EXAMCODE IN (" & gAllTestCd & ") " & vbCr
    If pDateType = 0 Then
        SQL = SQL & " ORDER BY EXAMDATE,BARCODE,SAVESEQ "
    Else
        SQL = SQL & " ORDER BY HOSPDATE,BARCODE,SAVESEQ "
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
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                    
                SetText frmMain.spdROrder, "1", intRow, colCHECKBOX
                SetText frmMain.spdROrder, Trim(RS.Fields("SAVESEQ")) & "", intRow, colSAVESEQ
                SetText frmMain.spdROrder, Trim(RS.Fields("EXAMDATE")) & "", intRow, colEXAMDATE
                SetText frmMain.spdROrder, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                SetText frmMain.spdROrder, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                SetText frmMain.spdROrder, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                SetText frmMain.spdROrder, Trim(RS.Fields("PID")) & "", intRow, colPID
                SetText frmMain.spdROrder, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                SetText frmMain.spdROrder, Trim(RS.Fields("PSEX")) & "", intRow, colPSEX
                SetText frmMain.spdROrder, Trim(RS.Fields("PAGE")) & "", intRow, colPAGE
                SetText frmMain.spdROrder, Trim(RS.Fields("PJUMIN")) & "", intRow, colPJUMIN
                'SetText frmMain.spdROrder, GetSampleITEM(intRow), intRow, colITEMS
                
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
    
    Screen.MousePointer = 0

End Sub

'-- 검사자 ITEM 가져오기
Function GetSampleITEM(ByVal asRow As Long) As String
    Dim strBarcode      As String
    Dim strRegDate      As String
    Dim lngExamNo       As Long
    Dim strItems        As String
    
    GetSampleITEM = ""
    
    strBarcode = Trim(GetText(frmWorkList.spdWork, asRow, colBARCODE))
    
    If strBarcode = "" Then
        Exit Function
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
    
        Case "MCC"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM " & vbCr
            SQL = SQL & "  FROM H7LIS_BCODE_ORD " & vbCr
            SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & " ORDER BY ORD_CD "
    
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


'-- 검사자 정보 가져오기
Function GetSampleInfo(ByVal asRow As Long) As Integer
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
    
    GetSampleInfo = -1
    intTestCnt = 0
    
    strBarcode = Trim(GetText(frmMain.spdOrder, asRow, colBARCODE))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    Select Case gOCS
        Case "PHILL"
            '-- 알러지일 경우 바코드번호(전송일자 + 환자번호)를 인터페이스에서 만들어서 전송하기 때문에
            '-- 전송일자,환자번호를 바코드번호에서 찾아와서 조회한다.
            strRegDate = Mid(strBarcode, 1, 8)
            lngExamNo = Val(Mid(strBarcode, 9))
            
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "  a.request_date AS HOSPDATE"
            SQL = SQL & ", a.exam_no AS PID"
            SQL = SQL & ", a.chart_no AS CHARTNO"
            SQL = SQL & ", a.personal_id AS JUMIN"
            SQL = SQL & ", a.person_name AS PNAME"
            SQL = SQL & ", a.person_sex AS SEX"
            SQL = SQL & ", a.person_age AS AGE"
            SQL = SQL & ", b.exam_code AS ITEM " & vbCr
            SQL = SQL & "  FROM TRUST a, TRURES b " & vbCr
            SQL = SQL & " WHERE a.request_date = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   AND a.request_date = b.request_date " & vbCr
            SQL = SQL & "   AND a.exam_no = '" & lngExamNo & "'"
            SQL = SQL & "   AND b.pro_code IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND b.exam_code <> 'X999' " & vbCr
            SQL = SQL & "   AND b.exam_code IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND a.exam_no = b.exam_no " & vbCr
            SQL = SQL & " ORDER BY a.request_date, a.exam_no "
    
        Case "BITHOSNEW"

    
    End Select
            
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With frmMain.spdOrder
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText frmMain.spdOrder, "1", .MaxRows, colCHECKBOX
                SetText frmMain.spdOrder, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText frmMain.spdOrder, Trim(RS.Fields("HOSPDATE")) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0"), asRow, colBARCODE
                SetText frmMain.spdOrder, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText frmMain.spdOrder, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText frmMain.spdOrder, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText frmMain.spdOrder, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText frmMain.spdOrder, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText frmMain.spdOrder, Trim(RS.Fields("JUMIN")) & "", asRow, colPJUMIN
                SetText frmMain.spdOrder, CStr(intTestCnt), asRow, colOCNT
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(frmMain.spdOrder, "◇", asRow, intCol)
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
        SQL = SQL & "   AND EQUIPNO = '" & gHOSP.HOSPCD & "' " & vbCr
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
            SQL = SQL & ", EXAMCODE"                        '병원검사코드"
            SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
            SQL = SQL & ", EXAMNAME"                        '검사명
            SQL = SQL & ", SEQNO" & vbCrLf                  '검사일련번호"
            SQL = SQL & ", SAMPLETYPE"                      '검체유형"
            SQL = SQL & ", INOUT"                           '입/외
            SQL = SQL & ", DISKNO"                          'Rack
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
            
            If Not DBExec(AdoCn_Local, SQL) Then
                'SaveQuery SQL
                Exit Function
            End If
            
        End If
        
    End With
    
End Function

'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
Public Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
          SQL = "SELECT MAX(SAVESEQ) as SEQ FROM PATRESULT  "
    SQL = SQL & " WHERE MID(EXAMDATE,1,8) = '" & strDate & "' " & vbCrLf
    
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        getMaxTestNum = Trim(RS.Fields("SEQ")) + 1
    Else
        getMaxTestNum = 1
    End If
    
    If getMaxTestNum >= 99999 Then
        getMaxTestNum = 99999
    End If
    
    RS.Close
    
End Function

