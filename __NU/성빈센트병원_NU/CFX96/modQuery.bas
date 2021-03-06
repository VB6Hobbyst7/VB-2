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
                Call SetText(frmMain.spdTest, IIf(AdoRs_Local.Fields("CUTOFFUSE").Value & "" = "Y", "1", "0"), intRow, colLCUTUSE)
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
                        gAllTestCd1 = AdoRs_Local.Fields("TESTCODE").Value
                    Else
                        gAllTestCd = gAllTestCd & ",'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                        gAllTestCd1 = gAllTestCd1 & "," & AdoRs_Local.Fields("TESTCODE").Value
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
                        gAllOrdCd1 = AdoRs_Local.Fields("ORDERCODE").Value
                    Else
                        gAllOrdCd = gAllOrdCd & ",'" & AdoRs_Local.Fields("ORDERCODE").Value & "'"
                        gAllOrdCd1 = gAllOrdCd1 & "," & AdoRs_Local.Fields("ORDERCODE").Value
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

'-- 사용자 이름 조회
Public Function GetUserNm(ByVal PID As String) As String
    
    GetUserNm = ""
    
    SQL = ""
    SQL = SQL & "SELECT usab.instcd,"               '-- 기관기호
    SQL = SQL & "       usab.USERID,"               '-- 사용자아이디
    SQL = SQL & "       usab.USERABBR" & vbCr       '-- 사용자성명
    SQL = SQL & "  FROM lis.lpcmusab usab,"         '-- 병리과 사용자관리
    SQL = SQL & "       com.zsumusrb usrb" & vbCr   '-- 시스템 사용자정보
    SQL = SQL & " WHERE usab.instcd     = ? " & vbCr
    SQL = SQL & "   AND usab.delflagcd  = ? " & vbCr
    SQL = SQL & "   AND usab.userid     = ? " & vbCr
    SQL = SQL & "   AND usab.userid     = usrb.userid" & vbCr
    SQL = SQL & "   AND to_char(sysdate,'yyyymmdd') BETWEEN usrb.userfromdd AND usrb.usertodd "
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("usab.instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("usab.delflagcd", adVarChar, adParamInput, 1000, "0")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("usab.userid", adVarChar, adParamInput, 1000, PID)
    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic

    Call SetSQLData("로그인", SQL)

    If AdoRs.BOF = False Then
        GetUserNm = AdoRs.Fields("USERABBR").Value & ""
    Else
        GetUserNm = ""
    End If
    
    Set AdoCmd = Nothing
    Set AdoRs = Nothing
    
End Function

'-- 시스템 날짜 가져오기
Public Function GetSysDtTm() As Boolean
    
    GetSysDtTm = False
    
    SQL = ""
    SQL = SQL & "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') AS sysdd"
    SQL = SQL & "     , TO_CHAR(SYSDATE,'HH24MISS') AS systm"
    SQL = SQL & "  FROM DUAL"
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic

    Call SetSQLData("결과저장", SQL, "A")
    
    If AdoRs.BOF = False Then
        gSysDate = AdoRs.Fields("sysdd").Value & ""
        gSysTime = AdoRs.Fields("systm").Value & ""
        GetSysDtTm = True
    Else
        GetSysDtTm = False
    End If
    
    Set AdoCmd = Nothing
    Set AdoRs = Nothing
    
End Function

'-- 워크리스트 조회
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String, Optional ByVal pState As String)
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
    
    
On Error GoTo RST
    
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
            SQL = SQL & " ORDER BY a.chart_no,a.request_date, a.exam_no "
        
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
            SQL = SQL & "  RECEIPTDATE AS HOSPDATE"
            SQL = SQL & ", SPECIMENNUM AS BARCODE"
            SQL = SQL & ", RECEIPTNO AS CHARTNO"
            SQL = SQL & ", IPDOPD AS INOUT "
            SQL = SQL & ", PTNO AS PID"
            SQL = SQL & ", SNAME AS PNAME"
            SQL = SQL & ", COUNT(LABCODE) AS CNT " & vbCr
            SQL = SQL & "   FROM SLA_LabMaster " & vbCr
            SQL = SQL & " WHERE RECEIPTDATE between '" & Format(pFrom, "####-##-##") & "' and '" & Format(pTo, "####-##-##") & "'" & vbCr
            SQL = SQL & "   AND ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND LABCODE IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND JSTATUS < '3'" & vbCr
            SQL = SQL & "  ORDER BY RECEIPTDATE "
                
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
                        SetText frmWorkList.spdWork, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                        SetText frmWorkList.spdWork, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                        SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
                        SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                        'SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                        'SetText frmWorkList.spdWork, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                        'SetText frmWorkList.spdWork, Trim(RS.Fields("JUMIN")) & "", intRow, colPJUMIN
                        SetText frmWorkList.spdWork, IIf(Trim(RS.Fields("INOUT")) & "" = "I", "입원", "외래"), intRow, colINOUT
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
                    
        Case "NU"
            SQL = ""
            SQL = SQL & "SELECT /*+ leading(acpt) */ "
            SQL = SQL & "  ACPT.INSTCD "                        '기관기호
            SQL = SQL & ", ACPT.ACPTDD AS HOSPDATE"             '접수일자
            SQL = SQL & ", ACPT.ACPTNO AS BARCODE"              '접수번호
            SQL = SQL & ", ACPT.ACPTITEMNO AS SEQ"              '접수항목번호
            SQL = SQL & ", ACPT.ptno AS CHARTNO"                '병리번호
            SQL = SQL & ", ACPT.acptstatcd as STATCD "          '접수상태코드(0:접수(440),1:취소(100,240),2:예비결과(710),3:최종결과(730), 4:수정결과(740))"
            SQL = SQL & ", ACPT.pid AS PID"                     '등록번호
            SQL = SQL & ", ACPT.prcpgenrflag AS INOUT"          '입원/외래구분
            SQL = SQL & ", ACPT.ORDDEPTCD AS RACKNO"            '진료과코드
            SQL = SQL & ", DEPT.deptengabbr AS POSNO"           '진료과명
            SQL = SQL & ", ACPT.PRCPDD AS KEY1"                 '처방일자
            SQL = SQL & ", ACPT.EXECPRCPUNIQNO AS KEY2"         '실시처방유일번호
            SQL = SQL & ", ACPT.prcpno AS JUMIN"                '처방번호
            SQL = SQL & ", PTBS.hngnm AS PNAME"                 '환자명
            SQL = SQL & ", PTBS.SEX AS SEX"                     '성별
            SQL = SQL & ", PTBS.BRTHDD AS DOB"                  '생일
            SQL = SQL & ", com.fn_zz_getage(ptbs.rrgstno1, ptbs.rrgstno2, acpt.acptdd, 'A', ptbs.brthdd) as AGE, " & vbCr  '접수일자기준 나이"
            
            SQL = SQL & "( SELECT min(rslt.rsltrgstno)" & vbCr
            SQL = SQL & "    FROM lis.lprmrslt rslt" & vbCr
            SQL = SQL & "   Where ACPT.instcd = rslt.instcd" & vbCr
            SQL = SQL & "     AND ACPT.ptno   = rslt.ptno" & vbCr
            SQL = SQL & "     AND ACPT.pid    = rslt.pid" & vbCr
            SQL = SQL & "     AND rslt.delflagcd = '0'" & vbCr
            SQL = SQL & "     AND rslt.rsltrgsthistno = 1" & vbCr
            SQL = SQL & ") as rsltrgstno" & vbCr
            
            SQL = SQL & "  FROM lis.lpjmacpt ACPT, lis.lpcmtest TEST, lis.lpcmspcm SPCM, pam.pmcmptbs PTBS, com.zsdddept DEPT" & vbCr
            SQL = SQL & " WHERE ACPT.instcd= ? " & vbCr                        '성빈센트병원(instcd = 017)는 고정입니다.
            SQL = SQL & "   AND ACPT.acptdd BETWEEN ? AND ? " & vbCr
            SQL = SQL & "   AND ACPT.testcd IN (?) " & vbCr                       '검사처방코드 : PMO12040 고정 GFX96 에서 다른처방 처리시 in 절로 변경
            
            'SQL = SQL & "   AND ACPT.acptstatcd = 0 " & vbCr                                     '접수상태코드 : (0:접수,1:취소,2:예비결과,3:최종결과, 4:수정결과 고정
            SQL = SQL & "   AND ACPT.acptstatcd IN (?) " & vbCr                                     '접수상태코드 : (0:접수,1:취소,2:예비결과,3:최종결과, 4:수정결과 고정
            
            SQL = SQL & "   AND ACPT.instcd          = TEST.instcd" & vbCr
            SQL = SQL & "   AND ACPT.testcd          = TEST.testcd" & vbCr
            SQL = SQL & "   AND ACPT.instcd          = SPCM.instcd" & vbCr
            SQL = SQL & "   AND ACPT.spccd           = SPCM.spccd" & vbCr
            SQL = SQL & "   AND ACPT.instcd          = PTBS.instcd" & vbCr
            SQL = SQL & "   AND ACPT.pid             = PTBS.pid" & vbCr
            SQL = SQL & "   AND ACPT.instcd          = DEPT.instcd" & vbCr
            SQL = SQL & "   AND ACPT.orddeptcd       = DEPT.deptcd" & vbCr
            SQL = SQL & "   AND ACPT.prcpdd BETWEEN DEPT.valifromdd AND DEPT.valitodd"
            SQL = SQL & " ORDER BY ACPT.ptno"
            
            Call SetSQLData("워크조회", SQL)
            frmWorkList.txtQuery.Text = SQL
            
            Set AdoCmd = New ADODB.Command
            Set AdoCmd.ActiveConnection = AdoCn
            AdoCmd.CommandType = adCmdText
            AdoCmd.CommandText = SQL
            AdoCmd.Parameters.Append AdoCmd.CreateParameter("acpt.instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
            AdoCmd.Parameters.Append AdoCmd.CreateParameter("acpt.acptdd", adVarChar, adParamInput, 1000, pFrom)
            AdoCmd.Parameters.Append AdoCmd.CreateParameter("acpt.acptdd", adVarChar, adParamInput, 1000, pTo)
            AdoCmd.Parameters.Append AdoCmd.CreateParameter("acpt.testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
            AdoCmd.Parameters.Append AdoCmd.CreateParameter("acpt.acptstatcd", adVarChar, adParamInput, 1000, pState)
            
            Set AdoRs = New ADODB.Recordset
            AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic
        
            If Not AdoRs.EOF = True And Not AdoRs.BOF = True Then
                frmWorkList.spdWork.MaxRows = 0
                Do Until AdoRs.EOF
                    iCnt = iCnt + 1
                    With frmWorkList.spdWork
                        .ReDraw = False
                        .MaxRows = .MaxRows + 1
                        intRow = .MaxRows
                        'strBarcode = Trim(AdoRs.Fields("HOSPDATE")) & PedLeftStr(Trim(AdoRs.Fields("PID")), 5, "0")
                            
                        SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("BARCODE")) & "", intRow, colBARCODE
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("CHARTNO")) & "", intRow, colCHARTNO
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("PID")) & "", intRow, colPID
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("PNAME")) & "", intRow, colPNAME
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("SEQ")) & "", intRow, colSEQNO
                            
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("RACKNO")) & "", intRow, colRACKNO  '진료과코드
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("POSNO")) & "", intRow, colPOSNO    '진료과명
                        
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("KEY1")) & "", intRow, colKEY1      '처방일자
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("KEY2")) & "", intRow, colKEY2      '실시처방유일번호
                        
'                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("JUMIN")) & "", intRow, colPJUMIN   '처방번호
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("rsltrgstno")) & "", intRow, colPJUMIN   '처방번호
                        
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("SEX")) & "", intRow, colPSEX
                        SetText frmWorkList.spdWork, Trim(AdoRs.Fields("AGE")) & "", intRow, colPAGE
                        
                        SetText frmWorkList.spdWork, IIf(Trim(AdoRs.Fields("INOUT")) & "" = "I", "입원", "외래"), intRow, colINOUT
                    End With
                    DoEvents
                    
                    AdoRs.MoveNext
                Loop
                frmWorkList.chkAll.Value = "1"
            End If
            
            Set AdoCmd = Nothing
            Set AdoRs = Nothing
            
            
            
'
'            '-- Record Count 가져옴
'            AdoCn.CursorLocation = adUseClient
'            Set RS = AdoCn.Execute(SQL, , 1)
'            If Not RS.EOF = True And Not RS.BOF = True Then
'                frmWorkList.spdWork.MaxRows = 0
'                strItems = ""
'                Do Until RS.EOF
'                    iCnt = iCnt + 1
'                    With frmWorkList.spdWork
'                        .ReDraw = False
'                        .MaxRows = .MaxRows + 1
'                        intRow = .MaxRows
'                        strBarcode = Trim(RS.Fields("HOSPDATE")) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0")
'
'                        SetText frmWorkList.spdWork, "1", intRow, colCHECKBOX
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("PID")) & "", intRow, colPID
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
'
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("RACKNO")) & "", intRow, colRACKNO  '진료과코드
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("POSNO")) & "", intRow, colPOSNO    '진료과명
'
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("KEY1")) & "", intRow, colKEY1      '처방일자
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("KEY2")) & "", intRow, colKEY2      '실시처방유일번호
'
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("JUMIN")) & "", intRow, colPJUMIN   '처방번호
'
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
'                        SetText frmWorkList.spdWork, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
'
'                        SetText frmWorkList.spdWork, IIf(Trim(RS.Fields("INOUT")) & "" = "I", "입원", "외래"), intRow, colINOUT
'                        'SetText frmWorkList.spdWork, frmWorkList.txtSeq.Text, intRow, colSEQNO
'                        'SetText frmWorkList.spdWork, RS.Fields("CNT"), intRow, colOCNT
'                        'SetText frmWorkList.spdWork, GetSampleITEM(intRow), intRow, colITEMS
'
'                        'frmWorkList.txtSeq.Text = frmWorkList.txtSeq.Text + 1
'                    End With
'                    DoEvents
'
'                    RS.MoveNext
'                Loop
'                frmWorkList.chkAll.Value = "1"
'            Else
'                frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
'                frmWorkList.chkAll.Value = "0"
'            End If
'
'            RS.Close
'
    End Select

     
    frmWorkList.spdWork.RowHeight(-1) = 12
    frmWorkList.spdWork.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMsg = "위    치 : " & "GetWorkList" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal
    
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
    SQL = SQL & ",CHARTNO,INOUT,PID,PNAME,PSEX,PAGE,PJUMIN,SENDFLAG,SENDDATE,EXAMUID,HOSPITAL,REFJUDGE,REFVALUE " & vbCr
    '-- 검사결과
    SQL = SQL & ",SEQNO,EXAMNAME,RESULT" & vbCr
    
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
                
'                Debug.Print Trim(RS.Fields("SAVESEQ"))
'                Debug.Print Trim(RS.Fields("EXAMDATE"))
                If strSaveSeq <> Trim(RS.Fields("SAVESEQ")) & "" Or strExamDate <> Trim(RS.Fields("EXAMDATE")) & "" Then
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
                    
                    SetText frmMain.spdROrder, Trim(RS.Fields("REFJUDGE")) & "", intRow, colOCNT
                    SetText frmMain.spdROrder, Trim(RS.Fields("REFVALUE")) & "", intRow, colRCNT
                    
                    
                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                    Case "0"
                            SetText frmMain.spdROrder, "장비결과", intRow, colSTATE
                    Case "1"
                            SetText frmMain.spdROrder, "전송완료", intRow, colSTATE
                    End Select
                    'SetText frmMain.spdROrder, GetSampleITEM(intRow), intRow, colITEMS
                
                End If
                
                For intCol = colSTATE + 1 To .MaxCols
                    .Row = 0
                    .Col = intCol
                    If Trim(RS.Fields("EXAMNAME")) & "" = Trim(.Text) Then
                        SetText frmMain.spdROrder, Trim(RS.Fields("RESULT")) & "", intRow, intCol
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
     
    frmMain.spdROrder.RowHeight(-1) = 16
    frmMain.spdROrder.ReDraw = True
    
    Call frmMain.GetPatTRestResult_Search(1)
    
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
    
        Case "JWINFO"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT LABCODE AS ITEM " & vbCr
            SQL = SQL & "   FROM SLA_LabMaster " & vbCr
            SQL = SQL & " WHERE ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   AND LABCODE IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND SPECIMENNUM = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND JSTATUS < '3'" & vbCr
            SQL = SQL & "  ORDER BY LABCODE "
    
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
Function GetSampleInfo(ByVal asRow As Long, ByVal SPD As vaSpread, Optional ByVal pBarno As String = "", Optional ByVal pState As String) As Integer
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
    
On Error GoTo DBErr
    
    GetSampleInfo = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    If pBarno = "" Then
        strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
        If Len(strBarcode) <> 9 Then
            strBarcode = Trim(GetText(SPD, asRow, colCHARTNO))
        End If
    Else
        strBarcode = pBarno
        SPD.MaxRows = SPD.MaxRows + 1
        asRow = SPD.MaxRows
    End If
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    

    SQL = ""
    SQL = SQL & "SELECT /*+ leading(acpt) */ "
    SQL = SQL & "  ACPT.INSTCD "                       '기관기호
    SQL = SQL & ", ACPT.ACPTDD AS HOSPDATE"            '접수일자
    SQL = SQL & ", ACPT.ACPTNO AS BARCODE"             '접수번호
    SQL = SQL & ", ACPT.ACPTITEMNO AS SEQ"             '접수항목번호
    SQL = SQL & ", ACPT.ptno AS CHARTNO"               '병리번호
    SQL = SQL & ", ACPT.pid AS PID"                    '등록번호
    SQL = SQL & ", ACPT.testcd AS ITEM"                '검사코드
    'SQL = SQL & ", TEST.testengnm"                     '영문검사명
    'SQL = SQL & ", SPCM.spcnm"                         '검체명
    SQL = SQL & ", ACPT.prcpgenrflag AS INOUT"         '입원/외래구분
    SQL = SQL & ", ACPT.ORDDEPTCD AS RACKNO"           '진료과코드
    SQL = SQL & ", DEPT.deptengabbr AS POSNO"          '진료과명
    SQL = SQL & ", ACPT.PRCPDD AS KEY1"                '처방일자
    SQL = SQL & ", ACPT.EXECPRCPUNIQNO AS KEY2"        '실시처방유일번호
    SQL = SQL & ", ACPT.prcpno AS JUMIN"               '처방번호
    SQL = SQL & ", PTBS.hngnm AS PNAME"                '환자명
    SQL = SQL & ", PTBS.SEX AS SEX"                    '성별
    SQL = SQL & ", PTBS.BRTHDD AS DOB"                 '생일
    SQL = SQL & ", com.fn_zz_getage(ptbs.rrgstno1, ptbs.rrgstno2, acpt.acptdd, 'A', ptbs.brthdd) as AGE, " & vbCr  '접수일자기준 나이"
    SQL = SQL & "( SELECT min(rslt.rsltrgstno)" & vbCr
    SQL = SQL & "    FROM lis.lprmrslt rslt" & vbCr
    SQL = SQL & "   Where ACPT.instcd = rslt.instcd" & vbCr
    SQL = SQL & "     AND ACPT.ptno   = rslt.ptno" & vbCr
    SQL = SQL & "     AND ACPT.pid    = rslt.pid" & vbCr
    SQL = SQL & "     AND rslt.delflagcd = '0'" & vbCr
    SQL = SQL & "     AND rslt.rsltrgsthistno = 1" & vbCr
    SQL = SQL & ") as rsltrgstno" & vbCr
    SQL = SQL & "  FROM lis.lpjmacpt ACPT, lis.lpcmtest TEST, lis.lpcmspcm SPCM, pam.pmcmptbs PTBS, com.zsdddept DEPT" & vbCr
    SQL = SQL & " WHERE ACPT.instcd= ? " & vbCr                        '성빈센트병원(instcd = 017)는 고정입니다.
    SQL = SQL & "   AND ACPT.ptno = ? " & vbCr
    SQL = SQL & "   AND ACPT.testcd IN (?) " & vbCr                       '검사처방코드 : PMO12040 고정 GFX96 에서 다른처방 처리시 in 절로 변경

'    SQL = SQL & "   AND ACPT.acptstatcd = 0 " & vbCr                                     '접수상태코드 : (0:접수,1:취소,2:예비결과,3:최종결과, 4:수정결과 고정
    SQL = SQL & "   AND ACPT.acptstatcd IN (?) " & vbCr                                     '접수상태코드 : (0:접수,1:취소,2:예비결과,3:최종결과, 4:수정결과 고정
    
    SQL = SQL & "   AND ACPT.instcd          = TEST.instcd" & vbCr
    SQL = SQL & "   AND ACPT.testcd          = TEST.testcd" & vbCr
    SQL = SQL & "   AND ACPT.instcd          = SPCM.instcd" & vbCr
    SQL = SQL & "   AND ACPT.spccd           = SPCM.spccd" & vbCr
    SQL = SQL & "   AND ACPT.instcd          = PTBS.instcd" & vbCr
    SQL = SQL & "   AND ACPT.pid             = PTBS.pid" & vbCr
    SQL = SQL & "   AND ACPT.instcd          = DEPT.instcd" & vbCr
    SQL = SQL & "   AND ACPT.orddeptcd       = DEPT.deptcd" & vbCr
    SQL = SQL & "   AND ACPT.prcpdd BETWEEN DEPT.valifromdd AND DEPT.valitodd"
        
    Call SetSQLData("바코드조회", SQL)
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("A.instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("A.ACPTNO", adVarChar, adParamInput, 1000, strBarcode)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("A.testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ACPT.acptstatcd", adVarChar, adParamInput, 1000, pState)
    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic
               
    If Not AdoRs.EOF = True And Not AdoRs.BOF = True Then
        Do Until AdoRs.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(AdoRs.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(AdoRs.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(AdoRs.Fields("SEQ")), asRow, colSEQNO    '접수항목번호
                SetText SPD, Trim(AdoRs.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(AdoRs.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(AdoRs.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(AdoRs.Fields("RACKNO")) & "", asRow, colRACKNO  '진료과코드
                SetText SPD, Trim(AdoRs.Fields("POSNO")) & "", asRow, colPOSNO    '진료과명
                SetText SPD, Trim(AdoRs.Fields("KEY1")) & "", asRow, colKEY1      '처방일자
                SetText SPD, Trim(AdoRs.Fields("KEY2")) & "", asRow, colKEY2      '실시처방유일번호
                SetText SPD, Trim(AdoRs.Fields("JUMIN")) & "", asRow, colPJUMIN   '처방번호
                SetText SPD, Trim(AdoRs.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(AdoRs.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, IIf(Trim(AdoRs.Fields("INOUT")) & "" = "I", "입원", "외래"), asRow, colINOUT
                SetText SPD, Trim(AdoRs.Fields("rsltrgstno")) & "", asRow, colOCNT
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(AdoRs.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        Exit For
                    End If
                Next
                gPatOrdCd = gPatOrdCd & "'" & Trim(AdoRs.Fields("ITEM")) & "',"
                
            End With
            DoEvents
            
            AdoRs.MoveNext
        Loop
    End If
    
    AdoRs.Close
            

    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_Save(ByVal asRow As Long, ByVal SPD As vaSpread, Optional ByVal pBarno As String = "", Optional ByVal pState As String) As Integer
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
    
On Error GoTo DBErr
    
    GetSampleInfo_Save = -1
    intTestCnt = 0
    gPatOrdCd = ""
    
    If pBarno = "" Then
        strBarcode = Trim(GetText(SPD, asRow, colCHARTNO))
    Else
        strBarcode = pBarno
        SPD.MaxRows = SPD.MaxRows + 1
        asRow = SPD.MaxRows
    End If
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    

    SQL = ""
    SQL = SQL & "SELECT /*+ leading(acpt) */ "
    SQL = SQL & "  ACPT.INSTCD "                       '기관기호
    SQL = SQL & ", ACPT.ACPTDD AS HOSPDATE"            '접수일자
    SQL = SQL & ", ACPT.ACPTNO AS BARCODE"             '접수번호
    SQL = SQL & ", ACPT.ACPTITEMNO AS SEQ"             '접수항목번호
    SQL = SQL & ", ACPT.ptno AS CHARTNO"               '병리번호
    SQL = SQL & ", ACPT.pid AS PID"                    '등록번호
    SQL = SQL & ", ACPT.testcd AS ITEM"                        '검사코드
    SQL = SQL & ", ACPT.prcpgenrflag AS INOUT"         '입원/외래구분
    SQL = SQL & ", ACPT.ORDDEPTCD AS RACKNO"           '진료과코드
    SQL = SQL & ", DEPT.deptengabbr AS POSNO"          '진료과명
    SQL = SQL & ", ACPT.PRCPDD AS KEY1"                '처방일자
    SQL = SQL & ", ACPT.EXECPRCPUNIQNO AS KEY2"        '실시처방유일번호
    SQL = SQL & ", ACPT.prcpno AS JUMIN"               '처방번호
    SQL = SQL & ", PTBS.hngnm AS PNAME"                '환자명
    SQL = SQL & ", PTBS.SEX AS SEX"                    '성별
    SQL = SQL & ", PTBS.BRTHDD AS DOB"                 '생일
    SQL = SQL & ", com.fn_zz_getage(ptbs.rrgstno1, ptbs.rrgstno2, acpt.acptdd, 'A', ptbs.brthdd) as AGE, " & vbCr  '접수일자기준 나이"
    SQL = SQL & "( SELECT min(rslt.rsltrgstno)" & vbCr
    SQL = SQL & "    FROM lis.lprmrslt rslt" & vbCr
    SQL = SQL & "   Where ACPT.instcd = rslt.instcd" & vbCr
    SQL = SQL & "     AND ACPT.ptno   = rslt.ptno" & vbCr
    SQL = SQL & "     AND ACPT.pid    = rslt.pid" & vbCr
    SQL = SQL & "     AND rslt.delflagcd = '0'" & vbCr
    SQL = SQL & "     AND rslt.rsltrgsthistno = 1" & vbCr
    SQL = SQL & ") as rsltrgstno" & vbCr
    SQL = SQL & ", ACPT.acptstatcd " & vbCr
    SQL = SQL & "  FROM lis.lpjmacpt ACPT, lis.lpcmtest TEST, lis.lpcmspcm SPCM, pam.pmcmptbs PTBS, com.zsdddept DEPT" & vbCr
    SQL = SQL & " WHERE ACPT.instcd= ? " & vbCr                        '성빈센트병원(instcd = 017)는 고정입니다.
    SQL = SQL & "   AND ACPT.ptno = ? " & vbCr
    SQL = SQL & "   AND ACPT.testcd IN (?) " & vbCr                       '검사처방코드 : PMO12040 고정 GFX96 에서 다른처방 처리시 in 절로 변경
    
    'SQL = SQL & "   AND ACPT.acptstatcd = 0 " & vbCr                                     '접수상태코드 : (0:접수,1:취소,2:예비결과,3:최종결과, 4:수정결과 고정
    SQL = SQL & "   AND ACPT.acptstatcd IN (?) " & vbCr                                     '접수상태코드 : (0:접수,1:취소,2:예비결과,3:최종결과, 4:수정결과 고정
    
    SQL = SQL & "   AND ACPT.instcd          = TEST.instcd" & vbCr
    SQL = SQL & "   AND ACPT.testcd          = TEST.testcd" & vbCr
    SQL = SQL & "   AND ACPT.instcd          = SPCM.instcd" & vbCr
    SQL = SQL & "   AND ACPT.spccd           = SPCM.spccd" & vbCr
    SQL = SQL & "   AND ACPT.instcd          = PTBS.instcd" & vbCr
    SQL = SQL & "   AND ACPT.pid             = PTBS.pid" & vbCr
    SQL = SQL & "   AND ACPT.instcd          = DEPT.instcd" & vbCr
    SQL = SQL & "   AND ACPT.orddeptcd       = DEPT.deptcd" & vbCr
    SQL = SQL & "   AND ACPT.prcpdd BETWEEN DEPT.valifromdd AND DEPT.valitodd"
                
                
    Call SetSQLData("바코드조회", SQL)
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("A.instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("A.ACPTNO", adVarChar, adParamInput, 1000, strBarcode)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("A.testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ACPT.acptstatcd", adVarChar, adParamInput, 1000, pState)
    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic
               
    If Not AdoRs.EOF = True And Not AdoRs.BOF = True Then
        Do Until AdoRs.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(AdoRs.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(AdoRs.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(AdoRs.Fields("SEQ")), asRow, colSEQNO    '접수항목번호
                SetText SPD, Trim(AdoRs.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(AdoRs.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(AdoRs.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(AdoRs.Fields("RACKNO")) & "", asRow, colRACKNO  '진료과코드
                SetText SPD, Trim(AdoRs.Fields("POSNO")) & "", asRow, colPOSNO    '진료과명
                SetText SPD, Trim(AdoRs.Fields("KEY1")) & "", asRow, colKEY1      '처방일자
                SetText SPD, Trim(AdoRs.Fields("KEY2")) & "", asRow, colKEY2      '실시처방유일번호
                SetText SPD, Trim(AdoRs.Fields("JUMIN")) & "", asRow, colPJUMIN   '처방번호
                SetText SPD, Trim(AdoRs.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(AdoRs.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, IIf(Trim(AdoRs.Fields("INOUT")) & "" = "I", "입원", "외래"), asRow, colINOUT
                
                SetText SPD, Trim(AdoRs.Fields("rsltrgstno")) & "", asRow, colOCNT
                SetText SPD, Trim(AdoRs.Fields("acptstatcd")) & "", asRow, colRCNT
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(AdoRs.Fields("ITEM")) & "',"
                
            End With
            DoEvents
            
            AdoRs.MoveNext
        Loop
    End If
    
    AdoRs.Close
            

    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_Save = -1
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
            SQL = SQL & ", ORDERCODE"                       '병원처방코드"
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
            
'Public Const colSEQNO = 6       '일련번호                       ==> 접수항목번호
'Public Const colKEY1 = 16       '여분1                          ==> 처방일자
'Public Const colKEY2 = 17       '여분2                          ==> 실시처방유일번호
'Public Const colRCNT = 19       '결과갯수                       ==> 검체명
            
            
            
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "'" & vbCrLf      '==> 접수번호
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRORDERCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRSUBCD)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTNM)) & "'"              '==> 검사명
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRSEQNO)) & "'" & vbCrLf
            SQL = SQL & ",''"                                                   '검체유형
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colINOUT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colRACKNO)) & "'"                '==> 진료과코드
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPOSNO)) & "'"                 '==> 진료과명
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRMACHRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRLISRESULT)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRJUDGE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRFLAG)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRREF)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colCHARTNO)) & "'"               '==> 병리번호
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPID)) & "'"                   '==> 등록번호
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPNAME)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPSEX)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPAGE)) & "'"
            SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPJUMIN)) & "'"                '==> 처방번호
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

