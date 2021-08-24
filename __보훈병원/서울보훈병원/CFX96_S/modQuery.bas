Attribute VB_Name = "modQuery"
Option Explicit

Public SQL              As String
Public RS               As ADODB.Recordset
Public blnSameRecord    As Boolean



'한 장비채널에 검사코드가 1개이상 존재 (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_AU680(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_AU680 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
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

'한 장비채널에 검사코드가 1개이상 존재 (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_HITACHI7180(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim intIntBase      As Integer
    Dim strItems        As String           '전송할 검사항목
    Dim blnISE          As Boolean          'Na, K, Cl 검사여부

    strItems = String$(88, "0")
    GetEquipExamCode_HITACHI7180 = strItems
    strExamCode = ""
    blnISE = False
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
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
                    '## 계산항목: 93~100
                    If intIntBase >= 93 And intIntBase <= 100 Then
                        'GoTo Skip1
                    Else
                        '## Na, K, Cl 검사여부 Check
                        If intIntBase = 87 Or intIntBase = 88 Or intIntBase = 89 Then
                            blnISE = True
                        Else
                            Mid$(strItems, intIntBase, 1) = "1"
                        End If
                    End If
                    
                    '## TIBC 이면 UIBC,FE 오더를 준다.
                    'If lngIntBase = 98 Then
                    '    Mid$(strItems, 22, 1) = "1"     'FE
                    '    Mid$(strItems, 23, 1) = "1"     'UIBC
                    'End If
                            
                    '## B/C  (025)항목은 계산항목이라 오더를 보내면 안됨(BUN,CREA)
                    '## A/G  (026)항목은 계산항목이라 오더를 보내면 안됨
                    '## GLOB (032)항목은 계산항목이라 오더를 보내면 안됨
                    '## I.Bil(033)항목은 계산항목이라 오더를 보내면 안됨
                    '## T.P  (002)항목은 검체가 Urine일때 검사를 하면 안됨
                    '## HbA1C(23)항목은 Hgb(20)와 A1C(21) 오더를 보내야 함
                    '## LDL-C(99)항목은 계산항목이라 오더를 보내면 안됨(CHOL, T.G, HDL-C)
                    mOrder.SendCnt = mOrder.SendCnt + 1
                End If
            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    '## Na, K, Cl 검사여부 Check
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
        
        '-- Local에서 환자별로 결과값 가져오기
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
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    SQL = "" '
                    SQL = SQL & "Update TB_H141_LISTAKEBODY                     " & vbCrLf
                    SQL = SQL & "   SET H141_RSLTYN    ='Y'                     " & vbCrLf
                    SQL = SQL & " WHERE H141_TSAMPLENO = '" & strBarcode & "'   " & vbCrLf
                    SQL = SQL & "   AND H141_SUGACD    = '" & strTestCd & "'    " & vbCrLf
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                    
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H131_SPPRESULT                       " & vbCrLf
                    SQL = SQL & "   SET H131_RESULT  = '" & sResult & "'        " & vbCrLf
                    SQL = SQL & " WHERE H131_SPPTYPE = '" & gHOSP.PARTCD & "'   " & vbCrLf    'L010
                    SQL = SQL & "   AND H131_SEQNO   = '" & strTestCdSub & "'   " & vbCrLf
                        
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H130_SPPRECEIVE                              " & vbCrLf
                    SQL = SQL & "   SET H130_RSLTDAT = TO_CHAR(SYSDATE, 'YYYYMMDD')     " & vbCrLf
                    SQL = SQL & "      ,H130_RSLTTM  = TO_CHAR(SYSDATE, 'HH24:MI:SS')   " & vbCrLf
                    SQL = SQL & " WHERE H130_SPPTYPE = '" & gHOSP.PARTCD & "'           " & vbCrLf    'L010
                    SQL = SQL & "   AND H130_SEQNO   = '" & strTestCdSub & "'           " & vbCrLf
                        
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H140_LISTAKEHEAD                     " & vbCrLf
                    SQL = SQL & "   SET H140_RSLTYN    = 'Y'                    " & vbCrLf
                    SQL = SQL & " WHERE H140_TSAMPLENO = '" & strBarcode & "'   " & vbCrLf
                                        
                    Call SetSQLData("결과저장", SQL, "A")
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SaveTransData_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
'        '-- Local에서 환자별로 결과값 가져오기
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
'                '-- 장비결과적용
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
'            sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"     '사용자ID|기관코드
'            sParam = sParam & "bcno=" & strBarcode & "&"                                    '검체번호(바코드)
'            sParam = sParam & "result=" & strAllResult & "&"                                '결과
'            sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '기관코드
'            sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '장비코드
'            sParam = sParam & "userid=" & gHOSP.USERID & "&"                                '사용자ID
'
'            sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'            Call SetSQLData("결과저장", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
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
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SaveTransData_NU" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
'        '-- Local에서 환자별로 결과값 가져오기
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
'                '-- 장비결과적용
'                If gHOSP.SAVELIS = "Y" Then
'                    sResult = sResult2
'                Else
'                    sResult = sResult1
'                End If
'
''SERVERIP "/himed2/.live?submit_id=" + argId + "&business_id=lis&bcno=" + argBarcode + "&result=" + argResult + "&eqmtcd=" + strLIS_EQCD + "&instcd=053&userid=LISBC&paste=Y&retestyn=N&nmeddilute=0"
'' -> (서버IP)/himed2/.live?
''submit_id=TXLII00101&
''business_id=lis&
''bcno=(바코드번호)&
''result=(결과:검사코드%17결과%17%17입력시간%171%03·····검사코드%17결과%17%17입력시간%171)&eqmtcd=(장비코드)&
''instcd=053&
''userid=LISBC&
''paste=Y&
''retestyn=N&
''nmeddilute=0
'
''JC메디컴
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
''            sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"     '사용자ID|기관코드
'            sParam = sParam & "bcno=" & strBarcode & "&"                                    '검체번호(바코드)
'            sParam = sParam & "result=" & strAllResult & "&"                                '결과
'            sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '장비코드
'            sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '기관코드
'            sParam = sParam & "userid=" & gHOSP.USERID & "&"                                '사용자ID
'            sParam = sParam & "paste=Y&"
'            sParam = sParam & "retestyn=N&"
'            sParam = sParam & "nmeddilute=0"
'
'            sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'            Call SetSQLData("결과저장", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
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
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_HDINFO" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
    strCmnt = "실시간 역전사 종합효소연쇄반응법(Realtime reverse transcriptase PCR)"
    
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
        
        '-- Local에서 환자별로 결과값 가져오기
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
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If

                If sResult <> "" Then
                    '-- 검사결과저장 = PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEMGT_U02
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
                         
                        SetRawData "[결과저장]" & gHOSP.MACHCD & "," & strBarcode & "," & strTestCd & "," & sResult & ",N, null" & "," & strPatID & "," & strSpcCd & "N" & "," & sResult & "," & strCmnt
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_HDINFO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
    
    'strCmnt = "실시간 역전사 종합효소연쇄반응법(Realtime reverse transcriptase PCR)"
        
    strHeader = ""
    strFooter = ""
    strCmnt = ""
    strCmnt = strCmnt & " ▣ Test method         Real-time RT-PCR" & vbCrLf
    strCmnt = strCmnt & " ▣ Reference range     Negative" & vbCrLf
    strCmnt = strCmnt & " ▣ Comment" & vbCrLf
    strCmnt = strCmnt & "    ▶PCR 검사는 검체 내 균수가 적거나 부적절한 검체희석,  증폭억제물질이 " & vbCrLf
    strCmnt = strCmnt & "      존재하는 경우 위음성이 나올수 있습니다." & vbCrLf
    strCmnt = strCmnt & "     또한, PCR 검사는 유전자 유무를 검사하므로 생존균과 사균의 구분이 안되어 " & vbCrLf
    strCmnt = strCmnt & "      위양성의 가능성이 있습니다." & vbCrLf
    strCmnt = strCmnt & "     결과 해석시 환자의 임상 양상을 고려하여 판단하시기 바랍니다." & vbCrLf
    strCmnt = strCmnt & "    " & vbCrLf
    strCmnt = strCmnt & "    " & vbCrLf
    'strCmnt = strCmnt & "검사자:정이진 M.T./보고자:배은신 M.D (tel.3943) "
        
    'strCmnt = strCmnt & "Positive "
    'strCmnt = strCmnt & ""
    'strCmnt = strCmnt & ""
    'strCmnt = strCmnt & " ▣ Test method         Real-time RT-PCR"
    'strCmnt = strCmnt & " ▣ Reference range     Negative"
    'strCmnt = strCmnt & " ▣ Comment"
    'strCmnt = strCmnt & "    ▶PCR 검사는 검체 내 균수가 적거나 부적절한 검체희석,  증폭억제물질이 "
    'strCmnt = strCmnt & "      존재하는 경우 위음성이 나올수 있습니다."
    'strCmnt = strCmnt & "     또한, PCR 검사는 유전자 유무를 검사하므로 생존균과 사균의 구분이 안되어 "
    'strCmnt = strCmnt & "      위양성의 가능성이 있습니다."
    'strCmnt = strCmnt & "     결과 해석시 환자의 임상 양상을 고려하여 판단하시기 바랍니다."
    'strCmnt = strCmnt & "    "
    'strCmnt = strCmnt & "    "
    'strCmnt = strCmnt & "검사자:정이진 M.T./보고자:배은신 M.D (tel.3943) "
        
    ''Positive
    ''
    ''의뢰하신 검체에서 코로나19 (상기도) 양성 반응이 검출되었습니다.
    ''
    ''
    ''코로나19:
    '' [감염병의 예방 및 관리에 관한 법률 시행규칙] 의 개정에 따라 감염병 진단기준에 의거하여 해당 의료기관 관할 보건소에 신고됨을 참고하시기  바랍니다.
    ''
    strPCmnt = ""
    strPCmnt = strPCmnt & " ▣ Test method           Real-time RT-PCR" & vbCrLf
    strPCmnt = strPCmnt & " ▣ Reference range       Negative" & vbCrLf
    strPCmnt = strPCmnt & " ▣ Comment" & vbCrLf
    strPCmnt = strPCmnt & "    ▶PCR 검사는 검체 내 균수가 적거나 부적절한 검체 희석 또는 증폭 억제물질이" & vbCrLf
    strPCmnt = strPCmnt & "      존재하는 경우  위음성이 나올수 있습니다. " & vbCrLf
    strPCmnt = strPCmnt & "     또한, PCR 검사는 유전자 유무를 검사 하므로 생존균과 사균의 구분이 안되어 " & vbCrLf
    strPCmnt = strPCmnt & "      위양성의 가능성이 있습니다." & vbCrLf
    strPCmnt = strPCmnt & "     결과 해석 시, 환자의 임상 양상과 연관지어 판단하시기 바랍니다." & vbCrLf
    strPCmnt = strPCmnt & "    ▶식약처에서 긴급사용 승인제품을 이용한  검사입니다." & vbCrLf
    strPCmnt = strPCmnt & "    ▶코로나19(2019-nCoV)는 현재 질병관리본부에서 잠정적으로 제1급 감염병으로" & vbCrLf
    strPCmnt = strPCmnt & "      적용하는 질환으로 코로나19 유전자검사(PCR)결과로 확진이 되며," & vbCrLf
    strPCmnt = strPCmnt & "     질병관리본부 신고 대상입니다." & vbCrLf
    strPCmnt = strPCmnt & "    " & vbCrLf
    'strPCmnt = strPCmnt & "<Corona CT>"
    'strPCmnt = strPCmnt & "E Gene:"
    'strPCmnt = strPCmnt & "RdRP/S Gene :"
    'strPCmnt = strPCmnt & "N Gene:"
    'strPCmnt = strPCmnt & "    " & vbCrLf
    'strPCmnt = strPCmnt & "    " & vbCrLf
    
    ''검사자:구본상 M.T./보고자:배은신 M.D (tel.3943)
    ''
    strFooter = ""
    strFooter = "검사자:" & frmMain.txtTestNm.Text & " M.T./보고자:최윤미 M.D "
            
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
        
        '-- Local에서 환자별로 결과값 가져오기
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
                
                '-- 장비결과적용
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
                        strHeader = strHeader & "코로나19:"
                        strHeader = strHeader & " [감염병의 예방 및 관리에 관한 법률 시행규칙] 의 개정에 따라 감염병 진단기준에 의거하여 해당 의료기관 관할 보건소에 신고됨을 참고하시기  바랍니다." & vbCrLf
            
                        strCmnt = strHeader & strCmnt & strFooter
                   
                    ElseIf UCase(sResult) = "POSITIVE" Then
                        strHeader = ""
                        strHeader = strHeader & "Positive" & vbCrLf
                        strHeader = strHeader & "" & "의뢰하신 검체에서 코로나19 (상기도) 양성 반응이 검출되었습니다." & vbCrLf
                        strHeader = strHeader & "" & vbCrLf
                        strHeader = strHeader & "코로나19:"
                        strHeader = strHeader & " [감염병의 예방 및 관리에 관한 법률 시행규칙] 의 개정에 따라 감염병 진단기준에 의거하여 해당 의료기관 관할 보건소에 신고됨을 참고하시기  바랍니다." & vbCrLf
                        
                        strPCmnt = strPCmnt & "<Corona CT>" & vbCrLf
                        strPCmnt = strPCmnt & "E Gene:" & strEVal & vbCrLf
                        strPCmnt = strPCmnt & "RdRP/S Gene :" & strRVal & vbCrLf
                        strPCmnt = strPCmnt & "N Gene:" & strNVal & vbCrLf
                        strPCmnt = strPCmnt & "    " & vbCrLf
                        strPCmnt = strPCmnt & "    " & vbCrLf
                        
                        strCmnt = strHeader & strPCmnt & strFooter
                   
                   End If
                                   
                   If sResult <> "" Then
                       '-- 검사결과저장
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
                           
                           If Mid(strTestCd, 1, 2) = "L1" Then '요양병원
                               Set Param03 = .CreateParameter("IN_RESULT", adVarChar, adParamInput, 4000, sResult)
                           Else
                               Set Param03 = .CreateParameter("IN_RESULT", adVarChar, adParamInput, 4000, strCmnt)
                           End If
                           
                           .Parameters.Append Param03
                           
                           Set Param04 = .CreateParameter("IN_USERID", adVarChar, adParamInput, 100, frmMain.txtTestID.Text)
                           .Parameters.Append Param04
                           
                           Set Param05 = .CreateParameter("IN_RFLAG", adVarChar, adParamInput, 50, "D")            '(접수: "C", 입력: "D", 결과: "N")
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
                            
                           If Mid(strTestCd, 1, 2) = "L1" Then '요양병원
                               Call SetSQLData("결과저장", strBarcode & "," & strTestCd & "," & strCmnt & "," & frmMain.txtTestID.Text & ",D" & "," & gHOSP.MACHCD & ","",0," & strErrYN & "," & strErrMsg, "A")
                           Else
                               Call SetSQLData("결과저장", strBarcode & "," & strTestCd & "," & sResult & "," & frmMain.txtTestID.Text & ",D" & "," & gHOSP.MACHCD & ","",0," & strErrYN & "," & strErrMsg, "A")
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
    
    
        If Mid(strTestCd, 1, 2) = "L1" Then '요양병원
            SQL = ""
            SQL = SQL & "UPDATE SLA1COLMT                     " & vbCrLf
            SQL = SQL & "   SET RMK         = '" & strCmnt & "'   " & vbCrLf
            SQL = SQL & " WHERE SPC_NO      = '" & strBarcode & "'" & vbCrLf
            SQL = SQL & "   AND PT_NO       = '" & strPatID & "'  " & vbCrLf
'            SQL = SQL & "   AND EXAM_GRPCD  = '" & strPatID & "'  " & vbCrLf
            SQL = SQL & "   AND ORD_DATE    = '" & Mid(strHospDate, 1, 10) & "'  " & vbCrLf
                                
            Call SetSQLData("결과저장", SQL, "A")
            AdoCn.Execute SQL
        End If
        
    End With
    
    SaveTransData_VHS = 1

Exit Function

ErrHandle:
    SaveTransData_VHS = -1
    
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_VHS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
        
        '-- Local에서 환자별로 결과값 가져오기
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
                
                '-- 장비결과적용
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
                    
                    '서울병원은 바코드가 월일로 시작되고, 목동병원은 바코드가 년월일로 시작된다. 목동바코드는 무조건 13 이상이다!
                    If Len(strBarcode) = 11 And IsNumeric(strBarcode) Then
                        strHospGbn = Mid(strBarcode, 1, 2)
                        If CCur(strHospGbn) > 12 Then
                            gHospCode = "02"      '이대목동병원
                        Else
                            gHospCode = "01"      '이대서울병원
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
'                        sParam = sParam & "<P15><![CDATA[" & vbTab & strCmnt & vbTab & "]]></P15>"   '소견넣는다.
'                    Else
'                        Select Case strEqpCd
'                        Case "BP":  strCmnt = gCmnt.BPNCmnt
'                        Case "CP":  strCmnt = gCmnt.CPNCmnt
'                        Case "LP":  strCmnt = gCmnt.LPNCmnt
'                        Case "MP":  strCmnt = gCmnt.MPNCmnt
'                        End Select
'                        strCmnt = Replace(strCmnt, "*Specimen : ", "*Specimen : " & strSpcmCd)
'                        sParam = sParam & "<P15><![CDATA[" & vbTab & strCmnt & vbTab & "]]></P15>"   '소견넣는다.
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
'                        sParam = sParam & "<P15><![CDATA[" & vbTab & strCmnt & vbTab & "]]></P15>"   '소견넣는다.
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
'                        sParam = sParam & "<P15><![CDATA[" & vbTab & strCmnt & vbTab & "]]></P15>"   '소견넣는다.
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SaveTransData_EHWA" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal
    
End Function

'-- 검사마스터 조회
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
    '검사처방이 없는 결과일경우 처음걸로 가져가게 하기 위해서...
    SQL = SQL & " ORDER BY SEQNO ASC, TESTCODE DESC, TESTNAME "

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        '처방코드, SUB코드용 추가 16,17
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
            gArrEQP(intRow, 16) = ""    '처방코드로 사용
            gArrEQP(intRow, 17) = ""    'SUB코드로 사용

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

    '-- Record Count 가져옴
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


'-- 검사마스터명 조회
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

    '-- Record Count 가져옴
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


'-- 검사마스터 조회
Public Sub GetTestMaster(ByVal SPD As Object)
    Dim intRow          As Long

    SPD.MaxRows = 0
    intRow = 0

    SQL = ""
    SQL = SQL & "SELECT * " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count 가져옴
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


''-- 검사오더마스터 조회
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
''    '-- Record Count 가져옴
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


'-- 검사코드로 검사명 조회
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

    '-- Record Count 가져옴
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
''-- 검사명으로 결과채널 조회
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
'    '-- Record Count 가져옴
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
''-- 검사항목 조회
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
'    '-- Record Count 가져옴
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
'                strErrMsg = "위    치 : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show 'vbModal
'
'    Screen.MousePointer = 0
'
'End Function
'
''-- 워크리스트 조회
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)

    Select Case gEMR
        Case "VHS"
                Call GetWorkList_VHS(pFrom, pTo, SPD)
        
        Case "SWMC"                    '서울의료원
                Call GetWorkList_SWMC(pFrom, pTo, SPD)

'        Case "HDINFO"                       '현대정보
'                Call GetWorkList_HDINFO(pFrom, pTo, SPD)
        
        Case "PHILL"
'                Call GetWorkList_PHILL(pFrom, pTo, SPD)

        Case "MSINFOTEC"                    'MS인포텍
                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)

'        Case "NU"                           '평화IS
'                Call GetWorkList_NU(pFrom, pTo, SPD)

'        Case "AMIS"                         '아미스
'                Call GetWorkList_AMIS(pFrom, pTo, SPD)
'
'        Case "BIGUBCARE"
'                Call GetWorkList_BIGUBCARE(pFrom, pTo, SPD)
'
'        Case "BIT"                          '비트
'                Call GetWorkList_BIT(pFrom, pTo, SPD)
'
'        Case "BIT70"                        '비트 HIB70
'                Call GetWorkList_BIT70(pFrom, pTo, SPD)
'
'        Case "EMEDI"                        '이메디
'                Call GetWorkList_AMIS(pFrom, pTo, SPD)
'
'        Case "EASYS"                        '이지스, MCC
'                Call GetWorkList_EASYS(pFrom, pTo, SPD)

'        Case "EHWA"
'                Call GetWorkList_EHWA(pFrom, pTo, SPD)

'
''        Case "EONM"                         '이온엠
''                Call GetWorkList_EONM(pFrom, pTo, SPD)
'
'        Case "GINUS"                         '지누스
''                Call GetWorkList_GINUS(pFrom, pTo, SPD)
'
'        Case "GSEN"                         '지센커뮤니케이션즈(이챠트)
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "HWASAN"                       '화산
'                Call GetWorkList_HWASAN(pFrom, pTo, SPD)
'
'        Case "JAINCOM"                      '자인컴
'                Call GetWorkList_JAINCOM(pFrom, pTo, SPD)
'
'        Case "JWINFO"                       '중외정보
'                Call GetWorkList_JWINFO(pFrom, pTo, SPD)
'
'        Case "KCHART"                       '다대소프트
'                Call GetWorkList_KCHART(pFrom, pTo, SPD)
'
'        Case "KOMAIN"                       '중외정보
'                Call GetWorkList_KOMAIN(pFrom, pTo, SPD)
'
'        Case "KYU"                          '건양대학교병원 - 워크리스트 기능없음
'                'Call GetWorkList_KYU(pFrom, pTo, SPD)
'
'        Case "MEDICHART"                    '메디챠트
'                Call GetWorkList_MEDICHART(pFrom, pTo, SPD)
'
'        Case "MEDIIT"                       '메디IT(필의료재단)
'                Call GetWorkList_MEDIIT(pFrom, pTo, SPD)
'
'        Case "MEDITOLISS"                   '아름누리
'                Call GetWorkList_MEDITOLISS(pFrom, pTo, SPD)
'
'        Case "MCC"                          'MCC SP버전
'                Call GetWorkList_MCC(pFrom, pTo, SPD)
'
'        Case "MOD"                          'MOD 시스템
'                Call GetWorkList_MOD(pFrom, pTo, SPD)
'
'        Case "MSINFOTEC"                    'MS인포텍
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "NEOSOFT"                      '네오소프트
'                Call GetWorkList_NEOSOFT(pFrom, pTo, SPD)
'
'        Case "ONITGUM"                      '온아티 검진
'                Call GetWorkList_ONITGUM(pFrom, pTo, SPD)
'
'        Case "ONITEMR"                      '온아티 EMR
'                Call GetWorkList_ONITEMR(pFrom, pTo, SPD)
'
'        Case "PLIS"                         '포미스 슈바이처
'                Call GetWorkList_PLIS(pFrom, pTo, SPD)
'
'        Case "SY"                           'SY
'                Call GetWorkList_SY(Format(pFrom, "yyyy-mm-dd"), Format(pTo, "yyyy-mm-dd"), SPD)
'
'        Case "TWIN"                         '투윈정보
'                Call GetWorkList_TWIN(pFrom, pTo, SPD)
'
'        Case "UBCARE"                       '의사랑
'                Call GetWorkList_UBCARE(pFrom, pTo, SPD)

'        Case "WELL"                         '웰커머스
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

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
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
'        frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
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
    SQL = SQL & "   And a.OKFL <> 'Y'                                       " & vbCrLf   '-- 결과확정유무
    'SQL = SQL & "   And a.OKFL = 'N'                                       " & vbCrLf   '-- 결과확정유무
    'SQL = SQL & "   AND (a.RSFL IS NULL OR a.RSFL = 'N' OR a.RSFL = '')     " & vbCrLf
    SQL = SQL & " GROUP BY a.ORDT,a.SPNO,a.PAID,a.NWNO,b.PANM,b.SEXS,b.AGES " & vbCrLf
    SQL = SQL & " Order By a.ORDT,a.PAID,b.PANM                             " & vbCrLf

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
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
'            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
    
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
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
'            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_SWMC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
    '2020-07-30 수정 : 처방일자가 아닌 접수일자 로 가져와야 함
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
    '2021-05-13 L5300T 추가
    '2021-08-04 L5300U1 추가
    SQL = SQL & "   and A.EXAM_CD IN ('L5300T','L5300I','L5300J','L5300M','L9084','L9085','L5300U1')" & vbCrLf
    SQL = SQL & "   AND A.RPT_YN <> 'Y'" & vbCrLf
    SQL = SQL & "   AND B.EXAM_PROSS_STS IN ('C', 'D', 'M')" & vbCrLf
    SQL = SQL & " Union All " & vbCrLf
    SQL = SQL & "SELECT DISTINCT " & vbCrLf
    SQL = SQL & "       2                       AS SORT " & vbCrLf
    '2020-07-30 수정 : 처방일자가 아닌 접수일자 로 가져와야 함
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
    '2020-07-30 수정 : 처방일자가 아닌 접수일자 로 가져와야 함
    
    SQL = SQL & "   and A.ORD_DATE BETWEEN  '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    
    
    
    'SQL = SQL & "   and A.EXAM_CD IN (" & gAllTestCd & ")" & vbCrLf
    '2021-06-15 추가 ,'L1204' 추가
    SQL = SQL & "   and A.EXAM_CD IN ('L1202','L1203','L1204')" & vbCrLf
    SQL = SQL & "   AND A.RPT_YN <> 'Y'"
    SQL = SQL & "   AND B.EXAM_PROSS_STS IN ('C', 'D', 'M')"

'    SQL = SQL & " GROUP BY PART_JUBSU_DATE, SPECIMEN_SER, BUNHO, FKCPL0201, SPECIMEN_CODE, SUNAME, AGE, SEX, GWA   " & vbCrLf
    SQL = SQL & " ORDER BY SORT, HOSPDATE, CHARTNO"
    
    
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
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
'            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_VHS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
                
                '환자 성별/나이
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
                
                '오더갯수
'                SetText SPD, CStr(intTestCnt), asRow, colOCNT

                '오더정보에 저장
                With mOrder
                    .BarNo = strBarcode
                    .PID = gPatInfo_Select.PT_NO
                    .PNAME = gPatInfo_Select.PT_NM
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With

                '-- 화면에 표시
                'If Trim(varRcvData(10) & "") <> "" Then
'                    For intCol = colSTATE + 1 To .MaxCols
'                        If gExam_Select(i).TST_CD = gArrEQP(intCol - colSTATE, 2) Then
'                            .Row = asRow
'                            .Col = intCol
'                            .BackColor = vbYellow
'                            Call SetText(SPD, "◇", asRow, intCol)
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
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_NU" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
    
    
'''    sParam = ""
'''    sParam = sParam & "submit_id=TRLII00101&"                                   'submit ID
'''    sParam = sParam & "business_id=li&"                                         'business_id
'''    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '기관코드
'''
'''    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'''
'''    Call SetSQLData("워크조회", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
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
'''                    strTestCds = Replace(strTestCds, "▦", "")
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
''''            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
''''    'strTestCds = "LIM305▦LIM306▦"
''''    'strTestCds = "LIM305"
''''
''''
''''    If optState(0).Value = True Then
''''        'sParam = "submit_id=TRLII00119&"                                           'submit ID
''''        sParam = "submit_id=TRLII00101&"                                            'submit ID
''''        sParam = sParam & "business_id=lis&"                                        'business_id
''''        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '사용자ID|기관코드
''''        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                            '기관코드
''''        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                            '장비코드
''''        sParam = sParam & "startdd=" & sSch1 & "&"                                  '시작작업일자
''''        sParam = sParam & "enddd=" & sSch2 & "&"                                    '종료작업일자
''''    Else
''''        sParam = "submit_id=TRLQI00101&"                                            'submit ID
''''        sParam = sParam & "business_id=lis&"                                        'business_id
''''        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '사용자ID|기관코드
''''        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                            '기관코드
''''        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                            '장비코드
''''        sParam = sParam & "startdd=" & sSch1 & "&"                                  '시작작업일자
''''        sParam = sParam & "enddd=" & sSch2 & "&"                                    '종료작업일자
''''    End If
''''
''''    '==> 서버로 오더조회
''''    'SetRawData "[WL_IN]" & sParam
''''        'spcacptdt 접수일자
''''        'acptflag 입원외래구분
''''        'bcno 검체번호
''''        'PID 등록번호
''''        'patnm 환자명
''''        'sexage 나이성별
''''        'erprcpflag 응급구분
''''        'workno 작업번호
''''        'tsectnm 검사계명
''''        'ifreqcdlist 장비요청코드
''''        'tclscdlist 검사리스트
''''        'urinextrvol 유린값
''''        'retestyn 재검여부
''''        'rsltstat 결과상태
''''    sRcvData = OpenURLWithIE2(NUAPI.APIURL & sParam, Inet1)
''''
''''    Call SetSQLData("워크조회", NUAPI.APIURL & sParam & vbNewLine & sRcvData)
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
''''                strTestCds = Replace(strTestCds, "▦", "")
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
''''    MsgBox "조회 오류", vbOKOnly + vbCritical, Me.Caption
'On Error GoTo ErrHandle
'
'    Screen.MousePointer = 11
'    blnSame = False
'    strTestCds = ""
'
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00101&"                                   'submit ID
'    sParam = sParam & "business_id=li&"                                         'business_id
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '기관코드
'
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("워크조회", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
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
'                    strTestCds = Replace(strTestCds, "▦", "")
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
''            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
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
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                        '기관코드
'    sParam = sParam & "startdd=" & pFrom & "&"                              '시작작업일자
'    sParam = sParam & "enddd=" & pTo & "&"                                  '종료작업일자
'
'    If gTest = "MTB" Then
'        'sParam = sParam & "testcd=C6021C&"                                  '검사코드
'        sParam = sParam & "testcd=" & gComm.MTBORD & "&"                                  '검사코드
'
'    ElseIf gTest = "RP" Then
'        If l = 0 Then
''            sParam = sParam & "testcd=VB8506B&"                             '검사코드
'            sParam = sParam & "testcd=" & gComm.RP19ORD_1 & "&"                                   '검사코드
'        Else
''            sParam = sParam & "testcd=D6802060&"                            '검사코드
'            sParam = sParam & "testcd=" & gComm.RP19ORD_2 & "&"                                   '검사코드
'        End If
'
'    ElseIf gTest = "PB" Then
''        sParam = sParam & "testcd=D6801&"                                   '검사코드
'        sParam = sParam & "testcd=" & gComm.PB6ORD & "&"                                    '검사코드
'
'    ElseIf gTest = "RPPB" Then
'        If l = 0 Then
''            sParam = sParam & "testcd=VB8506B&"                             '검사코드    RP19종
'            sParam = sParam & "testcd=" & gComm.RP19ORD_1 & "&"                                   '검사코드
'        ElseIf l = 1 Then
''            sParam = sParam & "testcd=D6802060&"                            '검사코드    RP19종
'            sParam = sParam & "testcd=" & gComm.RP19ORD_2 & "&"                                   '검사코드
'        Else
''            sParam = sParam & "testcd=D6801&"                               '검사코드    PB 6종
'            sParam = sParam & "testcd=" & gComm.PB6ORD & "&"                                    '검사코드
'        End If
'    End If
'
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("워크조회", "Param:" & gHOSP.APIURL & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
'
'    '<worklist>
'        '<bcno><![CDATA[3041900020]]></bcno>
'        '<patnm><![CDATA[이명숙]]></patnm>
'        '<prgstno><![CDATA[600603-2******]]></prgstno>
'        '<pid><![CDATA[000137388]]></pid>
'        '<sex><![CDATA[F]]></sex>
'        '<age><![CDATA[59]]></age>
'        '<spcnm><![CDATA[Throat swab]]></spcnm>
'        '<spccd><![CDATA[023]]></spccd>
'        '<tclscd><![CDATA[VB6012A]]></tclscd>
'        '<spcstat><![CDATA[4]]></spcstat>
'        '<rsltstat><![CDATA[4]]></rsltstat>                         '결과상태
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
'                '2019-12-11 메모
'                '   rsltstat 로 처리 가능 할 둣...
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
'        MsgBox "조회 대상자가 없습니다.", vbOKOnly + vbCritical, "워크리스트 조회"
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
'                strErrMsg = "위    치 : " & gHOSP.MACHNM & "_GetWorkList_HDINFO" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show 'vbModal
'
'    Screen.MousePointer = 0
'
'End Sub

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
            'SQL = SQL & "   AND O.H141_NOTYYN = 'N'                         " & vbCr
            SQL = SQL & "   AND O.H141_NOTYYN       IN ('N','T')                 " & vbCr '결과대기:T
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
            SQL = SQL & "Select DISTINCT ORCD AS ITEM       " & vbCrLf
            SQL = SQL & "  From emr.LRESULT                 " & vbCrLf
            SQL = SQL & " Where PAID = '" & strPatID & "'   " & vbCrLf
            SQL = SQL & "   And SPNO =  '" & strBarcode & "'" & vbCrLf
            SQL = SQL & "   And ORCD IN (" & gAllTestCd & ")" & vbCrLf
            SQL = SQL & "   And OKFL <> 'Y'                 " & vbCrLf   '-- 결과확정유무
            SQL = SQL & " Order By ORCD                     " & vbCrLf
        
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


Public Sub LetEqpMaster(ByVal pEqpCD As String)

    SQL = ""
    SQL = SQL & "UPDATE EQPMASTER SET EQUIPCD = '" & pEqpCD & "'"

    Call DBExec(AdoCn_Local, SQL)

End Sub

'-- 장비결과 조회
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
    '-- 검사결과
    SQL = SQL & ",SEQNO,EXAMNAME,RESULT,PREVRESULT,REFJUDGE" & vbCr

    SQL = SQL & "  FROM PATRESULT " & vbCr
    '-- 검사일자
    SQL = SQL & " WHERE EXAMDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
'    SQL = SQL & "   AND EXAMCODE IN (" & gAllTestCd & ") " & vbCr
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,BARCODE,SEQNO"

    '-- Record Count 가져옴
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
                            SetText SPD, "장비결과", intRow, colSTATE
                    Case "1"
                            SetText SPD, "저장에러", intRow, colSTATE
                    Case "2"
                            SetText SPD, "전송완료", intRow, colSTATE
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
        'frmMain.lblStatus.Caption = ">> 조회 대상자가 없습니다."
        'frmMain.chkRAll.Value = "0"
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

'    Call frmMain.GetPatTRestResult_Search(1)

    Screen.MousePointer = 0

End Sub

''-- 검사자 정보 가져오기
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
'        Case "KYU"                  '건양대학교병원
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
'        Case "MEDITOLISS"                   '아름누리
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
'        Case "ONITGUM"                      '온아티 검진
'                Call GetSampleInfo_ONITGUM(asRow, SPD)
'
'        Case "ONITEMR"                      '온아티 EMR
'                Call GetSampleInfo_ONITEMR(asRow, SPD)
'
        Case "PHILL"
                Call GetSampleInfo_PHILL(asRow, SPD)
                
'        Case "NU"
'                Call GetSampleInfo_NU(asRow, SPD)
                
'        Case "PLIS"                      '온아티 EMR
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

'-- 검사자 정보 가져오기
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
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
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
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_PHILL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
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
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
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
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If GetTestNm(Trim(RS.Fields("ITEM"))) = gArrEQPNm(intCol - colSTATE, 6) Then
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
    
    GetSampleInfo_SWMC = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_SWMC = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_SWMC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_VHS(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    '-- 바코드 번호로 오더 조회
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
    
    Call SetSQLData("바코드조회", "PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEEQP_S01 >> " & strBarcode)
    
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
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .PID = Trim(RS.Fields("PT_NO")) & ""
                    .PNAME = Trim(RS.Fields("PT_NAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If GetTestNm(Trim(RS.Fields("EXAM_CD"))) = gArrEQPNm(intCol - colSTATE, 6) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
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
        
        'CT값 처리용
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
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_VHS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- 검사자 정보 가져오기
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
'    sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"     '사용자ID|기관코드
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '기관코드
'    sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '장비코드
'    sParam = sParam & "bcno=" & strBarcode                                          '바코드
'
''    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("바코드조회", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
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
'                '환자 성별/나이
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
'                '오더갯수
'                SetText SPD, CStr(intTestCnt), asRow, colOCNT
'
'                '오더정보에 저장
'                With mOrder
'                    .BarNo = varRcvData(3) & ""
'                    .PID = varRcvData(4) & ""
'                    .PNAME = varRcvData(5) & ""
'                    .Count = CStr(intTestCnt)
'                    .NoOrder = False
'                End With
'
'                '-- 화면에 표시
'                If Trim(varRcvData(10) & "") <> "" Then
'                    varTstCode = Split(varRcvData(11), "▦")
'                    For j = 0 To UBound(varTstCode) - 1
'                        gPatOrdCd = gPatOrdCd & "'" & Trim(varTstCode(j)) & "',"
'
'                        For intCol = colSTATE + 1 To .MaxCols
'                            If Trim(varTstCode(j)) = gArrEQP(intCol - colSTATE, 2) Then
'                                .Row = asRow
'                                .Col = intCol
'                                .BackColor = vbYellow
'                                Call SetText(SPD, "◇", asRow, intCol)
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
''    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_NU" & vbNewLine & vbNewLine
''    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
''    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
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
    'xmlDoc.Load "D:\프로젝트\VB\광주포유병리과의원\참고\Info.xml"
    
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
        '<patnm><![CDATA[박성일]]></patnm>
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
                            Case "AGE":             XmlSelectS.AGE(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '나이           [78]
                            Case "BCNO":            XmlSelectS.BCNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                '검체번호       [3010700030]
                            Case "EXECprcpuniqno":  XmlSelectS.EXECprcpuniqno(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue      '챠트번호?      [2002638354]
                            Case "ORDDEPTCD":       XmlSelectS.ORDDEPTCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '처방부서코드?  [NU]
                            
                            Case "PATNM":           XmlSelectS.PATNM(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '환자명         [박성일]
                            Case "PID":             XmlSelectS.PID(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '환자번호       [000132623]
                            Case "PRCPDD":          XmlSelectS.PRCPDD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '처방일?        [20181217]
                            Case "PRGSTNO":         XmlSelectS.PRGSTNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '주민번호       [400321-1******]
                            
                            Case "RETESTYN":        XmlSelectS.RETESTYN(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue            '재검여부       [N]
                            Case "RSLTSTAT":        XmlSelectS.RSLTSTAT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue            '결과상태       [-]
                            
                            Case "SEX":             XmlSelectS.SEX(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '성별           [M]
                            
                            Case "SPCACPTDT":       XmlSelectS.SPCACPTDT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '검체채취시간?  [20181217094414]
                            Case "SPCCD":           XmlSelectS.SPCCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '검체코드       [023]
                            Case "SPCNM":           XmlSelectS.SPCNM(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '검체명         [Throat swab]
                            Case "SPCSTAT":         XmlSelectS.SPCSTAT(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '검체상태       [4]
                            
                            Case "TCLSCD":          XmlSelectS.TCLSCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '처방코드       [VB6012A]
                            Case "TESTCD":          XmlSelectS.TESTCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '검사코드       [VB6012A]
                            Case "TESTLRGCD":       XmlSelectS.TESTLRGCD(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '결과그룹코드?  [I]
                            Case "WORKNO":          XmlSelectS.WORKNO(j) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '워크번호       [20181217I20002]
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

'-- 검사자 정보 가져오기
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
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '기관코드
'    sParam = sParam & "bcno=" & strBarcode                                      '검체번호(=바코드)
'
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("바코드조회", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
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
'                '환자 성별/나이
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
'                            Call SetText(SPD, "◇", asRow, intCol)
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
'        'MsgBox "조회 대상자가 없습니다.", vbOKOnly + vbCritical, "워크리스트 조회"
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
''    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_NU" & vbNewLine & vbNewLine
''    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
''    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
''    frmErrMsg.txtErr = vbNewLine & strErrMsg
''    frmErrMsg.Show
'
'End Function

'-- 검사자 정보 가져오기
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
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '기관코드
'    sParam = sParam & "bcno=" & asBarcode                                      '검체번호(=바코드)
'
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("바코드테스트조회", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
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

'-- 검사자 정보 가져오기
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
    
    '서울병원은 바코드가 월일로 시작되고, 목동병원은 바코드가 년월일로 시작된다. 목동바코드는 무조건 13 이상이다!
    If Len(strBarcode) = 11 And IsNumeric(strBarcode) Then
        strHospGbn = Mid(strBarcode, 1, 2)
        If CCur(strHospGbn) > 12 Then
            gHospCode = "02"      '이대목동병원
        Else
            gHospCode = "01"      '이대서울병원
        End If
    End If
    
    Screen.MousePointer = 11
  
    sRes = Online_XML(gXml_ORDER_SELECT, strBarcode, "GETQUERY", "", "") ' "PKG_MSE_LM_INTERFACE.PC_MSE_ORDER_SELECT"
  
'    sRes = Online_XML(gXml_LOGIN, "", "GETQUERY", txtID.Text, txtPW.Text) ' "PKG_MSE_LM_INTERFACE.PC_MSE_ORDER_SELECT"
  
  
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00101&"                                       'submit ID
'    sParam = sParam & "business_id=li&"                                             'business_id
'    sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"     '사용자ID|기관코드
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '기관코드
'    sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '장비코드
'    sParam = sParam & "bcno=" & strBarcode                                          '바코드
        
'    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'    Call SetSQLData("바코드조회", "Param:" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)
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
                
                '환자 성별/나이
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
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT

                '오더정보에 저장
                With mOrder
                    .BarNo = strBarcode
                    .PID = gPatInfo_Select.PT_NO
                    .PNAME = gPatInfo_Select.PT_NM
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With

                '-- 화면에 표시
                'If Trim(varRcvData(10) & "") <> "" Then
                    For intCol = colSTATE + 1 To .MaxCols
                        If gExam_Select(i).TST_CD = gArrEQP(intCol - colSTATE, 2) Then
                            .Row = asRow
                            .Col = intCol
                            .BackColor = vbYellow
                            'Call SetText(SPD, "◇", asRow, intCol)
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
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_NU" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
    
End Function

Function SetJudge(asResult As String, asEquipCode As String)

    Select Case gEMR
        Case "AMIS"                         '아미스
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EMEDI"                        '이메디
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "BIT"                          '비트
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)

        Case "BIT70"                        '비트 HIB70
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EASYS"                        '이지스
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EONM"                         '이온엠
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "GSEN"                         '지센커뮤니케이션즈(이챠트)
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "HWASAN"                       '화산
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JAINCOM"                       '자인컴
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JWINFO"                       '중외정보
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KCHART"                       '다대소프트
                SetJudge = SetJudge_KCHART(asResult, asEquipCode)
        
        Case "KOMAIN"                       '중외정보
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KYU"                          '건양대학교병원
                '워크리스트 기능없음
                'SetJudge =  SetJudge_KYU(asResult,asEquipCode)
        Case "MEDICHART"                    '메디챠트
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDIIT"
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDITOLISS"                    '
                SetJudge = SetJudge_MEDITOLISS(asResult, asEquipCode)
            
        Case "MSINFOTEC"                    'MS인포텍
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
    
'    SQL = SQL & "  L.진료검사ID AS R, " & vbCrLf
'    SQL = SQL & "  L.진료지원ID AS P, " & vbCrLf

    '성인남 참고치0~참고치1,
    '성인여 참고치2~참고치3,
    '소아남 참고치4~참고치5,
    '소아여 참고치6~참고치7
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       A.환자성별 AS 성별                                          " & vbCr
    SQL = SQL & "     , L.참고치0, L.참고치1, L.참고치2, L.참고치3                  " & vbCr
    SQL = SQL & "     , L.참고치4, L.참고치5, L.참고치6, L.참고치7                  " & vbCr
    SQL = SQL & "     , (L.처방코드 + L.서브코드) AS ITEM                           " & vbCr
    SQL = SQL & "  FROM             TB_진료검사 L                                   " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료지원 J ON (L.진료지원ID = J.진료지원ID)  " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료일반 A ON (J.진료일자   = A.진료일자     " & vbCr
    SQL = SQL & "                                AND  J.챠트번호   = A.챠트번호     " & vbCr
    SQL = SQL & "                                AND  J.진료번호   = A.진료번호)    " & vbCr
    SQL = SQL & "  Where L.검체번호 = '" & mResult.BarNo & "'                       " & vbCr
    SQL = SQL & "    AND L.검사상태 < 5                                             " & vbCr
    SQL = SQL & "    AND (L.처방코드 + L.서브코드) = '" & sEquipCode & "'           " & vbCr
                                                                 

     Call SetSQLData("참고치조회", SQL)
     
     '-- Record Count 가져옴
     AdoCn.CursorLocation = adUseClient
     Set RS1 = AdoCn.Execute(SQL, , 1)
     If Not RS1.EOF = True And Not RS1.BOF = True Then
         Do Until RS1.EOF
            strRefL = ""
            strRefH = ""
            If Trim(RS1.Fields("성별")) & "" = "M" Then
                If Trim(RS1.Fields("참고치0")) & "" <> "" Then
                    strRefL = Trim(RS1.Fields("참고치0")) & ""
                    strRefH = Trim(RS1.Fields("참고치1")) & ""
                End If
            Else
                If Trim(RS1.Fields("성별")) & "" = "F" Then
                    If Trim(RS1.Fields("참고치2")) & "" <> "" Then
                        strRefL = Trim(RS1.Fields("참고치2")) & ""
                        strRefH = Trim(RS1.Fields("참고치3")) & ""
                    Else
                        strRefL = Trim(RS1.Fields("참고치0")) & ""
                        strRefH = Trim(RS1.Fields("참고치1")) & ""
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
            Case "변함없음"
                    sResult = Trim(asResult)
            Case "정량"
                    sResult = Trim(asResult)
            Case "정성"
                    sResult = Trim(sResult)
            Case "정량(정성)"
                    sResult = asResult & "(" & Trim(sResult) & ")"
            Case "정성(정량)"
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
            SQL = SQL & "  EQUIPNO"                         '장비코드
            SQL = SQL & ", EXAMDATE"                        '검사일자
            SQL = SQL & ", EXAMTIME"                        '검사시간
            SQL = SQL & ", SAVESEQ"                         '저장순번(날짜별)
            SQL = SQL & ", HOSPDATE" & vbCrLf               '병원접수일자
            
            SQL = SQL & ", BARCODE"                         '검체번호
            SQL = SQL & ", PID"                             '병록번호(내원번호)
            SQL = SQL & ", CHARTNO"                         '챠트번호
            SQL = SQL & ", SPECIMEN"                        '검체
            SQL = SQL & ", DEPT" & vbCrLf                   '의뢰과
            
            SQL = SQL & ", INOUT"                           '입/외
            SQL = SQL & ", ERYN"                            '응급여부
            SQL = SQL & ", RETESTYN"                        '재검여부
            SQL = SQL & ", PNAME"                           '이름
            SQL = SQL & ", PSEX" & vbCrLf                   '성별(M,F)
            
            SQL = SQL & ", PAGE"                            '나이
            SQL = SQL & ", EXAMUID"                         '검사자ID
            SQL = SQL & ", DISKNO"                          'Rack
            SQL = SQL & ", POSNO"                           'Pos
            SQL = SQL & ", EQPSEQNO" & vbCrLf               '장비검사번호
            '============================================================
            
            SQL = SQL & ", SEQNO"                           '검사일련번호
            SQL = SQL & ", EQUIPCODE"                       '검사채널
            SQL = SQL & ", ORDERCODE"                       '병원처방코드
            SQL = SQL & ", EXAMCODE"                        '병원검사코드
            SQL = SQL & ", EXAMCODESUB" & vbCrLf            '병원검사코드(SUB)"
            
            SQL = SQL & ", EXAMNAME"                        '검사명
            SQL = SQL & ", EQUIPRESULT"                     '장비결과"
            SQL = SQL & ", RESULT"                          '소수점적용"
            SQL = SQL & ", PREVRESULT"                      '이전결과"
            SQL = SQL & ", REFJUDGE" & vbCrLf               '판정(H,L)
            
            SQL = SQL & ", REFFLAG"                         'flag
            SQL = SQL & ", REFVALUE"                        '참고치
            SQL = SQL & ", PANICVALUE"                      'Delta
            SQL = SQL & ", DELTAVALUE"                      'Panic
            SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
            SQL = SQL & ", SENDDATE)" & vbCrLf               '전송일자
            
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
            
            'SetRawData "[로컬저장]" & SQL
            
            If Not DBExec(AdoCn_Local, SQL) Then
                Exit Function
            End If

        End If
    End With

End Function

'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
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
