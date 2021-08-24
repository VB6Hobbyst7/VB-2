Attribute VB_Name = "modTwCtrl"
Option Explicit

Public Sub startTWView()
    
    '검색창에 시간 정보 셋팅
    'setTwSechCondition
    
    mainFrm.chkTimer_Tw = 1
    '검색
    'goTwSearch
End Sub

Public Sub setTwStationID()

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    strQry = "SELECT STATION_NAME, STATION_ID " + vbCrLf
    strQry = strQry + "FROM WRN.T_WRN_STATION " + vbCrLf
    strQry = strQry + "WHERE STATION_ID IN (SELECT DISTINCT STATION_ID FROM WRN.T_WRN_TW_BUOY WHERE TO_CHAR(OBS_TIME, 'YYYY') > '2011') " + vbCrLf
    strQry = strQry + "  AND STATION_ID > ' ' " + vbCrLf
    strQry = strQry + "ORDER BY STATION_NAME " + vbCrLf

    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    mainFrm.cmbxSechTWID.AddItem "전체"
    'mainFrm.CboTw_NM.AddItem "전체"

    If Not (dtRs.EOF Or dtRs.BOF) Then

        Dim cnt As Integer

        For cnt = 0 To dtRs.RecordCount - 1
            mainFrm.cmbxSechTWID.AddItem dtRs.Fields("STATION_NAME")
            mainFrm.CboTw_NM.AddItem dtRs.Fields("STATION_NAME")
            mainFrm.CboTw_ID.AddItem dtRs.Fields("STATION_ID")
            
            dtRs.MoveNext
        Next cnt
    
        mainFrm.cmbxSechTWID.ListIndex = 0
        mainFrm.CboTw_NM.ListIndex = 0
        mainFrm.CboTw_ID.ListIndex = mainFrm.CboTw_NM.ListIndex
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub setTwSechCondition()

    Dim cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '검색 일자 설정
    mainFrm.txtSechTWStDate.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtSechTWEdDate.Text = Format(Now, "YYYY-MM-DD")
    
    
End Sub

'Public Sub chkTwSechCondition()
'
'    If Not IsDate(mainFrm.txtSechTWStDate.Text) Then
'        MsgBox "검색 범위 시작일자를 확인해주세요."
'        mainFrm.txtSechTWStDate.SetFocus
'        Exit Sub
'    End If
'
'    If Not IsDate(mainFrm.txtSechTWEdDate.Text) Then
'        MsgBox "검색 범위 종료일시를 확인해주세요."
'        mainFrm.txtSechTWEdDate.SetFocus
'        Exit Sub
'    End If
'
'    If mainFrm.cmbxTwRownum.Text = "ALL" Then
'    ElseIf Not IsNumeric(mainFrm.cmbxTwRownum.Text) Then
'        MsgBox "출력건수는 숫자를 입력해주세요."
'        mainFrm.cmbxTwRownum.SetFocus
'        Exit Sub
'    End If
'End Sub

Public Sub goTwSearch()
    '검색조건 체크 st
    If Not IsDate(mainFrm.txtSechTWStDate.Text) Then
        MsgBox "검색 범위 시작일자를 확인해주세요."
        mainFrm.txtSechTWStDate.SetFocus
        Exit Sub
    End If
    If Not IsDate(mainFrm.txtSechTWEdDate.Text) Then
        MsgBox "검색 범위 종료일시를 확인해주세요."
        mainFrm.txtSechTWEdDate.SetFocus
        Exit Sub
    End If
    If mainFrm.cmbxTwRownum.Text = "ALL" Then
    ElseIf Not IsNumeric(mainFrm.cmbxTwRownum.Text) Then
        MsgBox "출력건수는 숫자를 입력해주세요."
        mainFrm.cmbxTwRownum.SetFocus
        Exit Sub
    End If
    '검색조건 체크 end
    
    '검색초기화
    Init_fpSpread_TwLog

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    Dim stDate As String
    Dim edDate As String
    
    'DB접속
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn
    
    stDate = mainFrm.txtSechTWStDate.Text
    edDate = mainFrm.txtSechTWEdDate.Text
    
    strQry = ""
    strQry = strQry + "SELECT * " + vbCrLf
    strQry = strQry + "FROM ( " + vbCrLf
    strQry = strQry + "SELECT /*+ INDEX(A, IDX_LOG_TW_BUOY_01) */  CASE WHEN STATION_ID='0' Then '' ELSE STATION_ID End STATION_ID  " + vbCrLf
    strQry = strQry + "       ,CASE WHEN STATION_ID='0' THEN '공통' " + vbCrLf
    strQry = strQry + "       ELSE (SELECT STATION_NAME FROM WRN.T_WRN_STATION WHERE STATION_ID = A.STATION_ID) " + vbCrLf
    strQry = strQry + "       END STATION_NAME " + vbCrLf
    strQry = strQry + "       , CASE WHEN STATION_ID<>'0' Then TO_CHAR(OBS_TIME,'yyyy/mm/dd hh24:mi:ss') End OBS_TIME " + vbCrLf
    strQry = strQry + "       , TO_CHAR(REG_DATE,'yyyy/mm/dd hh24:mi:ss') REG_DATE, LOG_CONTENT " + vbCrLf
    strQry = strQry + "FROM WRN.LOG_TW_BUOY A " + vbCrLf
    strQry = strQry + "WHERE REG_DATE BETWEEN TO_DATE('" + stDate + "'||'000000', 'YYYY-MM-DDHH24MISS') AND TO_DATE('" + edDate + "'||'232359', 'YYYY-MM-DDHH24MISS')  " + vbCrLf
    If Not mainFrm.cmbxSechTWID.Text = "전체" Then
        strQry = strQry + "  AND STATION_ID IN ('0', (SELECT STATION_ID FROM WRN.T_WRN_STATION WHERE STATION_NAME = '" + mainFrm.cmbxSechTWID.Text + "')) " + vbCrLf
    End If
    strQry = strQry + ") " + vbCrLf
    
    If Not mainFrm.cmbxTwRownum.Text = "ALL" Then
    strQry = strQry + "WHERE ROWNUM <= " + mainFrm.cmbxTwRownum.Text + vbCrLf
    End If
    
    
    
'LogWrite strQry

    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
'        If AdoDBConn.Errors.Count = 0 Then
'            mainFrm.fpSpread_TwLog.DataSource = dtRs.DataSource
'        Else
'            AdoDBConn.Errors.Clear
'        End If
        Dim cnt As Integer
        
        mainFrm.fpSpread_TwLog.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_TwLog.MaxRows = dtRs.RecordCount
        
    
        For cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_TwLog.SetText(j + 1, cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next cnt
        
        
        '색상처리
'        For cnt = 0 To mainFrm.fpSpread_TwLog.MaxRows - 1
'            mainFrm.fpSpread_TwLog.Col = 2
'            mainFrm.fpSpread_TwLog.Row = cnt + 1
'
'            If DateDiff("d", CDate(mainFrm.fpSpread_TwLog.Text), Now) >= 1 Then
'                '-- 셀의 배경색상 변경
'                For j = 0 To mainFrm.fpSpread_TwLog.MaxCols - 1
'                    mainFrm.fpSpread_TwLog.Col = j + 1
'                    mainFrm.fpSpread_TwLog.Row = cnt + 1
'                    mainFrm.fpSpread_TwLog.BackColor = vbRed
'                    mainFrm.fpSpread_TwLog.ForeColor = vbWhite
'                Next j
'
'                mainFrm.fpSpread_TwLog.Col = 2
'                mainFrm.fpSpread_TwLog.Row = cnt + 1
'            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_TwLog.Text), Now) >= strJowiVPNCautionMin Then
'                    '-- 셀의 배경색상 변경
'                    For j = 0 To mainFrm.fpSpread_TwLog.MaxCols - 1
'                        mainFrm.fpSpread_TwLog.Col = j + 1
'                        mainFrm.fpSpread_TwLog.Row = cnt + 1
'                        mainFrm.fpSpread_TwLog.BackColor = vbYellow
'                        mainFrm.fpSpread_TwLog.ForeColor = vbBlack
'                    Next j
'
'                    mainFrm.fpSpread_TwLog.Col = 2
'                    mainFrm.fpSpread_TwLog.Row = cnt + 1
'                    'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
'            Else
'                For j = 0 To mainFrm.fpSpread_TwLog.MaxCols - 1
'                    mainFrm.fpSpread_TwLog.Col = j + 1
'                    mainFrm.fpSpread_TwLog.Row = cnt + 1
'                    mainFrm.fpSpread_TwLog.BackColor = vbWhite
'                    mainFrm.fpSpread_TwLog.ForeColor = vbBlack
'                Next j
'
'            End If
'        Next cnt
        
    
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    
    'DB접속종료
    If AdoDBConn.State = adStateOpen Then
       AdoDBConn.Close
    End If
    
    If Not AdoDBConn Is Nothing Then
        Set AdoDBConn = Nothing
    End If
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub Init_fpSpread_TwLog()
    
    With mainFrm.fpSpread_TwLog
    
        .Reset
        
        .OperationMode = OperationModeRow
        .GridSolid = False
        
        .Appearance = Appearance3D
                
        'Hide row header
        .RowHeadersShow = False
        
        'Turn off font bold
        .Col = -1
        .Row = -1
        .FontBold = False
        
        'Change the amount of data each cell will hold
        .Col = -1
        .Row = -1
        .TypeEditLen = 200
        
        'Set column display type
        .ColHeaderDisplay = DispBlank
        .AllowCellOverflow = True
        .ReDraw = True
        
        .ShowScrollTips = ShowScrollTipsVertical
        .GrayAreaBackColor = &HFFFFFF
        
        .TextTip = TextTipFloating
        
        .MaxCols = 5
        .MaxRows = 0
        
'        .Col = .MaxCols - 1
'        .ColHidden = True
'
'        .Col = .MaxCols
'        .ColHidden = True
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "관측소ID"
        .ColWidth(1) = 10
        .SetText 2, 0, "관측소명"
        .ColWidth(2) = 10
        .SetText 3, 0, "관측시간"
        .ColWidth(3) = 15
        .SetText 4, 0, "로그기록시간"
        .ColWidth(4) = 15
        .SetText 5, 0, "로그내용"
        .ColWidth(5) = 45
    End With
    
End Sub



