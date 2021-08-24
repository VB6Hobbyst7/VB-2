Attribute VB_Name = "modTotCtrl"
Option Explicit

Public dTimes As Double
Public strAdoDBConn As String
Public AdoDBConn As ADODB.Connection
Public strQry As String
Public strJowiVPNCautionMin As Integer
Public strJowiCDMACautionMin As Integer
Public strTwCautionMin As Integer
Public strAgCautionMin As Integer
Public strRtCautionMin As Integer
Public strUsnCautionMin As Integer

Type DatabaseInfo
    ID As String
    PW As String
    DataSource As String
End Type

Public CfgDb As DatabaseInfo
Public CfgTw As DatabaseInfo
Public CfgAg As DatabaseInfo
Public CfgRt As DatabaseInfo

'-- DataDiff
'Year:  yy , yyyy
'Quarter:  qq , q
'Month:  mm , m
'DayofYear:  dy , y
'Day:  dd , d
'Week:  wk , ww
'Weekday:  dw , w
'Hour:  Hh
'Minute:  mi , n
'Second:  ss , s
'Millisecond:  Ms

Public Sub Init_fpSpread_Tot_DtVPN()
    
    With mainFrm.fpSpread_Tot_DtVPN
    
        '.Reset
        .MaxRows = 0
        .MaxRows = 500
        
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
        
        .MaxCols = 2
        .MaxRows = 0
        
'        .Col = .MaxCols - 1
'        .ColHidden = True
'
'        .Col = .MaxCols
'        .ColHidden = True
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "관측소"
        '.ColWidth(1) = 7
        .SetText 2, 0, "관측시간"
        '.ColWidth(2) = 15
    End With
    
End Sub

Public Sub Init_fpSpread_Tot_DtCDMA()
    
    With mainFrm.fpSpread_Tot_DtCDMA
    
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
        
        .MaxCols = 2
        .MaxRows = 0
        
'        .Col = .MaxCols - 1
'        .ColHidden = True
'
'        .Col = .MaxCols
'        .ColHidden = True
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "관측소"
        .ColWidth(1) = 7
        .SetText 2, 0, "관측시간"
        .ColWidth(2) = 15
    End With
End Sub

Public Sub Init_fpSpread_Tot_Tw()
    
    With mainFrm.fpSpread_Tot_Tw
    
        '.Reset
        .MaxRows = 0
        .MaxRows = 500
        
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
        
        .MaxCols = 3
        .MaxRows = 0
        
'        .Col = .MaxCols - 1
'        .ColHidden = True
'
'        .Col = .MaxCols
'        .ColHidden = True
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "관측소"
        '.ColWidth(1) = 7
        .SetText 2, 0, "관측시간"
        '.ColWidth(2) = 15
        .SetText 3, 0, "업체"
    End With
End Sub

Public Sub Init_fpSpread_Tot_Rt()
    With mainFrm.fpSpread_Tot_Rt
    
        '.Reset
        .MaxRows = 0
        .MaxRows = 500
        
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
        
        .MaxCols = 3
        .MaxRows = 0
        
'        .Col = .MaxCols - 1
'        .ColHidden = True
'
'        .Col = .MaxCols
'        .ColHidden = True
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "관측소"
        '.ColWidth(1) = 7
        .SetText 2, 0, "관측시간"
        '.ColWidth(2) = 15
        .SetText 3, 0, "업체"
    End With
End Sub

Public Sub Init_fpSpread_Tot_Ag()
    With mainFrm.fpSpread_Tot_Ag
    
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
        
        .MaxCols = 2
        .MaxRows = 0
        
'        .Col = .MaxCols - 1
'        .ColHidden = True
'
'        .Col = .MaxCols
'        .ColHidden = True
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "부이ID"
        .ColWidth(1) = 7
        .SetText 2, 0, "관측시간"
        .ColWidth(2) = 15
    End With
End Sub

Public Sub Init_fpSpread_Tot_Usn()
    With mainFrm.fpSpread_Tot_Usn
    
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
        
        .MaxCols = 2
        .MaxRows = 0
        
'        .Col = .MaxCols - 1
'        .ColHidden = True
'
'        .Col = .MaxCols
'        .ColHidden = True
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "부이ID"
        .ColWidth(1) = 7
        .SetText 2, 0, "관측시간"
        .ColWidth(2) = 15
    End With
End Sub

Public Sub getJowiVpnList()

Dim Cnt As Integer

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
             strQry = " SELECT TS_NAME, TO_CHAR(DT_TIME,'yyyy/mm/dd hh24:mi:ss') DT_TIME " + vbCrLf
    strQry = strQry + "   FROM ( SELECT A.TS_NAME, B.DT_TIME " + vbCrLf
    strQry = strQry + "            FROM RTDB.TIDAL_STATION A, RTDB.DT_MAX_DT_TIME B " + vbCrLf
    strQry = strQry + "           WHERE A.TS_ID = B.DT_TS_ID " + vbCrLf
    strQry = strQry + "             AND A.COLLECTION_TYPE = 'V' " + vbCrLf
    strQry = strQry + "             AND TO_CHAR(B.DT_TIME,'yyyy') > '2011' ) " + vbCrLf
    strQry = strQry + "  ORDER BY DT_TIME ASC "
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    If Not (dtRs.EOF Or dtRs.BOF) Then
        mainFrm.fpSpread_Tot_DtVPN.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_Tot_DtVPN.MaxRows = dtRs.RecordCount + 1
        
        mainFrm.fpSpread_Tot_DtVPN.SetActiveCell 1, mainFrm.fpSpread_Tot_DtVPN.MaxRows
        mainFrm.fpSpread_Tot_DtVPN.ShowCell 1, 1, PositionBottomCenter
        
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_Tot_DtVPN.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next Cnt
        
        For Cnt = 0 To mainFrm.fpSpread_Tot_DtVPN.MaxRows - 1
            mainFrm.fpSpread_Tot_DtVPN.Col = 2
            mainFrm.fpSpread_Tot_DtVPN.Row = Cnt + 1

            If DateDiff("d", CDate(mainFrm.fpSpread_Tot_DtVPN.Text), Now) >= 1 Then
                For j = 0 To mainFrm.fpSpread_Tot_DtVPN.MaxCols - 1
                    mainFrm.fpSpread_Tot_DtVPN.Col = j + 1
                    mainFrm.fpSpread_Tot_DtVPN.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_DtVPN.BackColor = vbRed
                    mainFrm.fpSpread_Tot_DtVPN.ForeColor = vbWhite
                Next j
                
                mainFrm.fpSpread_Tot_DtVPN.Col = 2
                mainFrm.fpSpread_Tot_DtVPN.Row = Cnt + 1
            
            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_Tot_DtVPN.Text), Now) >= strJowiVPNCautionMin Then
                    For j = 0 To mainFrm.fpSpread_Tot_DtVPN.MaxCols - 1
                        mainFrm.fpSpread_Tot_DtVPN.Col = j + 1
                        mainFrm.fpSpread_Tot_DtVPN.Row = Cnt + 1
                        mainFrm.fpSpread_Tot_DtVPN.BackColor = vbYellow
                        mainFrm.fpSpread_Tot_DtVPN.ForeColor = vbBlack
                    Next j
                    
                    mainFrm.fpSpread_Tot_DtVPN.Col = 2
                    mainFrm.fpSpread_Tot_DtVPN.Row = Cnt + 1
            Else
                For j = 0 To mainFrm.fpSpread_Tot_DtVPN.MaxCols - 1
                    mainFrm.fpSpread_Tot_DtVPN.Col = j + 1
                    mainFrm.fpSpread_Tot_DtVPN.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_DtVPN.BackColor = vbWhite
                    mainFrm.fpSpread_Tot_DtVPN.ForeColor = vbBlack
                Next j
            End If
        Next Cnt
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub getJowiCdmaList()

Dim Cnt As Integer

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
             strQry = "SELECT TS_NAME , TO_CHAR(DT_TIME,'yyyy/mm/dd hh24:mi:ss') DT_TIME " + vbCrLf
    strQry = strQry + "  FROM ( " + vbCrLf
    strQry = strQry + "    SELECT A.TS_NAME, B.DT_TIME " + vbCrLf
    strQry = strQry + "    FROM RTDB.TIDAL_STATION A, RTDB.DT_MAX_DT_TIME B " + vbCrLf
    strQry = strQry + "    Where A.TS_ID = B.DT_TS_ID " + vbCrLf
    strQry = strQry + "      AND A.COLLECTION_TYPE = 'C' " + vbCrLf
    strQry = strQry + "    ) " + vbCrLf
    strQry = strQry + " ORDER BY DT_TIME ASC " + vbCrLf
    
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
        mainFrm.fpSpread_Tot_DtCDMA.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_Tot_DtCDMA.MaxRows = dtRs.RecordCount + 1
        
        mainFrm.fpSpread_Tot_DtCDMA.SetActiveCell 1, mainFrm.fpSpread_Tot_DtCDMA.MaxRows
    
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_Tot_DtCDMA.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next Cnt
        
        '색상처리
        For Cnt = 0 To mainFrm.fpSpread_Tot_DtCDMA.MaxRows - 1
            mainFrm.fpSpread_Tot_DtCDMA.Col = 2
            mainFrm.fpSpread_Tot_DtCDMA.Row = Cnt + 1

            If DateDiff("d", CDate(mainFrm.fpSpread_Tot_DtCDMA.Text), Now) >= 1 Then
                '-- 셀의 배경색상 변경
                For j = 0 To mainFrm.fpSpread_Tot_DtCDMA.MaxCols - 1
                    mainFrm.fpSpread_Tot_DtCDMA.Col = j + 1
                    mainFrm.fpSpread_Tot_DtCDMA.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_DtCDMA.BackColor = vbRed
                    mainFrm.fpSpread_Tot_DtCDMA.ForeColor = vbWhite
                Next j
                
                mainFrm.fpSpread_Tot_DtCDMA.Col = 2
                mainFrm.fpSpread_Tot_DtCDMA.Row = Cnt + 1
                'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_Tot_DtCDMA.Text), Now) >= strJowiCDMACautionMin Then
                '-- 셀의 배경색상 변경
                For j = 0 To mainFrm.fpSpread_Tot_DtCDMA.MaxCols - 1
                    mainFrm.fpSpread_Tot_DtCDMA.Col = j + 1
                    mainFrm.fpSpread_Tot_DtCDMA.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_DtCDMA.BackColor = vbYellow
                    mainFrm.fpSpread_Tot_DtCDMA.ForeColor = vbBlack
                Next j
                
                mainFrm.fpSpread_Tot_DtCDMA.Col = 2
                mainFrm.fpSpread_Tot_DtCDMA.Row = Cnt + 1
                'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
            Else
                For j = 0 To mainFrm.fpSpread_Tot_DtCDMA.MaxCols - 1
                    mainFrm.fpSpread_Tot_DtCDMA.Col = j + 1
                    mainFrm.fpSpread_Tot_DtCDMA.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_DtCDMA.BackColor = vbWhite
                    mainFrm.fpSpread_Tot_DtCDMA.ForeColor = vbBlack
                Next j
            End If
        Next Cnt
        
    
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub getTWList()

Dim Cnt As Integer

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Dim intGeoCnt   As Integer  'Geo시스템 해양부이 갯수
    Dim intOceanCnt As Integer  '오션테크 해양부이 갯수
        
    Set dtRs = New ADODB.Recordset

             strQry = "SELECT STATION_NAME, TO_CHAR(OBS_TIME,'yyyy/mm/dd hh24:mi:ss') OBS_TIME, COMPANY " + vbCrLf
    strQry = strQry + "  FROM(  " + vbCrLf
    strQry = strQry + "SELECT STATION_NAME, C.OBS_TIME, '오션테크' AS COMPANY  " + vbCrLf
    strQry = strQry + "FROM WRN.T_WRN_STATION A, WRN.T_WRN_TW_BUOY B, (SELECT STATION_ID, MAX(OBS_TIME) AS OBS_TIME " + vbCrLf
    strQry = strQry + "                                                  FROM WRN.T_WRN_TW_BUOY" + vbCrLf
    strQry = strQry + "                                                 WHERE TO_CHAR(OBS_TIME,'yyyy') > '2011' " + vbCrLf
    strQry = strQry + "                                              GROUP BY STATION_ID) C  " + vbCrLf
    strQry = strQry + "WHERE A.STATION_ID = B.STATION_ID  " + vbCrLf
    strQry = strQry + "  AND B.STATION_ID = C.STATION_ID  " + vbCrLf
    strQry = strQry + "  AND B.OBS_TIME = C.OBS_TIME  " + vbCrLf
    strQry = strQry + "UNION ALL " + vbCrLf
    strQry = strQry + "SELECT STATION_NAME, C.OBS_TIME, 'GEO'  " + vbCrLf
    strQry = strQry + "  FROM WRN.T_WRN_STATION A, WRN.T_RDCP_BUOY B, (SELECT STATION_ID, MAX(RD_OBS_TIME) AS OBS_TIME " + vbCrLf
    strQry = strQry + "                                                 FROM WRN.T_RDCP_BUOY" + vbCrLf
    strQry = strQry + "                                                 WHERE TO_CHAR(RD_OBS_TIME,'yyyy') > '2011' " + vbCrLf
    strQry = strQry + "                                              GROUP BY STATION_ID) C  " + vbCrLf
    strQry = strQry + "WHERE A.STATION_ID = B.STATION_ID  " + vbCrLf
    strQry = strQry + "  AND B.STATION_ID = C.STATION_ID  " + vbCrLf
    strQry = strQry + "  AND B.RD_OBS_TIME = C.OBS_TIME  " + vbCrLf
    strQry = strQry + ")  " + vbCrLf
    strQry = strQry + "ORDER BY OBS_TIME ASC " + vbCrLf

    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
        '변수 초기화
        intGeoCnt = 0
        intOceanCnt = 0
                
        mainFrm.fpSpread_Tot_Tw.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_Tot_Tw.MaxRows = dtRs.RecordCount + 1
        
        mainFrm.fpSpread_Tot_Tw.SetActiveCell 1, mainFrm.fpSpread_Tot_Tw.MaxRows
        mainFrm.fpSpread_Tot_Tw.ShowCell 1, 1, PositionBottomCenter
    
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_Tot_Tw.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            
            If dtRs.Fields("COMPANY") = "GEO" Then  '각 회사별 해양관측 부이 Count
                intGeoCnt = intGeoCnt + 1   'Geo
            Else
                intOceanCnt = intOceanCnt + 1   'Ocean
            End If
            
            dtRs.MoveNext
            
        Next Cnt
                
        '색상처리
        For Cnt = 0 To mainFrm.fpSpread_Tot_Tw.MaxRows - 1
            mainFrm.fpSpread_Tot_Tw.Col = 2
            mainFrm.fpSpread_Tot_Tw.Row = Cnt + 1

            If DateDiff("d", CDate(mainFrm.fpSpread_Tot_Tw.Text), Now) >= 1 Then    '전송안됨(빨간색)
                '-- 셀의 배경색상 변경
                For j = 0 To mainFrm.fpSpread_Tot_Tw.MaxCols - 1
                    mainFrm.fpSpread_Tot_Tw.Col = j + 1
                    mainFrm.fpSpread_Tot_Tw.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_Tw.BackColor = vbRed
                    mainFrm.fpSpread_Tot_Tw.ForeColor = vbWhite
                Next j
                
                mainFrm.fpSpread_Tot_Tw.Col = 2
                mainFrm.fpSpread_Tot_Tw.Row = Cnt + 1
                
            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_Tot_Tw.Text), Now) >= strTwCautionMin Then  '전송지연(노란색)
                '-- 셀의 배경색상 변경
                For j = 0 To mainFrm.fpSpread_Tot_Tw.MaxCols - 1
                    mainFrm.fpSpread_Tot_Tw.Col = j + 1
                    mainFrm.fpSpread_Tot_Tw.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_Tw.BackColor = vbYellow
                    mainFrm.fpSpread_Tot_Tw.ForeColor = vbBlack
                Next j
                
                mainFrm.fpSpread_Tot_Tw.Col = 2
                mainFrm.fpSpread_Tot_Tw.Row = Cnt + 1
                
            Else    '전송 전상처리
                For j = 0 To mainFrm.fpSpread_Tot_DtCDMA.MaxCols - 1
                    mainFrm.fpSpread_Tot_Tw.Col = j + 1
                    mainFrm.fpSpread_Tot_Tw.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_Tw.BackColor = vbWhite
                    mainFrm.fpSpread_Tot_Tw.ForeColor = vbBlack
                Next j
            End If
        Next Cnt
        
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
        
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub getAGList()

Dim Cnt As Integer

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
             strQry = "SELECT TYP_BUOY_ID, TO_CHAR(POS_TIME,'yyyy/mm/dd hh24:mi:ss') " + vbCrLf
    strQry = strQry + "  FROM ( " + vbCrLf
    strQry = strQry + "SELECT A.TYP_BUOY_ID, C.POS_TIME " + vbCrLf
    strQry = strQry + "FROM WRN.T_WRN_TYP_BUOY_ID A, WRN.T_WRN_TYP_BUOY B, (SELECT TYP_BUOY_ID, MAX(POS_TIME) AS POS_TIME FROM WRN.T_WRN_TYP_BUOY GROUP BY TYP_BUOY_ID) C " + vbCrLf
    strQry = strQry + "WHERE A.TYP_BUOY_ID = B.TYP_BUOY_ID " + vbCrLf
    strQry = strQry + "  AND B.TYP_BUOY_ID = C.TYP_BUOY_ID " + vbCrLf
    strQry = strQry + "  AND B.POS_TIME = C.POS_TIME " + vbCrLf
    strQry = strQry + ") " + vbCrLf
    strQry = strQry + "ORDER BY POS_TIME ASC " + vbCrLf

    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
        mainFrm.fpSpread_Tot_Ag.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_Tot_Ag.MaxRows = dtRs.RecordCount + 1
        
        mainFrm.fpSpread_Tot_Ag.SetActiveCell 1, mainFrm.fpSpread_Tot_Ag.MaxRows
        mainFrm.fpSpread_Tot_Ag.ShowCell 1, 1, PositionBottomCenter
    
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_Tot_Ag.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next Cnt
        
        
        '색상처리
        For Cnt = 0 To mainFrm.fpSpread_Tot_Ag.MaxRows - 1
            mainFrm.fpSpread_Tot_Ag.Col = 2
            mainFrm.fpSpread_Tot_Ag.Row = Cnt + 1

            If DateDiff("d", CDate(mainFrm.fpSpread_Tot_Ag.Text), Now) >= 1 Then
                '-- 셀의 배경색상 변경
                For j = 0 To mainFrm.fpSpread_Tot_Ag.MaxCols - 1
                    mainFrm.fpSpread_Tot_Ag.Col = j + 1
                    mainFrm.fpSpread_Tot_Ag.Row = Cnt + 1
                    'CCY 20100809 협회요구 mainFrm.fpSpread_Tot_Ag.BackColor = vbRed
                    'CCY 20100809 협회요구 mainFrm.fpSpread_Tot_Ag.ForeColor = vbWhite
                Next j
                
                mainFrm.fpSpread_Tot_Ag.Col = 2
                mainFrm.fpSpread_Tot_Ag.Row = Cnt + 1
            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_Tot_Ag.Text), Now) >= strAgCautionMin Then
                    '-- 셀의 배경색상 변경
                    For j = 0 To mainFrm.fpSpread_Tot_Ag.MaxCols - 1
                        mainFrm.fpSpread_Tot_Ag.Col = j + 1
                        mainFrm.fpSpread_Tot_Ag.Row = Cnt + 1
                        'CCY 20100809 협회요구 mainFrm.fpSpread_Tot_Ag.BackColor = vbYellow
                        'CCY 20100809 협회요구 mainFrm.fpSpread_Tot_Ag.ForeColor = vbBlack
                    Next j
                    
                    mainFrm.fpSpread_Tot_Ag.Col = 2
                    mainFrm.fpSpread_Tot_Ag.Row = Cnt + 1
            Else

                For j = 0 To mainFrm.fpSpread_Tot_DtCDMA.MaxCols - 1
                    mainFrm.fpSpread_Tot_Ag.Col = j + 1
                    mainFrm.fpSpread_Tot_Ag.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_Ag.BackColor = vbWhite
                    mainFrm.fpSpread_Tot_Ag.ForeColor = vbBlack
                Next j
            End If
        Next Cnt
        
    
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

'-- 2014.04.18 osw
'-- USN 모니터링
Public Sub subRDCPDisplay(ByVal spdDisplay As Object, ByVal strTbl As String)
        
Dim Cnt As Integer

On Error Resume Next
    
    Dim RsRDCP As ADODB.Recordset
    Set RsRDCP = New ADODB.Recordset
        
    '연결상태를 체크하여 재 연결
    If Not AdoDBConn.State = adStateOpen Then
        If AdoDBConn.State = adStateOpen Then AdoDBConn.Close
        If Not AdoDBConn Is Nothing Then Set AdoDBConn = Nothing
        
        Set AdoDBConn = New ADODB.Connection
        AdoDBConn.Open strAdoDBConn
    End If

    strQry = ""
    strQry = strQry & "SELECT b.station_name as 관측소명,MAX(a.OBS_TIME) as 관측시간"
    strQry = strQry & "  FROM " & strTbl & " a, TK_USN_STATION b"
    strQry = strQry & " WHERE a.SYSTEM_ID = b.SYSTEM_id"
    strQry = strQry & " GROUP BY b.station_name, a.SYSTEM_ID, a.NODE_ID"
    strQry = strQry & " ORDER BY a.SYSTEM_ID"

    RsRDCP.CursorLocation = adUseClient
    RsRDCP.Open strQry, AdoDBConn

    If Not (RsRDCP.EOF Or RsRDCP.BOF) Then
        spdDisplay.MaxCols = RsRDCP.Fields.Count
        spdDisplay.MaxRows = RsRDCP.RecordCount + 1
        
        spdDisplay.SetActiveCell 1, spdDisplay.MaxRows
        spdDisplay.ShowCell 1, 1, PositionBottomCenter
    
        For Cnt = 0 To RsRDCP.RecordCount - 1
            Dim j As Integer
            For j = 0 To RsRDCP.Fields.Count - 1
                Call spdDisplay.SetText(j + 1, Cnt + 1, RsRDCP(j))
            Next j
            RsRDCP.MoveNext
        Next Cnt
        
        '색상처리
        For Cnt = 0 To spdDisplay.MaxRows - 1
            spdDisplay.Col = 2
            spdDisplay.Row = Cnt + 1
            '-- 30분 이내
            If DateDiff("n", CDate(spdDisplay.Text), Now) > -1 And DateDiff("n", CDate(spdDisplay.Text), Now) < 30 Then
                For j = 0 To spdDisplay.MaxCols - 1
                    spdDisplay.Col = j + 1
                    spdDisplay.Row = Cnt + 1
                    spdDisplay.BackColor = vbWhite
                    spdDisplay.ForeColor = vbBlack
                Next j
                
                spdDisplay.Col = 2
                spdDisplay.Row = Cnt + 1
            '-- 30 ~ 65분
            ElseIf DateDiff("n", CDate(spdDisplay.Text), Now) > 29 And DateDiff("n", CDate(spdDisplay.Text), Now) < 65 Then
                    For j = 0 To spdDisplay.MaxCols - 1
                        spdDisplay.Col = j + 1
                        spdDisplay.Row = Cnt + 1
                        spdDisplay.BackColor = vbYellow
                        spdDisplay.ForeColor = vbBlack
                    Next j
                    
                    spdDisplay.Col = 2
                    spdDisplay.Row = Cnt + 1
            '-- 65분이상
            Else
                For j = 0 To mainFrm.fpSpread_Tot_DtCDMA.MaxCols - 1
                    spdDisplay.Col = j + 1
                    spdDisplay.Row = Cnt + 1
                    spdDisplay.BackColor = vbRed
                    spdDisplay.ForeColor = vbWhite
                Next j
            End If
        Next Cnt
    End If
    
End Sub

Public Sub getRTList()

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
        
    '연결상태를 체크하여 재 연결
    If Not AdoDBConn.State = adStateOpen Then
        If AdoDBConn.State = adStateOpen Then AdoDBConn.Close
        If Not AdoDBConn Is Nothing Then Set AdoDBConn = Nothing
        
        Set AdoDBConn = New ADODB.Connection
        AdoDBConn.Open strAdoDBConn
    End If
    
    strQry = ""
    
    strQry = "SELECT * " + vbCrLf
    strQry = strQry + "FROM ( " + vbCrLf
    strQry = strQry + "    SELECT NAME, TO_CHAR(DTIME,'yyyy/mm/dd hh24:mi:ss') DTIME, 'GEO' AS COMPANY " + vbCrLf
    strQry = strQry + "    FROM REALTIME.STATION A, (SELECT SID, MAX(DTIME) DTIME " + vbCrLf
    strQry = strQry + "                     FROM REALTIME.LOG_REALTIME_DATA " + vbCrLf
    strQry = strQry + "                     WHERE LOG_ID = 'L006'  " + vbCrLf
    strQry = strQry + "                       AND TO_CHAR(DTIME,'yyyy') > '2011' " + vbCrLf
    strQry = strQry + "                    GROUP BY SID) B  " + vbCrLf
    strQry = strQry + "   WHERE A.SID = B.SID " + vbCrLf
    strQry = strQry + "     AND A.USE_YN = 'Y'  " + vbCrLf
    strQry = strQry + ") " + vbCrLf
    strQry = strQry + "WHERE DTIME IS NOT NULL " + vbCrLf
    strQry = strQry + "ORDER BY DTIME ASC " + vbCrLf

'Call LogWrite(strQry)
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
'        If mainFrm.AdoDBConn.Errors.Count = 0 Then
'            mainFrm.fpSpread_Tot_Rt.DataSource = dtRs.DataSource
'        Else
'            AdoDBConn.Errors.Clear
'        End If
        Dim Cnt As Integer
        
        mainFrm.fpSpread_Tot_Rt.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_Tot_Rt.MaxRows = dtRs.RecordCount + 1
        
        mainFrm.fpSpread_Tot_Rt.SetActiveCell 1, mainFrm.fpSpread_Tot_Rt.MaxRows
        mainFrm.fpSpread_Tot_Rt.ShowCell 1, 1, PositionBottomCenter
    
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_Tot_Rt.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next Cnt
        
        
        '색상처리
        For Cnt = 0 To mainFrm.fpSpread_Tot_Rt.MaxRows - 1
            mainFrm.fpSpread_Tot_Rt.Col = 2
            mainFrm.fpSpread_Tot_Rt.Row = Cnt + 1

            If DateDiff("d", CDate(mainFrm.fpSpread_Tot_Rt.Text), Now) >= 1 Then
                '-- 셀의 배경색상 변경
                For j = 0 To mainFrm.fpSpread_Tot_Rt.MaxCols - 1
                    mainFrm.fpSpread_Tot_Rt.Col = j + 1
                    mainFrm.fpSpread_Tot_Rt.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_Rt.BackColor = vbRed
                    mainFrm.fpSpread_Tot_Rt.ForeColor = vbWhite
                Next j
                
                mainFrm.fpSpread_Tot_Rt.Col = 2
                mainFrm.fpSpread_Tot_Rt.Row = Cnt + 1
                'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_Tot_Rt.Text), Now) >= strRtCautionMin Then
                    '-- 셀의 배경색상 변경
                    For j = 0 To mainFrm.fpSpread_Tot_Rt.MaxCols - 1
                        mainFrm.fpSpread_Tot_Rt.Col = j + 1
                        mainFrm.fpSpread_Tot_Rt.Row = Cnt + 1
                        mainFrm.fpSpread_Tot_Rt.BackColor = vbYellow
                        mainFrm.fpSpread_Tot_Rt.ForeColor = vbBlack
                    Next j
                    
                    mainFrm.fpSpread_Tot_Rt.Col = 2
                    mainFrm.fpSpread_Tot_Rt.Row = Cnt + 1
                    'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
            Else
                For j = 0 To mainFrm.fpSpread_Tot_DtCDMA.MaxCols - 1
                    mainFrm.fpSpread_Tot_Rt.Col = j + 1
                    mainFrm.fpSpread_Tot_Rt.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_Rt.BackColor = vbWhite
                    mainFrm.fpSpread_Tot_Rt.ForeColor = vbBlack
                Next j
            End If
        Next Cnt
        
    
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub getUSNList()

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    strQry = "SELECT STATION_NAME, TO_CHAR(B.OBS_TIME,'yyyy/mm/dd hh24:mi:ss') OBS_TIME " + vbCrLf
    strQry = strQry + " FROM USN.TK_USN_STATION_CONFIG A , (SELECT STATION_ID, MAX(OBS_TIME) OBS_TIME FROM USN.TK_USN_OBSERVATION GROUP BY STATION_ID) B " + vbCrLf
    strQry = strQry + " Where A.STATION_ID = B.STATION_ID " + vbCrLf
    strQry = strQry + "ORDER BY OBS_TIME ASC " + vbCrLf

'Call LogWrite(strQry)

    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then

        Dim Cnt As Integer
        
        
        
        mainFrm.fpSpread_Tot_Usn.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_Tot_Usn.MaxRows = dtRs.RecordCount + 1
        
        mainFrm.fpSpread_Tot_Usn.SetActiveCell 1, mainFrm.fpSpread_Tot_Usn.MaxRows
        mainFrm.fpSpread_Tot_Usn.ShowCell 1, 1, PositionBottomCenter
        
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_Tot_Usn.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next Cnt
        
        
        '색상처리
        For Cnt = 0 To mainFrm.fpSpread_Tot_Usn.MaxRows - 1
            mainFrm.fpSpread_Tot_Usn.Col = 2
            mainFrm.fpSpread_Tot_Usn.Row = Cnt + 1

'MsgBox Now & "-" & CDate(mainFrm.fpSpread_Tot_Usn.Text) & "====" & DateDiff("n", CDate(mainFrm.fpSpread_Tot_Usn.Text), Now)
            If DateDiff("d", CDate(mainFrm.fpSpread_Tot_Usn.Text), Now) >= 1 Then
                '-- 셀의 배경색상 변경
                For j = 0 To mainFrm.fpSpread_Tot_Usn.MaxCols - 1
                    mainFrm.fpSpread_Tot_Usn.Col = j + 1
                    mainFrm.fpSpread_Tot_Usn.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_Usn.BackColor = vbRed
                    mainFrm.fpSpread_Tot_Usn.ForeColor = vbWhite
                Next j
                
                mainFrm.fpSpread_Tot_Usn.Col = 2
                mainFrm.fpSpread_Tot_Usn.Row = Cnt + 1
                'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_Tot_Usn.Text), Now) >= strAgCautionMin Then
                    '-- 셀의 배경색상 변경
                    For j = 0 To mainFrm.fpSpread_Tot_Usn.MaxCols - 1
                        mainFrm.fpSpread_Tot_Usn.Col = j + 1
                        mainFrm.fpSpread_Tot_Usn.Row = Cnt + 1
                        mainFrm.fpSpread_Tot_Usn.BackColor = vbYellow
                        mainFrm.fpSpread_Tot_Usn.ForeColor = vbBlack
                    Next j
                    
                    mainFrm.fpSpread_Tot_Usn.Col = 2
                    mainFrm.fpSpread_Tot_Usn.Row = Cnt + 1
                    'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
            Else

                For j = 0 To mainFrm.fpSpread_Tot_Usn.MaxCols - 1
                    mainFrm.fpSpread_Tot_Usn.Col = j + 1
                    mainFrm.fpSpread_Tot_Usn.Row = Cnt + 1
                    mainFrm.fpSpread_Tot_Usn.BackColor = vbWhite
                    mainFrm.fpSpread_Tot_Usn.ForeColor = vbBlack
                Next j
            End If
        Next Cnt
        
    
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub
