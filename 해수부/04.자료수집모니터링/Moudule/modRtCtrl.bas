Attribute VB_Name = "modRtCtrl"
Option Explicit

Public Sub startRTView()
    
    mainFrm.chkTimer_Rt = 1

End Sub

Public Sub setRtStationID()

Dim Cnt As Integer

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    strQry = ""
    strQry = strQry + "SELECT NAME, B.SID " + vbCrLf
    strQry = strQry + "  FROM REALTIME.STATION A, (SELECT DISTINCT SID " + vbCrLf
    strQry = strQry + "                     From REALTIME.LOG_REALTIME_DATA " + vbCrLf
    strQry = strQry + "                    ) B " + vbCrLf
    strQry = strQry + " WHERE A.SID = B.SID " + vbCrLf
    strQry = strQry + "   AND FILENAME IS NOT NULL " + vbCrLf
    strQry = strQry + " ORDER BY NAME " + vbCrLf
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn
    
    mainFrm.cmbxSechRTID.AddItem "전체"
    
    If Not (dtRs.EOF Or dtRs.BOF) Then
        For Cnt = 0 To dtRs.RecordCount - 1
            mainFrm.cmbxSechRTID.AddItem dtRs.Fields("NAME")
            mainFrm.CboRTID_NM.AddItem dtRs.Fields("NAME")
            mainFrm.CboRTID_ID.AddItem dtRs.Fields("SID")

            dtRs.MoveNext
        Next Cnt
        
        mainFrm.cmbxSechRTID.ListIndex = 0
        mainFrm.CboRTID_NM.ListIndex = 0
        mainFrm.CboRTID_ID.ListIndex = mainFrm.CboRTID_NM.ListIndex
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub setRtSechCondition()

    Dim Cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '검색 일자 설정
    mainFrm.txtSechRTStDate.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtSechRTEdDate.Text = Format(Now, "YYYY-MM-DD")
End Sub


Public Sub goRtSearch()
        
Dim Cnt As Integer
    
    '검색조건 체크 st
    If Not IsDate(mainFrm.txtSechRTStDate.Text) Then
        MsgBox "검색 범위 시작일자를 확인해주세요."
        mainFrm.txtSechRTStDate.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mainFrm.txtSechRTEdDate.Text) Then
        MsgBox "검색 범위 종료일시를 확인해주세요."
        mainFrm.txtSechRTEdDate.SetFocus
        Exit Sub
    End If
    
    If mainFrm.cmbxRtRownum.Text = "ALL" Then
    ElseIf Not IsNumeric(mainFrm.cmbxRtRownum.Text) Then
        MsgBox "출력건수는 숫자를 입력해주세요."
        mainFrm.cmbxRtRownum.SetFocus
        Exit Sub
    End If

    Call Init_fpSpread_RtLog

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    Dim stDate As String
    Dim edDate As String
    
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn
    
    stDate = mainFrm.txtSechRTStDate.Text
    edDate = mainFrm.txtSechRTEdDate.Text
    
    strQry = ""
    strQry = strQry + "SELECT * FROM( " + vbCrLf
    strQry = strQry + "SELECT  B.SID " + vbCrLf
    strQry = strQry + "       , NAME " + vbCrLf
    strQry = strQry + "       , CASE WHEN LOG_ID='L006' Then TO_CHAR(DTIME,'yyyy/mm/dd hh24:mi:ss') End DT_TIME " + vbCrLf
    strQry = strQry + "       , TO_CHAR(REG_DATE,'yyyy/mm/dd hh24:mi:ss') REG_DATE " + vbCrLf
    strQry = strQry + "       , LOG_CONTENT  " + vbCrLf
    strQry = strQry + "  FROM REALTIME.STATION A, REALTIME.LOG_REALTIME_DATA B  " + vbCrLf
    strQry = strQry + " WHERE A.SID = B.SID  " + vbCrLf
    strQry = strQry + "   AND B.SID  > 0  " + vbCrLf
    If Not mainFrm.cmbxSechRTID.Text = "전체" Then
        strQry = strQry + "  AND B.SID IN ('0', (SELECT SID FROM REALTIME.STATION WHERE NAME = '" + mainFrm.cmbxSechRTID.Text + "')) " + vbCrLf
    End If
    strQry = strQry + "  AND REG_DATE >= TO_DATE('" + stDate + "000000', 'YYYY-MM-DDHH24MISS')  " + vbCrLf
    strQry = strQry + "  AND REG_DATE <= TO_DATE('" + edDate + "235959', 'YYYY-MM-DDHH24MISS')  " + vbCrLf
    strQry = strQry + "ORDER BY REG_DATE DESC, SID DESC, LOG_ID DESC " + vbCrLf
    strQry = strQry + ") " + vbCrLf
    If Not mainFrm.cmbxRtRownum.Text = "ALL" Then
        strQry = strQry + "WHERE ROWNUM <= " + mainFrm.cmbxRtRownum.Text
    End If
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
        mainFrm.fpSpread_RtLog.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_RtLog.MaxRows = dtRs.RecordCount
    
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_RtLog.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next Cnt
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    
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

Public Sub Init_fpSpread_RtLog()
    
    With mainFrm.fpSpread_RtLog

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

'''-- 2014.04.17 osw
''Public Sub Init_fpSpread_RDCP()
''
''    With mainFrm.sprRDCP
''
''        .Reset
''
''        .OperationMode = OperationModeRow
''        .GridSolid = False
''
''        .Appearance = Appearance3D
''
''        'Hide row header
''        .RowHeadersShow = False
''
''        'Turn off font bold
''        .Col = -1
''        .Row = -1
''        .FontBold = False
''
''        'Change the amount of data each cell will hold
''        .Col = -1
''        .Row = -1
''        .TypeEditLen = 200
''
''        'Set column display type
''        .ColHeaderDisplay = DispBlank
''        .AllowCellOverflow = True
''        .ReDraw = True
''
''        .ShowScrollTips = ShowScrollTipsVertical
''        .GrayAreaBackColor = &HFFFFFF
''
''        .TextTip = TextTipFloating
''
''        .MaxCols = 5
''        .MaxRows = 0
''
''        .RowHeight(0) = 15
''
''        .SetText 1, 0, "관측소명"
''        .ColWidth(1) = 10
''        .SetText 2, 0, "관측시간"
''        .ColWidth(2) = 15
''    End With
''
''End Sub

