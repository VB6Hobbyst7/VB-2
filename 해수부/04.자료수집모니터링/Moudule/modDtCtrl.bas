Attribute VB_Name = "modDtCtrl"
Option Explicit

Public Sub startDTView()
    
    '검색창에 시간 정보 셋팅
    Call setDtSechCondition
    
    mainFrm.chkTimer_Dt = 1

End Sub

Public Sub setDtStationID()

Dim Cnt As Integer

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    strQry = "SELECT TS_NAME  " + vbCrLf
    strQry = strQry + "FROM RTDB.TIDAL_STATION " + vbCrLf
    strQry = strQry + "WHERE NOT VPN_IP IS NULL " + vbCrLf
    strQry = strQry + "ORDER BY TS_NAME ASC " + vbCrLf

    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    mainFrm.cmbxSechDTID.AddItem "전체"

    If Not (dtRs.EOF Or dtRs.BOF) Then
        For Cnt = 0 To dtRs.RecordCount - 1
            mainFrm.cmbxSechDTID.AddItem dtRs.Fields("TS_NAME")
            dtRs.MoveNext
        Next Cnt
    
        mainFrm.cmbxSechDTID.ListIndex = 0
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub setDtSechCondition()

    Dim Cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    
    strNowHour = Format(Now, "hh")
    '검색 일자 설정
    mainFrm.txtSechDTStDate.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtSechDTEdDate.Text = Format(Now, "YYYY-MM-DD")
    '검색 시간 설정
    For Cnt = 0 To 23
        If Cnt < 10 Then
            strHour = "0" & Cnt
        Else
            strHour = Cnt
        End If
        
        If strHour = strNowHour Then
             strNowIdx = Cnt
        End If
        mainFrm.cmbxSechDTStHour.AddItem strHour
        mainFrm.cmbxSechDTEdHour.AddItem strHour
    Next Cnt
    
    mainFrm.cmbxSechDTEdHour.ListIndex = strNowIdx
    If strNowIdx = 0 Then
        mainFrm.cmbxSechDTStHour.ListIndex = strNowIdx
    Else
        mainFrm.cmbxSechDTStHour.ListIndex = strNowIdx - 1
    End If
    
End Sub

Public Sub Sub_SetTwDate()

    Dim Cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '검색 일자 설정
    mainFrm.txtTwDate_From.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtTwDate_To.Text = Format(Now, "YYYY-MM-DD")
    '검색 시간 설정
    For Cnt = 0 To 23
        If Cnt < 10 Then
            strHour = "0" & Cnt
        Else
            strHour = Cnt
        End If
        
        If strHour = strNowHour Then
             strNowIdx = Cnt
        End If
        mainFrm.cboTwhh_From.AddItem strHour
        mainFrm.cboTwhh_To.AddItem strHour
    Next Cnt
    
    mainFrm.cboTwhh_To.ListIndex = strNowIdx
    If strNowIdx = 0 Then
        mainFrm.cboTwhh_From.ListIndex = strNowIdx
    Else
        mainFrm.cboTwhh_From.ListIndex = strNowIdx - 1
    End If
    
End Sub

Public Sub Sub_SetRTIDDate()

    Dim Cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '검색 일자 설정
    mainFrm.txtRTIDDate_From.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtRTIDDate_To.Text = Format(Now, "YYYY-MM-DD")
    '검색 시간 설정
    For Cnt = 0 To 23
        If Cnt < 10 Then
            strHour = "0" & Cnt
        Else
            strHour = Cnt
        End If
        
        If strHour = strNowHour Then
             strNowIdx = Cnt
        End If
        mainFrm.cboRTIDhh_From.AddItem strHour
        mainFrm.cboRTIDhh_To.AddItem strHour
    Next Cnt
    
    mainFrm.cboRTIDhh_To.ListIndex = strNowIdx
    If strNowIdx = 0 Then
        mainFrm.cboRTIDhh_From.ListIndex = strNowIdx
    Else
        mainFrm.cboRTIDhh_From.ListIndex = strNowIdx - 1
    End If
    
End Sub

Public Sub Sub_Tw()

    Dim Cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '검색 일자 설정
    mainFrm.txtTwDate_From.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtTwDate_To.Text = Format(Now, "YYYY-MM-DD")
    '검색 시간 설정
    For Cnt = 0 To 23
        If Cnt < 10 Then
            strHour = "0" & Cnt
        Else
            strHour = Cnt
        End If
        
        If strHour = strNowHour Then
             strNowIdx = Cnt
        End If
        mainFrm.cboTwhh_From.AddItem strHour
        mainFrm.cboTwhh_To.AddItem strHour
    Next Cnt
    
    mainFrm.cboTwhh_To.ListIndex = strNowIdx
    If strNowIdx = 0 Then
        mainFrm.cboTwhh_From.ListIndex = strNowIdx
    Else
        mainFrm.cboTwhh_From.ListIndex = strNowIdx - 1
    End If
    
End Sub


Public Sub goDtSearch()
    '검색조건 체크 st
    If Not IsDate(mainFrm.txtSechDTStDate.Text) Then
        MsgBox "검색 범위 시작일자를 확인해주세요."
        mainFrm.txtSechDTStDate.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(mainFrm.cmbxSechDTStHour.Text) = False Then
        MsgBox "검색 범위 시작일시를 확인해주세요."
        mainFrm.cmbxSechDTStHour.SetFocus
        Exit Sub
    End If
    
    If mainFrm.cmbxSechDTStHour.Text > 24 Then
        MsgBox "검색 범위 시작일시는 숫자를 입력해주세요."
        mainFrm.cmbxSechDTStHour.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mainFrm.txtSechDTEdDate.Text) Then
        MsgBox "검색 범위 종료일시를 확인해주세요."
        mainFrm.txtSechDTEdDate.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(mainFrm.cmbxSechDTEdHour.Text) = False Then
        MsgBox "검색 범위 종료일시를 확인해주세요."
        mainFrm.cmbxSechDTEdHour.SetFocus
        Exit Sub
    End If
    
    If mainFrm.cmbxSechDTEdHour.Text > 24 Then
        MsgBox "검색 범위 종료일시는 숫자를 입력해주세요."
        mainFrm.cmbxSechDTEdHour.SetFocus
        Exit Sub
    End If
    
    If mainFrm.cmbxDtRownum.Text = "ALL" Then
    ElseIf Not IsNumeric(mainFrm.cmbxDtRownum.Text) Then
        MsgBox "출력건수는 숫자를 입력해주세요."
        mainFrm.cmbxDtRownum.SetFocus
        Exit Sub
    End If
    
    '검색창초기화
    Call Init_fpSpread_DtLog

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    Dim stDateHour As String
    Dim edDateHour As String
    
    'DB접속
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn
    
    stDateHour = mainFrm.txtSechDTStDate.Text & mainFrm.cmbxSechDTStHour
    edDateHour = mainFrm.txtSechDTEdDate.Text & mainFrm.cmbxSechDTEdHour
    
             strQry = "SELECT * FROM( " + vbCrLf
    strQry = strQry + "SELECT /*+INDEX(B,PK_LOG_DT) */ DT_TS_ID " + vbCrLf
    strQry = strQry + ", TS_NAME" + vbCrLf
    strQry = strQry + ", CASE WHEN LOG_ID='V900' THEN TO_CHAR(DT_TIME,'yyyy/mm/dd hh24:mi:ss') END DT_TIME" + vbCrLf
    strQry = strQry + ", TO_CHAR(REG_DATE,'yyyy/mm/dd hh24:mi:ss') REG_DATE" + vbCrLf
    strQry = strQry + ", LOG_CONTENT " + vbCrLf
    strQry = strQry + "FROM RTDB.TIDAL_STATION A, RTDB.LOG_DT B " + vbCrLf
    strQry = strQry + "WHERE A.TS_ID = B.DT_TS_ID " + vbCrLf
    If mainFrm.cmbxSechDTID.Text = "전체" Then
        strQry = strQry + "  AND DT_TS_ID  > 0 " + vbCrLf
    Else
        strQry = strQry + "  AND DT_TS_ID  = (SELECT TS_ID FROM RTDB.TIDAL_STATION WHERE TS_NAME = '" + mainFrm.cmbxSechDTID.Text + "') " + vbCrLf
    End If
    strQry = strQry + "  AND REG_DATE >= TO_DATE('" + stDateHour + "0000', 'YYYY-MM-DDHH24MISS') " + vbCrLf
    strQry = strQry + "  AND REG_DATE <= TO_DATE('" + edDateHour + "5959', 'YYYY-MM-DDHH24MISS') " + vbCrLf
    strQry = strQry + ")" + vbCrLf
    If Not mainFrm.cmbxDtRownum.Text = "ALL" Then
        strQry = strQry + "WHERE ROWNUM <= " + mainFrm.cmbxDtRownum.Text
    End If
    
    Call LogWrite("goDtSearch")
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
        Dim Cnt As Integer
        
        mainFrm.fpSpread_DtLog.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_DtLog.MaxRows = dtRs.RecordCount
    
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_DtLog.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next Cnt
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

Public Sub Init_fpSpread_DtLog()
    
    With mainFrm.fpSpread_DtLog
    
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

