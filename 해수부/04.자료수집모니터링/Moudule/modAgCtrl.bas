Attribute VB_Name = "modAgCtrl"
Option Explicit

Public Sub startAGView()
    
    mainFrm.chkTimer_Ag = 1

End Sub

Public Sub setAgStationID()

Dim Cnt As Integer

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
             strQry = "SELECT A.TYP_BUOY_ID " + vbCrLf
    strQry = strQry + "  FROM WRN.T_WRN_TYP_BUOY_ID A, (SELECT TYP_BUOY_ID  FROM WRN.T_WRN_TYP_BUOY GROUP BY TYP_BUOY_ID) B " + vbCrLf
    strQry = strQry + " Where A.TYP_BUOY_ID = B.TYP_BUOY_ID " + vbCrLf
    strQry = strQry + " ORDER BY TYP_BUOY_ID " + vbCrLf
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    mainFrm.cmbxSechAGID.AddItem "전체"

    If Not (dtRs.EOF Or dtRs.BOF) Then
        For Cnt = 0 To dtRs.RecordCount - 1
            mainFrm.cmbxSechAGID.AddItem dtRs.Fields("TYP_BUOY_ID")
            dtRs.MoveNext
        Next Cnt
    End If
    
    mainFrm.cmbxSechAGID.ListIndex = 0
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub setAgSechCondition()

    Dim Cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '검색 일자 설정
    mainFrm.txtSechAGStDate.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtSechAGEdDate.Text = Format(Now, "YYYY-MM-DD")
    
End Sub

Public Sub goAgSearch()
        
Dim Cnt As Integer
    
    If Not IsDate(mainFrm.txtSechAGStDate.Text) Then
        MsgBox "검색 범위 시작일자를 확인해주세요."
        mainFrm.txtSechAGStDate.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mainFrm.txtSechAGEdDate.Text) Then
        MsgBox "검색 범위 종료일시를 확인해주세요."
        mainFrm.txtSechAGEdDate.SetFocus
        Exit Sub
    End If
    
    If mainFrm.cmbxAgRownum.Text = "ALL" Then
    ElseIf Not IsNumeric(mainFrm.cmbxAgRownum.Text) Then
        MsgBox "출력건수는 숫자를 입력해주세요."
        mainFrm.cmbxAgRownum.SetFocus
        Exit Sub
    End If
    
    '검색창초기화
    Call Init_fpSpread_AgLog

On Error Resume Next

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    Dim stDate As String
    Dim edDate As String
    
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn
    
    stDate = mainFrm.txtSechAGStDate.Text
    edDate = mainFrm.txtSechAGEdDate.Text
    
    strQry = ""
    strQry = strQry + "SELECT * " + vbCrLf
    strQry = strQry + "  FROM ( " + vbCrLf
    strQry = strQry + "SELECT B.TYP_BUOY_ID " + vbCrLf
    strQry = strQry + "       , B.LOG_ID " + vbCrLf
    strQry = strQry + "       , B.POS_TIME " + vbCrLf
    strQry = strQry + "       , B.REG_DATE " + vbCrLf
    strQry = strQry + "       , B.LOG_CONTENT " + vbCrLf
    strQry = strQry + "  FROM WRN.LOG_MASTER A, WRN.LOG_TYP_BUOY B " + vbCrLf
    strQry = strQry + " WHERE A.LOG_ID = B.LOG_ID " + vbCrLf
    strQry = strQry + "   AND B.LOG_ID > ' ' " + vbCrLf
    If Not mainFrm.cmbxSechAGID.Text = "전체" Then
        strQry = strQry + "  AND B.TYP_BUOY_ID ='" + mainFrm.cmbxSechAGID.Text + "' " + vbCrLf
    End If
    strQry = strQry + "ORDER BY  REG_DATE DESC " + vbCrLf
    strQry = strQry + ") " + vbCrLf
    If Not mainFrm.cmbxAgRownum.Text = "ALL" Then
        strQry = strQry + "WHERE ROWNUM <= " + mainFrm.cmbxAgRownum.Text + vbCrLf
    End If
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '동기화대상자료 있는가? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
        mainFrm.fpSpread_AgLog.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_AgLog.MaxRows = dtRs.RecordCount
    
        For Cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_AgLog.SetText(j + 1, Cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next Cnt
    Else
        MsgBox "검색 결과가 없습니다."
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

Public Sub Init_fpSpread_AgLog()
    
    With mainFrm.fpSpread_AgLog
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
        
        .SetText 1, 0, "부이ID"
        .ColWidth(1) = 10
        '.SetText 2, 0, "관측소명"
        '.ColWidth(2) = 10
        .SetText 2, 0, "관측시간"
        .ColWidth(2) = 15
        .SetText 3, 0, "로그기록시간"
        .ColWidth(3) = 20
        .SetText 4, 0, "로그내용"
        .ColWidth(4) = 50
    End With
    
End Sub
