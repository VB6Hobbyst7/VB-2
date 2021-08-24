Attribute VB_Name = "modRtCtrl"
Option Explicit

Public Sub startRTView()
    
    '�˻�â�� �ð� ���� ����
    'setRtSechCondition
    
    mainFrm.chkTimer_Rt = 1
    '�˻�
    'goRtSearch
End Sub

Public Sub setRtStationID()

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
'LogWrite ("setRtStationID")
    strQry = ""
    strQry = strQry + "    SELECT NAME, B.SID " + vbCrLf
    strQry = strQry + "    FROM REALTIME.STATION A, (SELECT DISTINCT SID " + vbCrLf
    strQry = strQry + "                     From REALTIME.LOG_REALTIME_DATA " + vbCrLf
    strQry = strQry + "                    ) B " + vbCrLf
    strQry = strQry + "   Where A.SID = B.SID " + vbCrLf
    strQry = strQry + "     AND FILENAME IS NOT NULL " + vbCrLf
    strQry = strQry + "   ORDER BY NAME " + vbCrLf
    
'LogWrite (strQry)
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn
    
    mainFrm.cmbxSechRTID.AddItem "��ü"
    'mainFrm.CboRTID_NM.AddItem "��ü"
    
    If Not (dtRs.EOF Or dtRs.BOF) Then

        Dim cnt As Integer

        For cnt = 0 To dtRs.RecordCount - 1
            mainFrm.cmbxSechRTID.AddItem dtRs.Fields("NAME")
            mainFrm.CboRTID_NM.AddItem dtRs.Fields("NAME")
            mainFrm.CboRTID_ID.AddItem dtRs.Fields("SID")

            dtRs.MoveNext
        Next cnt
        
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

    Dim cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '�˻� ���� ����
    mainFrm.txtSechRTStDate.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtSechRTEdDate.Text = Format(Now, "YYYY-MM-DD")
    
    
End Sub

'Public Sub chkRtSechCondition()
'
'    If Not IsDate(mainFrm.txtSechRTStDate.Text) Then
'        MsgBox "�˻� ���� �������ڸ� Ȯ�����ּ���."
'        mainFrm.txtSechRTStDate.SetFocus
'        Exit Sub
'    End If
'
'    If Not IsDate(mainFrm.txtSechRTEdDate.Text) Then
'        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
'        mainFrm.txtSechRTEdDate.SetFocus
'        Exit Sub
'    End If
'
'    If mainFrm.cmbxRtRownum.Text = "ALL" Then
'    ElseIf Not IsNumeric(mainFrm.cmbxRtRownum.Text) Then
'        MsgBox "��°Ǽ��� ���ڸ� �Է����ּ���."
'        mainFrm.cmbxRtRownum.SetFocus
'        Exit Sub
'    End If
'End Sub

Public Sub goRtSearch()
    '�˻����� üũ st
    If Not IsDate(mainFrm.txtSechRTStDate.Text) Then
        MsgBox "�˻� ���� �������ڸ� Ȯ�����ּ���."
        mainFrm.txtSechRTStDate.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mainFrm.txtSechRTEdDate.Text) Then
        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
        mainFrm.txtSechRTEdDate.SetFocus
        Exit Sub
    End If
    
    If mainFrm.cmbxRtRownum.Text = "ALL" Then
    ElseIf Not IsNumeric(mainFrm.cmbxRtRownum.Text) Then
        MsgBox "��°Ǽ��� ���ڸ� �Է����ּ���."
        mainFrm.cmbxRtRownum.SetFocus
        Exit Sub
    End If
    '�˻����� üũ end
    '�˻�â �ʱ�ȭ
    Init_fpSpread_RtLog

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    Dim stDate As String
    Dim edDate As String
    
    'DB����
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn
    
    stDate = mainFrm.txtSechRTStDate.Text
    edDate = mainFrm.txtSechRTEdDate.Text
    
    strQry = ""
    

    strQry = strQry + "SELECT * FROM( " + vbCrLf
    strQry = strQry + "SELECT  B.SID " + vbCrLf
    strQry = strQry + "       , NAME " + vbCrLf
    strQry = strQry + "       , CASE WHEN LOG_ID='L006' Then TO_CHAR(DTIME,'yyyy/mm/dd hh24:mi:ss') End DT_TIME " + vbCrLf
'    strQry = strQry + "       , LOG_ID " + vbCrLf
    strQry = strQry + "       , TO_CHAR(REG_DATE,'yyyy/mm/dd hh24:mi:ss') REG_DATE " + vbCrLf
    strQry = strQry + "       , LOG_CONTENT  " + vbCrLf
    strQry = strQry + "FROM REALTIME.STATION A, REALTIME.LOG_REALTIME_DATA B  " + vbCrLf
    strQry = strQry + "WHERE A.SID = B.SID  " + vbCrLf
    strQry = strQry + "  AND B.SID  > 0  " + vbCrLf
    If Not mainFrm.cmbxSechRTID.Text = "��ü" Then
        strQry = strQry + "  AND B.SID IN ('0', (SELECT SID FROM REALTIME.STATION WHERE NAME = '" + mainFrm.cmbxSechRTID.Text + "')) " + vbCrLf
    End If
    strQry = strQry + "  AND REG_DATE >= TO_DATE('" + stDate + "000000', 'YYYY-MM-DDHH24MISS')  " + vbCrLf
    strQry = strQry + "  AND REG_DATE <= TO_DATE('" + edDate + "235959', 'YYYY-MM-DDHH24MISS')  " + vbCrLf
    strQry = strQry + "ORDER BY REG_DATE DESC, SID DESC, LOG_ID DESC " + vbCrLf
    strQry = strQry + ") " + vbCrLf
    If Not mainFrm.cmbxRtRownum.Text = "ALL" Then
        strQry = strQry + "WHERE ROWNUM <= " + mainFrm.cmbxRtRownum.Text
    End If
    
    
    
'LogWrite strQry

    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '����ȭ����ڷ� �ִ°�? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
'        If AdoDBConn.Errors.Count = 0 Then
'            mainFrm.fpSpread_RtLog.DataSource = dtRs.DataSource
'        Else
'            AdoDBConn.Errors.Clear
'        End If
        Dim cnt As Integer
        
        mainFrm.fpSpread_RtLog.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_RtLog.MaxRows = dtRs.RecordCount
        
    
        For cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_RtLog.SetText(j + 1, cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next cnt
        
        
        '����ó��
'        For cnt = 0 To mainFrm.fpSpread_RtLog.MaxRows -1
'            mainFrm.fpSpread_RtLog.Col = 2
'            mainFrm.fpSpread_RtLog.Row = cnt + 1
'
'            If DateDiff("d", CDate(mainFrm.fpSpread_RtLog.Text), Now) >= 1 Then
'                '-- ���� ������ ����
'                For j = 0 To mainFrm.fpSpread_RtLog.MaxCols -1
'                    mainFrm.fpSpread_RtLog.Col = j + 1
'                    mainFrm.fpSpread_RtLog.Row = cnt + 1
'                    mainFrm.fpSpread_RtLog.BackColor = vbRed
'                    mainFrm.fpSpread_RtLog.ForeColor = vbWhite
'                Next j
'
'                mainFrm.fpSpread_RtLog.Col = 2
'                mainFrm.fpSpread_RtLog.Row = cnt + 1
'            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_RtLog.Text), Now) >= strJowiVPNCautionMin Then
'                    '-- ���� ������ ����
'                    For j = 0 To mainFrm.fpSpread_RtLog.MaxCols -1
'                        mainFrm.fpSpread_RtLog.Col = j + 1
'                        mainFrm.fpSpread_RtLog.Row = cnt + 1
'                        mainFrm.fpSpread_RtLog.BackColor = vbYellow
'                        mainFrm.fpSpread_RtLog.ForeColor = vbBlack
'                    Next j
'
'                    mainFrm.fpSpread_RtLog.Col = 2
'                    mainFrm.fpSpread_RtLog.Row = cnt + 1
'                    'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
'            Else
'                For j = 0 To mainFrm.fpSpread_RtLog.MaxCols -1
'                    mainFrm.fpSpread_RtLog.Col = j + 1
'                    mainFrm.fpSpread_RtLog.Row = cnt + 1
'                    mainFrm.fpSpread_RtLog.BackColor = vbWhite
'                    mainFrm.fpSpread_RtLog.ForeColor = vbBlack
'                Next j
'
'            End If
'        Next cnt
        
    
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    
    'DB��������
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
        
        .SetText 1, 0, "������ID"
        .ColWidth(1) = 10
        .SetText 2, 0, "�����Ҹ�"
        .ColWidth(2) = 10
        .SetText 3, 0, "�����ð�"
        .ColWidth(3) = 15
        .SetText 4, 0, "�αױ�Ͻð�"
        .ColWidth(4) = 15
        .SetText 5, 0, "�α׳���"
        .ColWidth(5) = 45
    End With
    
End Sub

