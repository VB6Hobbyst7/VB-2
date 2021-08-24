Attribute VB_Name = "modDtCtrl"
Option Explicit

Public Sub startDTView()
    
    '�˻�â�� �ð� ���� ����
    Call setDtSechCondition
    
    mainFrm.chkTimer_Dt = 1
    '�˻�
    'goDtSearch
End Sub

Public Sub setDtStationID()

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    strQry = "SELECT TS_NAME  " + vbCrLf
    strQry = strQry + "FROM RTDB.TIDAL_STATION " + vbCrLf
    strQry = strQry + "WHERE NOT VPN_IP IS NULL " + vbCrLf
    strQry = strQry + "ORDER BY TS_NAME ASC " + vbCrLf

    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    mainFrm.cmbxSechDTID.AddItem "��ü"

    If Not (dtRs.EOF Or dtRs.BOF) Then

        Dim cnt As Integer

        For cnt = 0 To dtRs.RecordCount - 1
            mainFrm.cmbxSechDTID.AddItem dtRs.Fields("TS_NAME")
            dtRs.MoveNext
        Next cnt
    
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

    Dim cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '�˻� ���� ����
    mainFrm.txtSechDTStDate.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtSechDTEdDate.Text = Format(Now, "YYYY-MM-DD")
    '�˻� �ð� ����
    For cnt = 0 To 23
        If cnt < 10 Then
            strHour = "0" & cnt
        Else
            strHour = cnt
        End If
        
        If strHour = strNowHour Then
             strNowIdx = cnt
        End If
        mainFrm.cmbxSechDTStHour.AddItem strHour
        mainFrm.cmbxSechDTEdHour.AddItem strHour
    Next cnt
    
    mainFrm.cmbxSechDTEdHour.ListIndex = strNowIdx
    If strNowIdx = 0 Then
        mainFrm.cmbxSechDTStHour.ListIndex = strNowIdx
    Else
        mainFrm.cmbxSechDTStHour.ListIndex = strNowIdx - 1
    End If
    
End Sub

Public Sub Sub_SetTwDate()

    Dim cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '�˻� ���� ����
    mainFrm.txtTwDate_From.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtTwDate_To.Text = Format(Now, "YYYY-MM-DD")
    '�˻� �ð� ����
    For cnt = 0 To 23
        If cnt < 10 Then
            strHour = "0" & cnt
        Else
            strHour = cnt
        End If
        
        If strHour = strNowHour Then
             strNowIdx = cnt
        End If
        mainFrm.cboTwhh_From.AddItem strHour
        mainFrm.cboTwhh_To.AddItem strHour
    Next cnt
    
    mainFrm.cboTwhh_To.ListIndex = strNowIdx
    If strNowIdx = 0 Then
        mainFrm.cboTwhh_From.ListIndex = strNowIdx
    Else
        mainFrm.cboTwhh_From.ListIndex = strNowIdx - 1
    End If
    
End Sub

Public Sub Sub_SetRTIDDate()

    Dim cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '�˻� ���� ����
    mainFrm.txtRTIDDate_From.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtRTIDDate_To.Text = Format(Now, "YYYY-MM-DD")
    '�˻� �ð� ����
    For cnt = 0 To 23
        If cnt < 10 Then
            strHour = "0" & cnt
        Else
            strHour = cnt
        End If
        
        If strHour = strNowHour Then
             strNowIdx = cnt
        End If
        mainFrm.cboRTIDhh_From.AddItem strHour
        mainFrm.cboRTIDhh_To.AddItem strHour
    Next cnt
    
    mainFrm.cboRTIDhh_To.ListIndex = strNowIdx
    If strNowIdx = 0 Then
        mainFrm.cboRTIDhh_From.ListIndex = strNowIdx
    Else
        mainFrm.cboRTIDhh_From.ListIndex = strNowIdx - 1
    End If
    
End Sub

Public Sub Sub_Tw()

    Dim cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '�˻� ���� ����
    mainFrm.txtTwDate_From.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtTwDate_To.Text = Format(Now, "YYYY-MM-DD")
    '�˻� �ð� ����
    For cnt = 0 To 23
        If cnt < 10 Then
            strHour = "0" & cnt
        Else
            strHour = cnt
        End If
        
        If strHour = strNowHour Then
             strNowIdx = cnt
        End If
        mainFrm.cboTwhh_From.AddItem strHour
        mainFrm.cboTwhh_To.AddItem strHour
    Next cnt
    
    mainFrm.cboTwhh_To.ListIndex = strNowIdx
    If strNowIdx = 0 Then
        mainFrm.cboTwhh_From.ListIndex = strNowIdx
    Else
        mainFrm.cboTwhh_From.ListIndex = strNowIdx - 1
    End If
    
End Sub


'Public Sub chkDtSechCondition()
'
'    If Not IsDate(mainFrm.txtSechDTStDate.Text) Then
'        MsgBox "�˻� ���� �������ڸ� Ȯ�����ּ���."
'        mainFrm.txtSechDTStDate.SetFocus
'        Exit Sub
'    ElseIf IsNumeric(mainFrm.cmbxSechDTStHour.Text) = False Then
'        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
'        mainFrm.cmbxSechDTStHour.SetFocus
'        Exit Sub
'    Else
'        If mainFrm.cmbxSechDTStHour.Text > 24 Then
'            MsgBox "�˻� ���� �����Ͻô� ���ڸ� �Է����ּ���."
'            mainFrm.cmbxSechDTStHour.SetFocus
'            Exit Sub
'        End If
'    End If
'
'    If Not IsDate(mainFrm.txtSechDTEdDate.Text) Then
'        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
'        mainFrm.txtSechDTEdDate.SetFocus
'        Exit Sub
'    ElseIf IsNumeric(mainFrm.cmbxSechDTEdHour.Text) = False Then
'        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
'        mainFrm.cmbxSechDTEdHour.SetFocus
'        Exit Sub
'    Else
'        If mainFrm.cmbxSechDTEdHour.Text > 24 Then
'            MsgBox "�˻� ���� �����Ͻô� ���ڸ� �Է����ּ���."
'            mainFrm.cmbxSechDTEdHour.SetFocus
'            Exit Sub
'        End If
'    End If
'
'    If mainFrm.cmbxDtRownum.Text = "ALL" Then
'    ElseIf Not IsNumeric(mainFrm.cmbxDtRownum.Text) Then
'        MsgBox "��°Ǽ��� ���ڸ� �Է����ּ���."
'        mainFrm.cmbxDtRownum.SetFocus
'        Exit Sub
'    End If
'End Sub

Public Sub goDtSearch()
    '�˻����� üũ st
    If Not IsDate(mainFrm.txtSechDTStDate.Text) Then
        MsgBox "�˻� ���� �������ڸ� Ȯ�����ּ���."
        mainFrm.txtSechDTStDate.SetFocus
        Exit Sub
    ElseIf IsNumeric(mainFrm.cmbxSechDTStHour.Text) = False Then
        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
        mainFrm.cmbxSechDTStHour.SetFocus
        Exit Sub
    Else
        If mainFrm.cmbxSechDTStHour.Text > 24 Then
            MsgBox "�˻� ���� �����Ͻô� ���ڸ� �Է����ּ���."
            mainFrm.cmbxSechDTStHour.SetFocus
            Exit Sub
        End If
    End If
    
    If Not IsDate(mainFrm.txtSechDTEdDate.Text) Then
        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
        mainFrm.txtSechDTEdDate.SetFocus
        Exit Sub
    ElseIf IsNumeric(mainFrm.cmbxSechDTEdHour.Text) = False Then
        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
        mainFrm.cmbxSechDTEdHour.SetFocus
        Exit Sub
    Else
        If mainFrm.cmbxSechDTEdHour.Text > 24 Then
            MsgBox "�˻� ���� �����Ͻô� ���ڸ� �Է����ּ���."
            mainFrm.cmbxSechDTEdHour.SetFocus
            Exit Sub
        End If
    End If
    
    If mainFrm.cmbxDtRownum.Text = "ALL" Then
    ElseIf Not IsNumeric(mainFrm.cmbxDtRownum.Text) Then
        MsgBox "��°Ǽ��� ���ڸ� �Է����ּ���."
        mainFrm.cmbxDtRownum.SetFocus
        Exit Sub
    End If
    '�˻����� üũ end
    
    '�˻�â�ʱ�ȭ
    Call Init_fpSpread_DtLog

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    Dim stDateHour As String
    Dim edDateHour As String
    
    'DB����
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn
    
    stDateHour = mainFrm.txtSechDTStDate.Text & mainFrm.cmbxSechDTStHour
    edDateHour = mainFrm.txtSechDTEdDate.Text & mainFrm.cmbxSechDTEdHour
    
'    strQry = "SELECT ROWNUM, TS_NAME, DT_TIME, REG_DATE, LOG_CONTENT FROM( " + vbCrLf
    strQry = "SELECT * FROM( " + vbCrLf
    strQry = strQry + "SELECT /*+INDEX(B,PK_LOG_DT) */ DT_TS_ID " + vbCrLf
    strQry = strQry + ", TS_NAME" + vbCrLf
    strQry = strQry + ", CASE WHEN LOG_ID='V900' THEN TO_CHAR(DT_TIME,'yyyy/mm/dd hh24:mi:ss') END DT_TIME" + vbCrLf
    'strQry = strQry + ", LOG_ID" + vbCrLf
    strQry = strQry + ", TO_CHAR(REG_DATE,'yyyy/mm/dd hh24:mi:ss') REG_DATE" + vbCrLf
    strQry = strQry + ", LOG_CONTENT " + vbCrLf
    strQry = strQry + "FROM RTDB.TIDAL_STATION A, RTDB.LOG_DT B " + vbCrLf
    strQry = strQry + "WHERE A.TS_ID = B.DT_TS_ID " + vbCrLf
    If mainFrm.cmbxSechDTID.Text = "��ü" Then
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
    
LogWrite "goDtSearch"
'LogWrite strQry
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '����ȭ����ڷ� �ִ°�? st
    If Not (dtRs.EOF Or dtRs.BOF) Then
'        If AdoDBConn.Errors.Count = 0 Then
'            mainFrm.fpSpread_DtLog.DataSource = dtRs.DataSource
'        Else
'            AdoDBConn.Errors.Clear
'        End If
        Dim cnt As Integer
        
        mainFrm.fpSpread_DtLog.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_DtLog.MaxRows = dtRs.RecordCount
        
    
        For cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_DtLog.SetText(j + 1, cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next cnt
        
        
        '����ó��
'        For cnt = 0 To mainFrm.fpSpread_DtLog.MaxRows - 1
'            mainFrm.fpSpread_DtLog.Col = 2
'            mainFrm.fpSpread_DtLog.Row = cnt + 1
'
'            If DateDiff("d", CDate(mainFrm.fpSpread_DtLog.Text), Now) >= 1 Then
'                '-- ���� ������ ����
'                For j = 0 To mainFrm.fpSpread_DtLog.MaxCols - 1
'                    mainFrm.fpSpread_DtLog.Col = j + 1
'                    mainFrm.fpSpread_DtLog.Row = cnt + 1
'                    mainFrm.fpSpread_DtLog.BackColor = vbRed
'                    mainFrm.fpSpread_DtLog.ForeColor = vbWhite
'                Next j
'
'                mainFrm.fpSpread_DtLog.Col = 2
'                mainFrm.fpSpread_DtLog.Row = cnt + 1
'            ElseIf DateDiff("n", CDate(mainFrm.fpSpread_DtLog.Text), Now) >= strJowiVPNCautionMin Then
'                    '-- ���� ������ ����
'                    For j = 0 To mainFrm.fpSpread_DtLog.MaxCols - 1
'                        mainFrm.fpSpread_DtLog.Col = j + 1
'                        mainFrm.fpSpread_DtLog.Row = cnt + 1
'                        mainFrm.fpSpread_DtLog.BackColor = vbYellow
'                        mainFrm.fpSpread_DtLog.ForeColor = vbBlack
'                    Next j
'
'                    mainFrm.fpSpread_DtLog.Col = 2
'                    mainFrm.fpSpread_DtLog.Row = cnt + 1
'                    'TS_Status_Panel.Caption = TS_Status_Panel.Caption + " " + fpSpread1.Text + ","
'            Else
'                For j = 0 To mainFrm.fpSpread_DtLog.MaxCols - 1
'                    mainFrm.fpSpread_DtLog.Col = j + 1
'                    mainFrm.fpSpread_DtLog.Row = cnt + 1
'                    mainFrm.fpSpread_DtLog.BackColor = vbWhite
'                    mainFrm.fpSpread_DtLog.ForeColor = vbBlack
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

