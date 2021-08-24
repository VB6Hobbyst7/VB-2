Attribute VB_Name = "modUsnCtrl"
Option Explicit

Public Sub startUSNView()
    
'    '�˻�â�� �ð� ���� ����
'    Call setUsnSechCondition
    
    mainFrm.chkTimer_Usn = 1
    '�˻�
    'goUsnSearch
End Sub

Public Sub setUsnStationID()

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    
    strQry = "SELECT /*+ INDEX(TK_USN_STATION_CONFIG INX_TK_USN_STATION_CONFIG_01) */ STATION_NAME " + vbCrLf
    strQry = strQry + " FROM USN.TK_USN_STATION_CONFIG " + vbCrLf
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    mainFrm.cmbxSechUSNID.AddItem "��ü"

    If Not (dtRs.EOF Or dtRs.BOF) Then

        Dim cnt As Integer

        For cnt = 0 To dtRs.RecordCount - 1
            mainFrm.cmbxSechUSNID.AddItem dtRs.Fields("STATION_NAME")
            dtRs.MoveNext
        Next cnt
    
        mainFrm.cmbxSechUSNID.ListIndex = 0
    End If
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("ERR : " & Err.Number & "-" & Err.Description)
    End If
End Sub

Public Sub setUsnSechCondition()

    Dim cnt As Integer
    Dim strHour As String
    Dim strNowHour As String
    Dim strNowIdx As Integer
    strNowHour = Format(Now, "hh")
    '�˻� ���� ����
    mainFrm.txtSechUSNStDate.Text = Format(Now, "YYYY-MM-DD")
    mainFrm.txtSechUSNEdDate.Text = Format(Now, "YYYY-MM-DD")
    
End Sub

'Public Sub chkUsnSechCondition()
'
'    If Not IsDate(mainFrm.txtSechUSNStDate.Text) Then
'        MsgBox "�˻� ���� �������ڸ� Ȯ�����ּ���."
'        mainFrm.txtSechUSNStDate.SetFocus
'        Exit Sub
'    End If
'
'    If Not IsDate(mainFrm.txtSechUSNEdDate.Text) Then
'        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
'        mainFrm.txtSechUSNEdDate.SetFocus
'        Exit Sub
'
'    End If
'
'    If mainFrm.cmbxUsnRownum.Text = "ALL" Then
'    ElseIf Not IsNumeric(mainFrm.cmbxUsnRownum.Text) Then
'        MsgBox "��°Ǽ��� ���ڸ� �Է����ּ���."
'        mainFrm.cmbxUsnRownum.SetFocus
'        Exit Sub
'    End If
'End Sub

Public Sub goUsnSearch()
    '�˻����� üũ st
    If Not IsDate(mainFrm.txtSechUSNStDate.Text) Then
        MsgBox "�˻� ���� �������ڸ� Ȯ�����ּ���."
        mainFrm.txtSechUSNStDate.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mainFrm.txtSechUSNEdDate.Text) Then
        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
        mainFrm.txtSechUSNEdDate.SetFocus
        Exit Sub

    End If
    
    If mainFrm.cmbxUsnRownum.Text = "ALL" Then
    ElseIf Not IsNumeric(mainFrm.cmbxUsnRownum.Text) Then
        MsgBox "��°Ǽ��� ���ڸ� �Է����ּ���."
        mainFrm.cmbxUsnRownum.SetFocus
        Exit Sub
    End If
    '�˻����� üũ end
    '�˻�â�ʱ�ȭ
    Init_fpSpread_UsnLog

On Error Resume Next
'CCY On Error GoTo ErrorHandler

    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    Dim stDate As String
    Dim edDate As String
    
    'DB����
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn
    
    stDate = mainFrm.txtSechUSNStDate.Text
    edDate = mainFrm.txtSechUSNEdDate.Text
    
    strQry = ""
    
    
    strQry = strQry + "SELECT * " + vbCrLf
    strQry = strQry + "  FROM ( " + vbCrLf
    strQry = strQry + "       SELECT /*+ INDEX_DESC(A PK_LOG_TK_USN_OBSERVATION) */ CASE WHEN LOG_ID=1000 THEN STATION_ID END STATION_ID " + vbCrLf
    'strQry = strQry + "       SELECT  CASE WHEN LOG_ID=1000 THEN STATION_ID END STATION_ID " + vbCrLf
    strQry = strQry + "          , CASE WHEN LOG_ID=1000 THEN (SELECT STATION_NAME FROM USN.TK_USN_STATION_CONFIG WHERE STATION_ID= A.STATION_ID) END STN_NM " + vbCrLf
    strQry = strQry + "          , CASE WHEN LOG_ID=1000 THEN TO_CHAR(OBS_TIME,'yyyy/mm/dd hh24:mi:ss') END OBS_TIME " + vbCrLf
    strQry = strQry + "          , TO_CHAR(REG_DATE,'yyyy/mm/dd hh24:mi:ss') REG_DATE " + vbCrLf
    strQry = strQry + "          , LOG_CONTENT " + vbCrLf
    strQry = strQry + "       FROM USN.LOG_TK_USN_OBSERVATION A " + vbCrLf
    strQry = strQry + "      WHERE LOG_ID > -1 " + vbCrLf
    If Not mainFrm.cmbxSechUSNID.Text = "��ü" Then
        strQry = strQry + "  AND STATION_ID IN ('000', (SELECT STATION_ID FROM USN.TK_USN_STATION_CONFIG WHERE STATION_NAME = '" + mainFrm.cmbxSechUSNID.Text + "') ) " + vbCrLf
    End If
    strQry = strQry + "  AND REG_DATE >= TO_DATE('" + mainFrm.txtSechUSNStDate.Text + "000000', 'YYYY-MM-DDHH24MISS') " + vbCrLf
    strQry = strQry + "  AND REG_DATE <= TO_DATE('" + mainFrm.txtSechUSNEdDate.Text + "235959', 'YYYY-MM-DDHH24MISS') " + vbCrLf
    'strQry = strQry + " ORDER BY REG_DATE DESC, LOG_ID DESC " + vbCrLf
    strQry = strQry + "  ) " + vbCrLf
    If Not mainFrm.cmbxUsnRownum.Text = "ALL" Then
        strQry = strQry + "WHERE ROWNUM <= " + mainFrm.cmbxUsnRownum.Text + vbCrLf
    End If
    
   
'LogWrite "goUsnSearch strQry=" & strQry

    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    '����ȭ����ڷ� �ִ°�? st
    If Not (dtRs.EOF Or dtRs.BOF) Then

        Dim cnt As Integer
        
        mainFrm.fpSpread_UsnLog.MaxCols = dtRs.Fields.Count
        mainFrm.fpSpread_UsnLog.MaxRows = dtRs.RecordCount
        
    
        For cnt = 0 To dtRs.RecordCount - 1
            Dim j As Integer
            For j = 0 To dtRs.Fields.Count - 1
                Call mainFrm.fpSpread_UsnLog.SetText(j + 1, cnt + 1, dtRs(j))
            Next j
            dtRs.MoveNext
        Next cnt
        
    Else
        MsgBox "�˻� ����� �����ϴ�."
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

Public Sub Init_fpSpread_UsnLog()
    
    With mainFrm.fpSpread_UsnLog

    
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
        
        
        .RowHeight(0) = 15
        
        .SetText 1, 0, "����ID"
        .ColWidth(1) = 10
        .SetText 2, 0, "�����Ҹ�"
        .ColWidth(2) = 10
        .SetText 3, 0, "�����ð�"
        .ColWidth(3) = 15
        .SetText 4, 0, "�αױ�Ͻð�"
        .ColWidth(4) = 20
        .SetText 5, 0, "�α׳���"
        .ColWidth(5) = 50
    End With
    
End Sub



