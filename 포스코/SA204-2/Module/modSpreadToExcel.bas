Attribute VB_Name = "modSpreadToExcel"
'================================================================================
' @(f)
'
' ���      : ��� ���������� ������ Excel�� Open�Ѵ�
'
' ���ϰ�    : ����
'
' ����      : ARG1 - Excel�� ������ ��� ��������
'             ARG2 - �ɼ�:Excel�� ���� ǥ�� ����(�⺻:True(ǥ��))
'
' ��ɼ���  : ��� ���������� ������ Excel�� Open�Ѵ�
'
' ���      : 1000���̻��Ͻô� �����ɸ��� ������ ���Ϸ� �����Ұ����� ���´�
'
'              ��������     ���� ����      �ۼ���
'             ========== ================= ======
' ��������  :
'================================================================================
Public Sub gsp_SetSpdTExcelExport(ByVal spd_Control As Object, _
                                  Optional bln_FlagLine As Boolean = True)

'    Dim XlAPP As Excel.Application
'    Dim xlWorkbook As Excel.Workbook
'    Dim xlWorksheet As Excel.Worksheet
'    Dim myrange()
'    Dim I As Long, j As Long, k As Long
'    Dim lng_a As Long
'    Dim lng_b As Long
'    Dim str_tmp As String
'
'    Dim lng_MaxRows As Long
'    Dim lng_MaxCols As Long
'    Dim lng_ColHiddenCount As Long
'
'    On Error Resume Next
'
'    If spd_Control.maxrows < 1 Then Exit Sub
'    '����Ÿ �Ǽ��� 10000���̻��϶��� ���Ϸ� ������ ���´�
'    If spd_Control.maxrows > 2000 Then
'        If MsgBox("����Ÿ�� 2000�� �̻��Դϴ�." & vbCrLf & vbCrLf & _
'            "����Ÿ�� ���� �����ɸ��ų� ��ǻ�Ͱ� �ٿ�� �� �ֽ��ϴ�" & vbCrLf & vbCrLf & _
'            "���α׷��� �������� �ʰ� ���Ϸ� �����ðڽ��ϱ�?", vbQuestion + vbYesNo, "���Ϲޱ�") = vbYes Then
'            gsp_SetSpdExcelFileExport spd_Control
'            Exit Sub
'        End If
'    End If
'
'    lng_MaxRows = spd_Control.maxrows
'    lng_MaxCols = spd_Control.MaxCols
'
'    ReDim myrange(lng_MaxRows + spd_Control.ColHeaderRows, lng_MaxCols)
'
'    Set XlAPP = CreateObject("Excel.Application")
'
'    'Excel�� ��ġ�Ǿ����� ������ ������ �߻��Ѵ�
'    If Err.Number > 0 Then
'        MsgBox " �� Excel ���α׷��� ���� �� ���� �����ϴ�." & vbCrLf & _
'               "    Excel ���α׷��� ��ġ�Ǿ� �ִ��� Ȯ�ιٶ��ϴ�.", vbInformation, App.Title
'
'        Exit Sub
'    End If
'
'    XlAPP.Visible = True
'
'    Set xlWorkbook = XlAPP.Workbooks.Add
'    Set xlWorksheet = xlWorkbook.Worksheets.Add
'
'
'    For I = 1 To spd_Control.ColHeaderRows
'        lng_ColHiddenCount = 0
'
'        If I = 0 Then
'            k = 0
'        Else
'            k = -1001 + I
'        End If
'
'        For j = 1 To lng_MaxCols
'            spd_Control.Row = k
'            spd_Control.Col = j
'            If spd_Control.ColHidden = False Then
'                myrange(I - 1, j - lng_ColHiddenCount - 1) = spd_Control.text
'            Else
'                lng_ColHiddenCount = lng_ColHiddenCount + 1
'            End If
'        Next
'    Next
'
'    k = spd_Control.ColHeaderRows - 1
'
'    For I = 1 To lng_MaxRows + 1
'        lng_ColHiddenCount = 0
'
'        For j = 1 To lng_MaxCols
'            spd_Control.Row = I
'            spd_Control.Col = j
'            If spd_Control.ColHidden = False Then
'
'                Select Case spd_Control.CellType
'                    Case CellTypeNumber, CellTypeCurrency
'
'                        myrange(k + I, j - lng_ColHiddenCount - 1) = spd_Control.text
'
'                    Case Else
'                        If IsNumeric(spd_Control.text) = True Then
'                            myrange(k + I, j - lng_ColHiddenCount - 1) = "'" & spd_Control.text
'                        Else
'                            myrange(k + I, j - lng_ColHiddenCount - 1) = spd_Control.text
'                        End If
'                End Select
''                myrange(k + i, j - lng_ColHiddenCount - 1) = spd_Control.Text
'            Else
'                lng_ColHiddenCount = lng_ColHiddenCount + 1
'            End If
'        Next
'    Next
'
'    lng_a = IIf((spd_Control.MaxCols - lng_ColHiddenCount) Mod 26 = 0, (spd_Control.MaxCols - lng_ColHiddenCount) \ 26 - 1, (spd_Control.MaxCols - lng_ColHiddenCount) \ 26)
'    If lng_a > 0 Then
'       str_tmp = Chr(lng_a + 64)
'    Else
'       str_tmp = ""
'    End If
'
'    str_tmp = str_tmp & Chr(IIf((spd_Control.MaxCols - lng_ColHiddenCount) Mod 26 = 0, 26, (spd_Control.MaxCols - lng_ColHiddenCount) Mod 26) + 64)
'
'    If bln_FlagLine Then
'
'        xlWorksheet.Range("A1:" & str_tmp & spd_Control.maxrows + spd_Control.ColHeaderRows) = myrange
'
'        xlWorksheet.Range("A1:" & str_tmp & spd_Control.ColHeaderRows).Font.Bold = True
'        xlWorksheet.Range("A1:" & str_tmp & spd_Control.ColHeaderRows).HorizontalAlignment = xlCenter
'        xlWorksheet.Range("A1:" & str_tmp & spd_Control.ColHeaderRows).VerticalAlignment = xlCenter
'
'
'        xlWorksheet.Range("A1:" & str_tmp & spd_Control.maxrows + spd_Control.ColHeaderRows).Borders.LineStyle = xlContinuous
'
'        xlWorksheet.Columns.AutoFit
'
'        xlWorksheet.Range("A1:" & str_tmp & "3").Insert
'
'        If spd_Control.ToolTipText = "" Then
'            If InStr(Replace(Trim(Screen.ActiveForm.Caption), ":::", ""), "[") > 0 Then
'                xlWorksheet.Range("A2") = Mid(Replace(Trim(Screen.ActiveForm.Caption), ":::", ""), 1, InStr(Replace(Trim(Screen.ActiveForm.Caption), ":::", ""), "[") - 1)
'            Else
'                xlWorksheet.Range("A2") = Trim(Screen.ActiveForm.Caption)
'            End If
'        Else
'            xlWorksheet.Range("A2") = Trim(spd_Control.ToolTipText)
'        End If
'
'        With xlWorksheet.Range("A2:" & str_tmp & "2")
'            .Merge
'            .Font.Underline = True
'            .Font.Bold = True
'            .Font.Size = 16
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'        End With
'
''        xlWorksheet.Columns.Font.Name = "����"
''        xlWorksheet.Columns.Font.Size = 9
'    End If
'
End Sub

'================================================================================
' @(f)
'
' ���      : Spread�� Export������� Excel�� Export
'
' ���ϰ�    : ����
'
' ����      : ARG1 - Excel�� ������ ��� ��������
'
' ��ɼ���  : ��� ���������� ������ Excel File�� �����Ѵ�
'
' ���      :
'
'              ��������     ���� ����      �ۼ���
'             ========== ================= ======
' ��������  :
'================================================================================
Public Sub gsp_SetSpdExcelFileExport(ByVal spd_Control As Object)

    Dim filePath    As String
    Dim logPath     As String
    Dim ret
    
    'File Name Set
    filePath = App.Path & "\" & Format(Now, "YYYYMMDD-HHMMSS") & ".xls"
    
    'Log Name Set
    logPath = App.Path & "\" & "Exceldown.log"
    
    ret = spd_Control.ExportToExcel(filePath, "Sheet1", logPath)
    
    If ret Then MsgBox "Saved to " & filePath, vbInformation, "Excel download"

End Sub



