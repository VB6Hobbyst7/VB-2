Attribute VB_Name = "modSpreadToExcel"
'================================================================================
' @(f)
'
' 기능      : 대상 스프레드의 내용을 Excel로 Open한다
'
' 리턴값    : 없음
'
' 인자      : ARG1 - Excel로 내려질 대상 스프레드
'             ARG2 - 옵션:Excel에 라인 표시 여부(기본:True(표시))
'
' 기능설명  : 대상 스프레드의 내용을 Excel로 Open한다
'
' 비고      : 1000건이상일시는 오래걸리기 때문에 파일로 저장할것인지 묻는다
'
'              변동일자     변경 내역      작성자
'             ========== ================= ======
' 변동사항  :
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
'    '데이타 건수가 10000건이상일때는 파일로 받을지 묻는다
'    If spd_Control.maxrows > 2000 Then
'        If MsgBox("데이타가 2000건 이상입니다." & vbCrLf & vbCrLf & _
'            "데이타가 많아 오래걸리거나 컴퓨터가 다운될 수 있습니다" & vbCrLf & vbCrLf & _
'            "프로그램을 실행하지 않고 파일로 받으시겠습니까?", vbQuestion + vbYesNo, "파일받기") = vbYes Then
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
'    'Excel이 설치되어있지 않으면 에러가 발생한다
'    If Err.Number > 0 Then
'        MsgBox " ▒ Excel 프로그램을 실행 할 수가 없습니다." & vbCrLf & _
'               "    Excel 프로그램이 설치되어 있는지 확인바랍니다.", vbInformation, App.Title
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
''        xlWorksheet.Columns.Font.Name = "굴림"
''        xlWorksheet.Columns.Font.Size = 9
'    End If
'
End Sub

'================================================================================
' @(f)
'
' 기능      : Spread의 Export기능으로 Excel로 Export
'
' 리턴값    : 없음
'
' 인자      : ARG1 - Excel로 내려질 대상 스프레드
'
' 기능설명  : 대상 스프레드의 내용을 Excel File로 생성한다
'
' 비고      :
'
'              변동일자     변경 내역      작성자
'             ========== ================= ======
' 변동사항  :
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



