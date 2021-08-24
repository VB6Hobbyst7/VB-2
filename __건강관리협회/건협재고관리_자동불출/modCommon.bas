Attribute VB_Name = "modCommon"
Option Explicit

Public gDbCn As ADODB.Connection, gSql As String, cDb As clsDbConnect
Public gUserId As String, gAreaCd As String, gCompany As String, gERPStkGroup As String, gERPStkCondition As String
Public gChangGoMng As Boolean, gWorkArea As Boolean
Public gKahpUser As String, gKahpUserTable As String
Public gHelpCode As String, gAutoEnter As Boolean

Public Const gMachCycleFgStr$ = "매일,매주,매월"
Public Const gMachCycleDayStr$ = "NONE,요일,일자"
Public Const gSpcStatusStr$ = "보관,대출,폐기"

Public gMachCycle() As String, gMachCycleDay() As String, gSpcStatus() As String

' ERP연동 관련 Table
Public Const gTBLstk$ = "MA_PITEM_KAHP_LAB@KAHP_ERP"
Public Const gTBLenter$ = "MM_QTIO_IN_KAHP_LAB@KAHP_ERP"
Public Const gTBLleave$ = "MM_QTIO_OUT_KAHP_LAB@KAHP_ERP"
Public Const gTBLPartner$ = "MA_PARTNER@KAHP_ERP"

Public Const gLockColor = &HE0E0E0, gEditColor = vbWhite
Public Const gGrpLineColor = vbWhite, gGrpBackColor = vbWhite
'Public Const gGrpLineColor = &H99A8AC, gGrpBackColor = vbWhite

Public Const gReasonCommon$ = "0", gReasonMach$ = "1", gReasonTest$ = "2", gReasonManual$ = "3", gReasonExpirydt$ = "4"

Public gDecimalMoney As Integer, gDecimalQtyI As Integer, gDecimalQtyO As Integer

Public gPrgBar As ProgressBar

' 마우스 좌표값 ---------------------------------
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Type POINTAPI
x As Long
y As Long
End Type
'------------------------------------------------
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
     ByVal lpFile As String, ByVal lpParameters As String, _
     ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub Main()
Dim sCmd() As String

    If App.PrevInstance Then
        MsgBox App.Title & "이 프로그램이 이미 실행중입니다.!", vbInformation
        End
    End If
    
    gMachCycle = Split(gMachCycleFgStr, ",")
    gMachCycleDay = Split(gMachCycleDayStr, ",")
    gSpcStatus = Split(gSpcStatusStr, ",")
    
    gUserId = "dev"
    gCompany = "1000"
    gERPStkGroup = "'004'"
    
    gDecimalMoney = 1
    gDecimalQtyI = 2
    gDecimalQtyO = 1
    
    ' 창고관리유무(창고관리 안할경우 입고와 동시에 불출처리)
    gChangGoMng = False
    ' 중앙센터(true)와 지부(false)의 구분
    gWorkArea = False
    gKahpUser = "TW_MIS_MED."
    gKahpUserTable = "TW_MIS_MED.TWMED_USER2013"
    
    frmMain.Show
    frmMain.Enabled = False
    
    sCmd = Split(Command$, ";")
    If UBound(sCmd) > 0 Then
        gUserId = sCmd(0)
        gSql = "SELECT A.EMPID AS USERID, A.EMPNM AS USER_NM, B.LOGINPASS AS PASSWORD " & vbNewLine & _
               "  FROM S2COM006 A INNER JOIN S2COM010 B ON A.EMPID=B.LOGINID WHERE A.EMPID='" & Trim(gUserId) & "'"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    frmMain.stsBar.Panels(3).Text = "" & .Fields("USER_NM").Value
                Else
                    MsgBox "등록되지 않은 사용자 입니다.!", vbCritical
                    End
                End If
                .Close
            End If
        End With
    Else
'        If gWorkArea = False Then
        frmMain.Enabled = False
        frmLogin.Show vbModal
'        End If
    End If
    frmMain.Enabled = True
    
    Call frmMain.psInitial
    
    Set gPrgBar = frmMain.prgBar

    gERPStkCondition = " AND X.CD_COMPANY='" & gCompany & "' AND X.CD_PLANT='" & gAreaCd & "'"
    
End Sub

Public Sub gsMousePoint(ByVal brForm As Form)
Dim sPnt As POINTAPI

    Call GetCursorPos(sPnt)
    
    brForm.Top = frmMain.Top + sPnt.y * 10
    brForm.Left = frmMain.Left + sPnt.x * 10
    
End Sub

Public Function HLeft(ByVal vString As String, ByVal vLen As Long) As String
' 한글포함된 문장에서 Left함수 사용
    HLeft = StrConv(LeftB(StrConv(vString, vbFromUnicode), vLen), vbUnicode)

End Function

Public Function HRight(ByVal vString As String, ByVal vLen As Long) As String
' 한글포함된 문장에서 right함수 사용

    HRight = StrConv(RightB(StrConv(vString, vbFromUnicode), vLen), vbUnicode)

End Function

Public Function HMid(ByVal vString As String, ByVal vLenF As Long, ByVal vLenT As Long) As String
' 한글포함된 문장에서 mid함수 사용

    HMid = StrConv(MidB(StrConv(vString, vbFromUnicode), vLenF, vLenT), vbUnicode)

End Function

Public Function HLen(ByVal vString As String) As Long
' 한글포함된 문장에서 len함수 사용

    HLen = LenB(StrConv(vString, vbFromUnicode))

End Function

Public Sub gsSpreadClear(ByVal brSpread As Object, Optional ByVal brRow As Long = 1000, Optional ByVal brColor As Boolean = False, Optional ByVal brHeight As Integer = 0, _
                         Optional ByVal brRowAdd As Boolean = False)
' 스프레드 Clear
    
   On Error GoTo gsSpreadClear_ERROR

    With brSpread
        .UserResize = UserResizeNone
        
        .MaxRows = brRow
        .RetainSelBlock = True
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
        If brRowAdd = False Then
            .Action = ActionClearText
        End If
        If brHeight = 0 Then
            .RowHeight(-1) = .FontSize * 1.5
        Else
            .RowHeight(-1) = brHeight
        End If
        .BlockMode = False
        
        If brColor Then
            .SetOddEvenRowColor vbWhite, vbBlack, &HF1F1F1, vbBlack
        End If
        .SetCellBorder 1, 1, .MaxCols, .MaxRows, 13, &HDEDEDE, CellBorderStyleSolid
        .SelectBlockOptions = SelectBlockOptionsAll
    End With

   Exit Sub
gsSpreadClear_ERROR:
   MsgBox Err.Numbe, vbCritical

End Sub

Public Sub gsButtonEnable(ByVal brBtn As Object, ByVal brOnOff As Boolean)

    If brOnOff Then
        brBtn.Enabled = True
'        brBtn.ForeColor = vbBlack
    Else
        brBtn.Enabled = False
'        brBtn.ForeColor = &HE0E0E0
    End If
    
End Sub

Public Function gfCurrencyStr(ByVal brVal As Currency) As String
Dim sStr As String
    
    sStr = "#,##0"
    If gDecimalMoney > 0 Then
        sStr = sStr & "." & String(gDecimalMoney, "0")
    End If
    gfCurrencyStr = Format(brVal, sStr)

End Function

Public Function gfQtyInputStr(ByVal brVal As Double) As String
Dim sStr As String
    
    sStr = "#,##0"
    If gDecimalQtyI > 0 Then
        sStr = sStr & "." & String(gDecimalQtyI, "0")
    End If
    gfQtyInputStr = Format(brVal, sStr)

End Function

Public Function gfQtyOutputStr(ByVal brVal As Double) As String
Dim sStr As String
    
    sStr = "#,##0"
    If gDecimalQtyO > 0 Then
        sStr = sStr & "." & String(gDecimalQtyO, "0")
    End If
    gfQtyOutputStr = Format(brVal, sStr)

End Function

Public Sub gsSpreadToExcel(ByVal brSpread As fpSpread, ByVal brTitle As String, _
                                  Optional brLineFg As Boolean = True)
Dim sExlAPP As Excel.Application, sExlBook As Excel.Workbook, sExlSheet As Excel.Worksheet
Dim sSpeadData()
Dim sValA As Long, sColStr As String
Dim sStrTmp As String, sColAp As Long, sCol As Long, sRow As Long, sFmtStr As String
Dim sMaxRows As Long, sMaxCols As Long, sHiddenCnt As Long, sHeadRow As Long

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    If brSpread.MaxRows < 1 Then Exit Sub
    '데이타 건수가 10000건이상일때는 파일로 받을지 묻는다
    If brSpread.MaxRows > 1000 Then
        If MsgBox("데이타가 1000건 이상입니다." & vbCrLf & vbCrLf & _
            "데이타가 많아 오래걸리거나 컴퓨터가 다운될 수 있습니다" & vbCrLf & vbCrLf & _
            "프로그램을 실행하지 않고 파일로 받으시겠습니까?", vbQuestion + vbYesNo, "파일받기") = vbYes Then
            Call gsp_SetSpdExcelFileExport(brSpread, brTitle)
            
            Exit Sub
        End If
    End If
    
    sMaxRows = brSpread.MaxRows
    sMaxCols = brSpread.MaxCols
    sHeadRow = brSpread.ColHeaderRows
    
    ReDim sSpeadData(sMaxRows + sHeadRow, sMaxCols)
    
    Set sExlAPP = CreateObject("Excel.Application")
    
    'Excel이 설치되어있지 않으면 에러가 발생한다
    If Err.Number > 0 Then
        MsgBox "Excel 프로그램을 실행 할 수가 없습니다." & vbCrLf & _
        "Excel 프로그램이 설치되어 있는지 확인바랍니다.", vbInformation, "프로그램 실행에러"
        
        Exit Sub
    End If
    
'    sExlAPP.Visible = True
    
    Set sExlBook = sExlAPP.Workbooks.Add
    Set sExlSheet = sExlBook.Worksheets("Sheet1")
    
    For sRow = 1 To sHeadRow
        sHiddenCnt = 0
        
        If sRow = 0 Then
            sColAp = 0
        Else
            sColAp = -1001 + sRow
        End If
        
        For sCol = 1 To sMaxCols
            brSpread.Row = sColAp
            brSpread.Col = sCol
            If brSpread.ColHidden = False Then
                sSpeadData(sRow - 1, sCol - sHiddenCnt - 1) = brSpread.Text
            Else
                sHiddenCnt = sHiddenCnt + 1
            End If
        Next
    Next
    
    sColAp = sHeadRow - 1
    
    For sRow = 1 To sMaxRows + 1
        sHiddenCnt = 0
        
        For sCol = 1 To sMaxCols
            brSpread.Row = sRow
            brSpread.Col = sCol
            ' Cell이 Hidden 아닌 자료만 전송
            If brSpread.ColHidden = False Then
                Select Case brSpread.CellType
                    Case CellTypeNumber, CellTypeCurrency
                        sSpeadData(sColAp + sRow, sCol - sHiddenCnt - 1) = brSpread.Text
                    Case Else
                        If IsNumeric(brSpread.Text) = True Then
                            ' 자료형이 숫자형 문자의 경우
                            sSpeadData(sColAp + sRow, sCol - sHiddenCnt - 1) = brSpread.Text
                        Else
                            sSpeadData(sColAp + sRow, sCol - sHiddenCnt - 1) = brSpread.Text
                        End If
                End Select
            Else
                sHiddenCnt = sHiddenCnt + 1
            End If
        Next
    Next
    
    sValA = IIf((brSpread.MaxCols - sHiddenCnt) Mod 26 = 0, (brSpread.MaxCols - sHiddenCnt) \ 26 - 1, (brSpread.MaxCols - sHiddenCnt) \ 26)
    If sValA > 0 Then
       sStrTmp = Chr(sValA + 64)
    Else
       sStrTmp = ""
    End If
    
    ' 엑셀의 MAXCOL에 해당하는 문자열
    sStrTmp = sStrTmp & Chr(IIf((brSpread.MaxCols - sHiddenCnt) Mod 26 = 0, 26, (brSpread.MaxCols - sHiddenCnt) Mod 26) + 64)

    sColStr = "A1:" & sStrTmp & sMaxRows + sHeadRow
    sExlSheet.Range(sColStr) = sSpeadData                           ' 자료복사
    sExlSheet.Range(sColStr).Font.Size = 9
    If brLineFg Then
        sExlSheet.Range(sColStr).Borders.LineStyle = xlContinuous   ' 외곽라인
    End If
    
    sColStr = "A1:" & sStrTmp & sHeadRow                            ' 항목타이틀
    sExlSheet.Range(sColStr).Font.Bold = False                      ' 타이틀 폰트
    sExlSheet.Range(sColStr).HorizontalAlignment = xlCenter         ' 타이틀 정렬
    sExlSheet.Range(sColStr).VerticalAlignment = xlCenter           ' 타이틀 정렬
    sExlSheet.Range(sColStr).Interior.Color = &HF1F1F1              ' 타이틀 Background Color
    
    brSpread.Row = 1
    For sCol = 1 To sMaxCols
        brSpread.Col = sCol
        If brSpread.ColHidden = False Then
            sColAp = sColAp + 1
            sColStr = Chr(64 + sColAp) & "1"
            ' Cell의 폭을 Spread와 동일하게
            sExlSheet.Range(sColStr).ColumnWidth = brSpread.ColWidth(sCol)
            
            If brSpread.CellType = CellTypeNumber Or brSpread.CellType = CellTypeCurrency Or IsNumeric(brSpread.Text) Then
                ' 데이터타입이 숫자인경우 ',' 삽입
                If brSpread.TypeNumberDecPlaces > 0 Then
                    ' Spread Type이 숫자일 경우 소수점자리 확인
                    sFmtStr = "#,##0." & String(brSpread.TypeNumberDecPlaces, "0")
                Else
                    sFmtStr = "#,##0"
                End If
                sColStr = Chr(64 + sColAp) & "2:" & Chr(64 + sColAp) & sMaxRows + sHeadRow
                sExlSheet.Range(sColStr).NumberFormat = sFmtStr
            ElseIf IsDate(brSpread.Text) Then
                sFmtStr = "yyyy-MM-dd"
                sColStr = Chr(64 + sColAp) & "2:" & Chr(64 + sColAp) & sMaxRows + sHeadRow
                sExlSheet.Range(sColStr).NumberFormat = sFmtStr
            End If
        End If
    Next sCol
    
    'sExlSheet.Columns.AutoFit
    
    ' 엑셀시트에 제목등록
    sExlSheet.Range("A1:" & sStrTmp & "3").Insert
    sExlSheet.Range("A2") = Trim(brTitle)
    
    With sExlSheet.Range("A2:" & sStrTmp & "2")
        .Merge
        .Font.Underline = True
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    sExlSheet.Columns.Font.Name = "굴림체"

    sExlAPP.Visible = True
    Screen.MousePointer = vbDefault

End Sub

Public Sub gsp_SetSpdExcelFileExport(ByVal brSpread As Object, ByVal brTitle As String)

    Dim sFile    As String
    Dim sLogPath     As String
    Dim sReturn As Boolean
    
    'File Name Set
    sFile = App.Path & "\" & brTitle & ".xls"
    
    'Log Name Set
    sLogPath = App.Path & "\" & "Exceldown.log"
    
    sReturn = brSpread.ExportToExcel(sFile, "Sheet1", sLogPath)
    
    If sReturn Then MsgBox "Saved to " & sFile, vbInformation, "Excel download"

End Sub


