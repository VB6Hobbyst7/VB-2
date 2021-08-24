Attribute VB_Name = "modCommon"
Option Explicit

Public gDbCn As ADODB.Connection, gSql As String, cDb As clsDbConnect
Public gUserId As String, gAreaCd As String, gCompany As String, gERPStkGroup As String, gERPStkCondition As String
Public gChangGoMng As Boolean, gWorkArea As Boolean
Public gKahpUser As String, gKahpUserTable As String
Public gHelpCode As String, gAutoEnter As Boolean

Public Const gMachCycleFgStr$ = "����,����,�ſ�"
Public Const gMachCycleDayStr$ = "NONE,����,����"
Public Const gSpcStatusStr$ = "����,����,���"

Public gMachCycle() As String, gMachCycleDay() As String, gSpcStatus() As String

' ERP���� ���� Table
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

' ���콺 ��ǥ�� ---------------------------------
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
        MsgBox App.Title & "�� ���α׷��� �̹� �������Դϴ�.!", vbInformation
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
    
    ' â���������(â����� ���Ұ�� �԰�� ���ÿ� ����ó��)
    gChangGoMng = False
    ' �߾Ӽ���(true)�� ����(false)�� ����
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
                    MsgBox "��ϵ��� ���� ����� �Դϴ�.!", vbCritical
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
' �ѱ����Ե� ���忡�� Left�Լ� ���
    HLeft = StrConv(LeftB(StrConv(vString, vbFromUnicode), vLen), vbUnicode)

End Function

Public Function HRight(ByVal vString As String, ByVal vLen As Long) As String
' �ѱ����Ե� ���忡�� right�Լ� ���

    HRight = StrConv(RightB(StrConv(vString, vbFromUnicode), vLen), vbUnicode)

End Function

Public Function HMid(ByVal vString As String, ByVal vLenF As Long, ByVal vLenT As Long) As String
' �ѱ����Ե� ���忡�� mid�Լ� ���

    HMid = StrConv(MidB(StrConv(vString, vbFromUnicode), vLenF, vLenT), vbUnicode)

End Function

Public Function HLen(ByVal vString As String) As Long
' �ѱ����Ե� ���忡�� len�Լ� ���

    HLen = LenB(StrConv(vString, vbFromUnicode))

End Function

Public Sub gsSpreadClear(ByVal brSpread As Object, Optional ByVal brRow As Long = 1000, Optional ByVal brColor As Boolean = False, Optional ByVal brHeight As Integer = 0, _
                         Optional ByVal brRowAdd As Boolean = False)
' �������� Clear
    
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
    '����Ÿ �Ǽ��� 10000���̻��϶��� ���Ϸ� ������ ���´�
    If brSpread.MaxRows > 1000 Then
        If MsgBox("����Ÿ�� 1000�� �̻��Դϴ�." & vbCrLf & vbCrLf & _
            "����Ÿ�� ���� �����ɸ��ų� ��ǻ�Ͱ� �ٿ�� �� �ֽ��ϴ�" & vbCrLf & vbCrLf & _
            "���α׷��� �������� �ʰ� ���Ϸ� �����ðڽ��ϱ�?", vbQuestion + vbYesNo, "���Ϲޱ�") = vbYes Then
            Call gsp_SetSpdExcelFileExport(brSpread, brTitle)
            
            Exit Sub
        End If
    End If
    
    sMaxRows = brSpread.MaxRows
    sMaxCols = brSpread.MaxCols
    sHeadRow = brSpread.ColHeaderRows
    
    ReDim sSpeadData(sMaxRows + sHeadRow, sMaxCols)
    
    Set sExlAPP = CreateObject("Excel.Application")
    
    'Excel�� ��ġ�Ǿ����� ������ ������ �߻��Ѵ�
    If Err.Number > 0 Then
        MsgBox "Excel ���α׷��� ���� �� ���� �����ϴ�." & vbCrLf & _
        "Excel ���α׷��� ��ġ�Ǿ� �ִ��� Ȯ�ιٶ��ϴ�.", vbInformation, "���α׷� ���࿡��"
        
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
            ' Cell�� Hidden �ƴ� �ڷḸ ����
            If brSpread.ColHidden = False Then
                Select Case brSpread.CellType
                    Case CellTypeNumber, CellTypeCurrency
                        sSpeadData(sColAp + sRow, sCol - sHiddenCnt - 1) = brSpread.Text
                    Case Else
                        If IsNumeric(brSpread.Text) = True Then
                            ' �ڷ����� ������ ������ ���
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
    
    ' ������ MAXCOL�� �ش��ϴ� ���ڿ�
    sStrTmp = sStrTmp & Chr(IIf((brSpread.MaxCols - sHiddenCnt) Mod 26 = 0, 26, (brSpread.MaxCols - sHiddenCnt) Mod 26) + 64)

    sColStr = "A1:" & sStrTmp & sMaxRows + sHeadRow
    sExlSheet.Range(sColStr) = sSpeadData                           ' �ڷẹ��
    sExlSheet.Range(sColStr).Font.Size = 9
    If brLineFg Then
        sExlSheet.Range(sColStr).Borders.LineStyle = xlContinuous   ' �ܰ�����
    End If
    
    sColStr = "A1:" & sStrTmp & sHeadRow                            ' �׸�Ÿ��Ʋ
    sExlSheet.Range(sColStr).Font.Bold = False                      ' Ÿ��Ʋ ��Ʈ
    sExlSheet.Range(sColStr).HorizontalAlignment = xlCenter         ' Ÿ��Ʋ ����
    sExlSheet.Range(sColStr).VerticalAlignment = xlCenter           ' Ÿ��Ʋ ����
    sExlSheet.Range(sColStr).Interior.Color = &HF1F1F1              ' Ÿ��Ʋ Background Color
    
    brSpread.Row = 1
    For sCol = 1 To sMaxCols
        brSpread.Col = sCol
        If brSpread.ColHidden = False Then
            sColAp = sColAp + 1
            sColStr = Chr(64 + sColAp) & "1"
            ' Cell�� ���� Spread�� �����ϰ�
            sExlSheet.Range(sColStr).ColumnWidth = brSpread.ColWidth(sCol)
            
            If brSpread.CellType = CellTypeNumber Or brSpread.CellType = CellTypeCurrency Or IsNumeric(brSpread.Text) Then
                ' ������Ÿ���� �����ΰ�� ',' ����
                If brSpread.TypeNumberDecPlaces > 0 Then
                    ' Spread Type�� ������ ��� �Ҽ����ڸ� Ȯ��
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
    
    ' ������Ʈ�� ������
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
    
    sExlSheet.Columns.Font.Name = "����ü"

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


