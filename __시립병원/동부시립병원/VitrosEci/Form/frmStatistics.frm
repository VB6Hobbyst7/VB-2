VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmStatistics 
   Caption         =   "결과통계"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   15210
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList imlList 
      Left            =   10770
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":059A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":0B34
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":10CE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":1668
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":1C02
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   8745
      Left            =   45
      TabIndex        =   12
      Top             =   585
      Width           =   15135
      _Version        =   65536
      _ExtentX        =   26696
      _ExtentY        =   15425
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin FPSpreadADO.fpSpread spdResult1 
         Height          =   8730
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   15090
         _Version        =   393216
         _ExtentX        =   26617
         _ExtentY        =   15399
         _StockProps     =   64
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   32
         MaxRows         =   20
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "frmStatistics.frx":219C
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   510
      Left            =   45
      TabIndex        =   6
      Top             =   45
      Width           =   15135
      _Version        =   65536
      _ExtentX        =   26696
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.OptionButton optCondition 
         Caption         =   "검사 건수"
         Height          =   285
         Index           =   0
         Left            =   5805
         TabIndex        =   9
         Top             =   135
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optCondition 
         Caption         =   "슬립 건수"
         Height          =   285
         Index           =   1
         Left            =   7200
         TabIndex        =   8
         Top             =   135
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdSerch 
         Caption         =   "조  회"
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   105
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1095
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   105
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM"
         Format          =   56688643
         CurrentDate     =   38056
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "조회월 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   30
      TabIndex        =   0
      Top             =   9330
      Width           =   15120
      Begin VB.CommandButton cmdAction 
         Caption         =   "Close"
         Height          =   375
         Index           =   3
         Left            =   10470
         TabIndex        =   4
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Clear"
         Height          =   375
         Index           =   2
         Left            =   9150
         TabIndex        =   3
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Print"
         Height          =   375
         Index           =   1
         Left            =   7830
         TabIndex        =   2
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Save"
         Height          =   375
         Index           =   0
         Left            =   6510
         TabIndex        =   1
         Top             =   120
         Width           =   1245
      End
   End
   Begin MSComctlLib.ListView lvwCuData 
      Height          =   3420
      Left            =   11490
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   6033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAdoRs As ADODB.Recordset
Private CallForm    As String

Private Const COL_WIDTH As Long = "900"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름
Private Const TEST_VALUES   As String = "VALUES"    '결과


Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case 0
            Call cmdExSave
        Case 1
            Call cmdPrint
        Case 2
            Call cmdClear
        Case 3
            Call cmdClose
        Case Else
    End Select
End Sub

Private Sub cmdExSave()
    Dim strPath   As String
    Dim strFilter As String
    Dim strFileName As String

    strPath = App.Path & "\Excel\"
    strFilter = "Excel"
    
    With spdResult1
        If .DataRowCnt > 0 Then
            strFileName = ShowSaveFile(strPath, strFilter)
            If strFileName = "" Then Exit Sub
            Call ssSaveAsExcel(spdResult1, strFileName)
        End If
    End With
    
End Sub


Private Function ssSaveAsExcel(ByVal tbl As fpSpread, ByVal strFileName As String) As String

    'Spread Sheet의 데이타 저장하기
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim tmpTxtFile As String
    Dim WithoutExtNm As String
    Dim DirName  As String
    Dim intPos As Integer

    Screen.ActiveForm.MousePointer = vbHourglass

    DirName = "C:\temp"
    If Dir(DirName, vbDirectory) = "" Then
        MkDir (DirName)
    End If

    'tmpTxtFile = App.Path & "Stastics.txt"
    tmpTxtFile = "C:\temp\ss.txt"

    intPos = InStr(1, strFileName, ".")
    If intPos = 0 Then
        WithoutExtNm = strFileName
    Else
        WithoutExtNm = Mid(strFileName, 1, intPos - 1)
    End If

Dim ExcelSheet As Object
Dim i, j As Integer
'Dim AlphaChr As Integer

Dim AlphaChr(40) As String

    AlphaChr(0) = "A"
    AlphaChr(1) = "B"
    AlphaChr(2) = "C"
    AlphaChr(3) = "D"
    AlphaChr(4) = "E"
    AlphaChr(5) = "F"
    AlphaChr(6) = "G"
    AlphaChr(7) = "H"
    AlphaChr(8) = "I"
    AlphaChr(9) = "J"
    AlphaChr(10) = "K"
    AlphaChr(11) = "L"
    AlphaChr(12) = "M"
    AlphaChr(13) = "N"
    AlphaChr(14) = "O"
    AlphaChr(15) = "P"
    AlphaChr(16) = "Q"
    AlphaChr(17) = "R"
    AlphaChr(18) = "S"
    AlphaChr(19) = "T"
    AlphaChr(20) = "U"
    AlphaChr(21) = "V"
    AlphaChr(22) = "W"
    AlphaChr(23) = "X"
    AlphaChr(24) = "Y"
    AlphaChr(25) = "Z"
    AlphaChr(26) = "AA"
    AlphaChr(27) = "AB"
    AlphaChr(28) = "AC"
    AlphaChr(29) = "AD"
    AlphaChr(30) = "AE"
    AlphaChr(31) = "AF"
    AlphaChr(32) = "AG"
    AlphaChr(33) = "AH"
    AlphaChr(34) = "AI"
    AlphaChr(35) = "AJ"
    AlphaChr(36) = "AK"
    AlphaChr(37) = "AL"
    AlphaChr(38) = "AM"
    AlphaChr(39) = "AN"
    AlphaChr(40) = "AO"
    
    Set ExcelSheet = CreateObject("Excel.Sheet") '<== 엑셀 객체 생성

    For i = 0 To spdResult1.MaxRows
        spdResult1.Row = i
        For j = 0 To spdResult1.MaxCols
            spdResult1.Col = j
            ExcelSheet.Application.Cells(i + 1, AlphaChr(j)).Value = spdResult1.Text
        Next
    Next
    ExcelSheet.SaveAs WithoutExtNm & ".xls"  '<== 엑셀 파일로 저장
    ExcelSheet.Application.Quit              '<== 응용 프로그램의 Quit 메서드로 Excel을 종료합니다
    
    Set ExcelSheet = Nothing                 '<== 개체 변수를 해제합니다
    MsgBox "저장되었습니다.", vbOKOnly + vbInformation, Me.Caption

    Screen.ActiveForm.MousePointer = vbDefault
    
End Function

'-- 엑셀로 저장하기///////////////////////////////////////////////////////////////////////////
Public Function ShowSaveFile(Optional strInitDir As String = "", Optional strFilter As String = "", _
                             Optional strDefaultExt As String = "") As String
    'CommonDialog.ShowSave
    If strInitDir = "" Then
        CommonDialog1.InitDir = App.Path
    Else
        CommonDialog1.InitDir = strInitDir
    End If
    CommonDialog1.Filter = strFilter
    CommonDialog1.DefaultExt = strDefaultExt
    CommonDialog1.ShowSave
    
    ShowSaveFile = CommonDialog1.FileName
    
End Function

Private Sub Load_From(ByVal frm As Form)
    
    With frm
        .Show
        .SetFocus
        
    End With
    
End Sub

Private Sub cmdPrint()
    Dim strPage As String
    Dim strArea As String
    Dim strPDate As String
    
    strPage = "Page : " & Space(7) & "/p" & " of " & spdResult1.PrintPageCount
    strArea = ""
    strPDate = "출력일자:" & Format(Now, "yyyy년mm월dd일")
    
    With SpPrint
        .strTitle = "/fn""굴림체""/fz""20""/fb1/fi0/fu1/fk0/fs1" _
                  & "/f1/c검사통계(" & Format(dtpDate.Value, "yyyy-mm") & ")/n"
        .strBaseDate = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs1" _
                     & "/f1/c" & "" & "/n/n"
        .strPageCount = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r" & strPage & "/n"
        .strAreaName = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/l" & strArea
        .strPrintDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r" & strPDate & ""
    End With

    Call Load_From(frmSpPreview)

End Sub

Private Sub cmdClose()
    
    Unload Me

End Sub

Private Sub cmdSerch_Click()
    
    Screen.MousePointer = 11
    spdResult1.Visible = False
    
    Call cmdClear
    Call f_subSet_ItemList
    Call f_subSet_StatList
    
    spdResult1.Visible = True
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
    
    dtpDate.Value = Format(Now, "yyyy-mm")
    
    Call cmdClear
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList

End Sub

Private Sub f_subSet_ItemHeader()
    
    '검사코드 테이블
    With lvwCuData
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HideColumnHeaders = True
        With .ColumnHeaders
            .Clear
            Call .Add(, TEST_NM_EQP, "ID", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_CD_LIS, "검사코드", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "검 사 명", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "검사결과", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFLM", "참고치남(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFHM", "참고치남(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFLF", "참고치여(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFHF", "참고치여(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "재검", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "검체코드", (lvwCuData.Width - 310) * 0.1)
        End With
        .HideColumnHeaders = False
    End With
    
   
End Sub

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intRow  As Integer, intCnt  As Integer
    Dim DayWeek As Integer
    Dim intWeek As Integer
    Dim MonLastDay As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 14
'        .ColWidth(-1) = 4
    End With
    
    intRow = 1
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFLM, REFHM, REFLF, REFHF, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") " & _
             "   and ((TESTCD <> '') and (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: spdResult1.MaxRows = adoRS.RecordCount
    
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TESTNM") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFLM") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFHM") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("REFLF") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REFHF") & "")
            itemX.SubItems(12) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(13) = Trim(adoRS.Fields("REMARK") & "")
            itemX.Tag = Trim(adoRS.Fields("TESTNM") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        With spdResult1
            If intRow > .MaxRows Then .MaxRows = .MaxRows + 1
            .SetText 0, intRow, Trim$(adoRS("TESTNM") & "")
        End With
        intRow = intRow + 1
        adoRS.MoveNext
    Loop
        
    intWeek = 0
    intCol = 1
    
    If Month(dtpDate) Mod 2 = 0 Then
        MonLastDay = 30
    Else
        MonLastDay = 31
    End If
    
    If Month(dtpDate) = 2 Then
        MonLastDay = 29
    End If
    
    spdResult1.MaxCols = MonLastDay
    
    For intCnt = 1 To MonLastDay
        DayWeek = Weekday(Format(dtpDate, "yyyy-mm") + "-" + Format(intCnt, "00"))
        '-- 일계
        With spdResult1
            If intCol + intWeek > .MaxCols Then .MaxCols = .MaxCols + 1: .ColWidth(.MaxCols) = 3.25
            .SetText intCol + intWeek, 0, intCnt
        End With
        
        If DayWeek = 1 Then
            '-- 주계
            intWeek = intWeek + 1
            With spdResult1
                If intCol + intWeek > .MaxCols Then .MaxCols = .MaxCols + 1: .ColWidth(.MaxCols) = 3.25
                .SetText intCol + intWeek, 0, "주계"
            End With
        End If
        
        intCol = intCol + 1
        
    Next
    
    With spdResult1
        If intCol + intWeek > .MaxCols Then .MaxCols = .MaxCols + 1: .ColWidth(.MaxCols) = 3.25
        .SetText intCol + intWeek, 0, "월계"
    End With
    
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub f_subSet_StatList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intRow  As Integer, intCnt  As Integer
    Dim DayWeek As Integer
    Dim intWeek As Integer
    Dim MonLastDay As Integer
    Dim varTmp
    Dim varTmpName
    Dim WeekStat() As Integer
    Dim MonthStat() As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_StatList()"
    
    
    Erase WeekStat
    Erase MonthStat
            
    With spdResult1
        ReDim WeekStat(.MaxRows) As Integer
        ReDim MonthStat(.MaxRows) As Integer
        
        For intCol = 1 To .MaxCols
            For intRow = 1 To .MaxRows
                .GetText intCol, 0, varTmp
                If InStr(varTmp, "주계") > 0 Then
                    .Col = intCol: .Row = intRow
                    .ForeColor = vbBlue
                    .BackColor = &HC0FFFF   'vbYellow
                    .SetText intCol, intRow, IIf(WeekStat(intRow) = 0, "", WeekStat(intRow)) 'WeekStat(intRow) '
                    If intRow = .MaxRows Then
                        For intWeek = 1 To .MaxRows
                            WeekStat(intWeek) = 0
                        Next
                    End If
                ElseIf InStr(varTmp, "월계") > 0 Then
                    .Col = intCol: .Row = intRow
                    .ForeColor = vbRed
                    .BackColor = &HC0C0FF
                    .SetText intCol, intRow, IIf(MonthStat(intRow) = 0, "", MonthStat(intRow)) 'MonthStat(intRow) '
                Else
                    varTmp = Format(dtpDate, "yyyy-mm") & "-" & Format(varTmp, "00")
                    
                    .GetText 0, intRow, varTmpName
                    Set itemX = lvwCuData.FindItem(Trim$(varTmpName & ""), lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then
                        sqlDoc = "Select count(*) From INTERFACE003" & _
                                 " Where TransDt = '" & varTmp & "' " & _
                                 "   and TestCd = '" & itemX.SubItems(1) & "' "
                    
                        adoRS.CursorLocation = adUseClient
                        adoRS.Open sqlDoc, AdoCn_Jet
                        
                        .Col = intCol: .Row = intRow
                        .ForeColor = vbBlack
                        .SetText intCol, intRow, IIf(adoRS.Fields(0) = 0, "", adoRS.Fields(0)) 'adoRS.Fields(0) '
                        
                        WeekStat(intRow) = WeekStat(intRow) + adoRS.Fields(0)
                        MonthStat(intRow) = MonthStat(intRow) + adoRS.Fields(0)
                        adoRS.Close
                    End If
                End If
            Next
        Next
    End With
   
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub cmdClear()
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        '.ColWidth(-1) = 3
        .RowHeight(-1) = 14
    End With

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
'    Call lvwStatics(1).Move(lvwStatics(0).left, lvwStatics(0).Top, lvwStatics(0).Width, lvwStatics(0).Height)
    Call fraCmdBar.Move(ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60)
    
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
End Sub
