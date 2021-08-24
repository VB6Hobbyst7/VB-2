VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmProficiencyDAU 
   Caption         =   "기기간 데이타 비교"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   13740
   Begin VB.CommandButton cmdReCal 
      Caption         =   "재계산"
      Height          =   315
      Left            =   6360
      TabIndex        =   21
      Top             =   1440
      Width           =   915
   End
   Begin VB.CommandButton cmdSch 
      Caption         =   "조회"
      Height          =   315
      Left            =   6360
      TabIndex        =   20
      Top             =   1080
      Width           =   915
   End
   Begin VB.CommandButton cmdSelect1 
      Caption         =   "선택"
      Height          =   315
      Left            =   5370
      TabIndex        =   19
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCall1 
      Caption         =   "가져오기"
      Height          =   315
      Left            =   4380
      TabIndex        =   18
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "장비간 비교"
      Height          =   345
      Left            =   2850
      TabIndex        =   17
      Top             =   1080
      Width           =   1485
   End
   Begin VB.CommandButton cmdExcelLoad 
      Caption         =   "ExcelLoad"
      Height          =   285
      Left            =   1770
      TabIndex        =   16
      Top             =   30
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint1 
      Caption         =   "출력1"
      Height          =   345
      Left            =   11700
      TabIndex        =   15
      Top             =   6210
      Width           =   825
   End
   Begin VB.CommandButton cmdPrint2 
      Caption         =   "출력2"
      Height          =   345
      Left            =   12540
      TabIndex        =   14
      Top             =   6210
      Width           =   825
   End
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   3915
      Left            =   5490
      TabIndex        =   13
      Top             =   2730
      Visible         =   0   'False
      Width           =   3045
      _Version        =   196613
      _ExtentX        =   5371
      _ExtentY        =   6906
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmProficiencyDAU.frx":0000
   End
   Begin ChartfxLibCtl.ChartFX ChartFX1 
      Height          =   3135
      Left            =   6750
      TabIndex        =   12
      Top             =   6210
      Width           =   6615
      _cx             =   11668
      _cy             =   5530
      Build           =   19
      TypeMask        =   109576193
      MarkerShape     =   2
      AxesStyle       =   3
      Axis(0).Max     =   90
      Axis(0).Style   =   14344
      Axis(2).Min     =   0
      Axis(2).Max     =   1
      Axis(2).Decimals=   2
      Axis(2).Style   =   10280
      RGBBk           =   16777216
      RGB2DBk         =   16777216
      nColors         =   16
      Colors          =   "frmProficiencyDAU.frx":9C56
      Multi           =   "frmProficiencyDAU.frx":9CF6
      MMask           =   8197
      _Data_          =   "frmProficiencyDAU.frx":9D5E
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      Height          =   315
      Left            =   5460
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "조회"
      Height          =   315
      Left            =   6060
      TabIndex        =   10
      Top             =   690
      Width           =   615
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   3450
      Width           =   2415
      _Version        =   196613
      _ExtentX        =   4260
      _ExtentY        =   7011
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmProficiencyDAU.frx":A03C
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "선택"
      Height          =   315
      Left            =   5460
      TabIndex        =   8
      Top             =   690
      Width           =   615
   End
   Begin FPSpread.vaSpread vasTest 
      Height          =   5985
      Left            =   6750
      TabIndex        =   7
      Top             =   210
      Width           =   6615
      _Version        =   196613
      _ExtentX        =   11668
      _ExtentY        =   10557
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeaderDisplay=   0
      RowsFrozen      =   1
      SpreadDesigner  =   "frmProficiencyDAU.frx":A233
   End
   Begin VB.ComboBox cboEquip 
      Height          =   315
      ItemData        =   "frmProficiencyDAU.frx":109FE
      Left            =   3780
      List            =   "frmProficiencyDAU.frx":10A0B
      TabIndex        =   4
      Top             =   690
      Width           =   1635
   End
   Begin VB.CheckBox chkAll 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   1
      Top             =   120
      Width           =   195
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   8985
      Left            =   210
      TabIndex        =   0
      Top             =   360
      Width           =   2445
      _Version        =   196613
      _ExtentX        =   4313
      _ExtentY        =   15849
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxRows         =   35
      ScrollBars      =   2
      SpreadDesigner  =   "frmProficiencyDAU.frx":10A21
   End
   Begin MSComCtl2.DTPicker dtpExamDate 
      Height          =   345
      Left            =   3780
      TabIndex        =   2
      Top             =   240
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24313857
      CurrentDate     =   38584
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7815
      Left            =   2820
      TabIndex        =   6
      Top             =   1530
      Width           =   3825
      _Version        =   196613
      _ExtentX        =   6747
      _ExtentY        =   13785
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      ColsFrozen      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      Protect         =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmProficiencyDAU.frx":13DF1
      UserResize      =   2
      ScrollBarTrack  =   1
      ShowScrollTips  =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사장비"
      Height          =   195
      Left            =   2880
      TabIndex        =   5
      Top             =   750
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사일자"
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   300
      Width           =   840
   End
End
Attribute VB_Name = "frmProficiencyDAU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iRow1, iRow2, iCol1, iCol2 As Long

Private Sub cboEquip_Click()
    Dim lsEquip As String
    
    If cboEquip.ListIndex < 0 Or cboEquip.ListIndex > cboEquip.ListCount - 1 Then
        Exit Sub
    End If
    
    Select Case cboEquip.ListIndex
    Case 0
        lsEquip = "XE"
    Case 1
        lsEquip = "IPU1"
    Case 2
        lsEquip = "IPU2"
    End Select
    
    ClearSpread vasList
    vasList.Row = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col = 1
    vasList.Col2 = 1
    vasList.BlockMode = True
    vasList.Value = 0
    vasList.BlockMode = False
    
    SQL = "Select barcode, max(diskno), max(posno) " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examtype = '" & lsEquip & "' " & vbCrLf & _
          "group by barcode " & vbCrLf & _
          "Order by barcode "

    res = db_select_Vas(gLocal, SQL, vasList, 1, 2)
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub chkAll_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasExam.DataRowCnt
        vasExam.Row = lRow
        vasExam.Col = 1
        vasExam.Value = chkAll.Value
    Next lRow
End Sub

Private Sub cmdCall1_Click()
    Dim lsEquip As String
    
    If cboEquip.ListIndex < 0 Or cboEquip.ListIndex > cboEquip.ListCount - 1 Then
        Exit Sub
    End If
    
    ClearSpread vasList
    vasList.Row = -1
    vasList.Col = 1
    vasList.Value = 0
    
    Select Case cboEquip.ListIndex
    Case 0
        lsEquip = "XE"
    Case 1
        lsEquip = "IPU1"
    Case 2
        lsEquip = "IPU2"
    End Select
    
    SQL = "select distinct 1, barcode, 'XE', '" & lsEquip & "' from pat_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "'  " & vbCrLf & _
          "  and barcode in (select barcode from pat_res " & vbCrLf & _
                            "where equipno = '" & gEquip & "' " & vbCrLf & _
                            "  and examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
                            "  and examtype = '" & lsEquip & "' )"
    res = db_select_Vas(gLocal, SQL, vasList, 1, 1)
          
End Sub

Private Sub cmdClear_Click()
    ClearSpread vasList
    vasList.Row = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col = 1
    vasList.Col2 = 1
    vasList.BlockMode = True
    vasList.Value = 0
    vasList.BlockMode = False
    
    ClearSpread vasTemp
    ClearSpread vasTest
    
End Sub

Private Sub cmdPrint_Click()
    'Print_P '세로
    Print_L '가로
End Sub

Sub Print_P()
Dim l, r, t, b As Integer
Dim px, py As Integer
Dim w, h, gap As Integer
Dim PageCount As Integer
Dim i         As Integer
Dim sHead       As String


'On Error GoTo PrtErr

sHead = "  검사일자 : " & dtpExamDate

Call vasTest.OwnerPrintPageCount(Printer.hDC, 550, 1200, (Printer.Width / 2) - 300, (Printer.Height) - 300, PageCount)

Printer.Orientation = 1

Printer.FontSize = 13
Printer.Print ""
Printer.Print Tab(5); "▣ Proficiency Testing Worksheet ▣ "
Printer.FontSize = 10
Printer.Print ""
Printer.Print Tab(5); sHead; Tab(150); "Page : " & i & "/" & PageCount
Printer.FontSize = 9
Printer.Print ""


'Call vasTest.OwnerPrintDraw(Printer.hDC, 550, (Printer.Height / 4) - 300, (Printer.Width - 300), (Printer.Height / 2) + 300, 1)
Call vasTest.OwnerPrintDraw(Printer.hDC, 550, 1200, (Printer.Width / 2) - 300, (Printer.Height) - 300, 1)

px = Printer.TwipsPerPixelX
py = Printer.TwipsPerPixelY
w = Printer.Width
h = Printer.Height
gap = 100 / px
t = 10 * gap
b = ((h / 2) / px)
r = (w / px) - (gap * 5)
l = 3 * gap


t = b + (10 * gap)
b = (h / px) - (5 * gap)

r = (w / px) - (gap * 5)
ChartFX1.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0

Printer.EndDoc

End Sub

Sub Print_L()
Dim l, r, t, b As Integer
Dim px, py As Integer
Dim w, h, gap As Integer
Dim PageCount As Integer
Dim i         As Integer
Dim sHead       As String


On Error GoTo PrtErr

sHead = " 검사일자 : " & dtpExamDate

Call vasTest.OwnerPrintPageCount(Printer.hDC, 550, 1200, (Printer.Width / 2) - 300, Printer.Height - 300, PageCount)

Printer.Orientation = 2

Printer.FontSize = 13
Printer.Print ""
'Printer.Print Tab(15); "▣ Proficiency Testing Worksheet ▣ "
Printer.Print Tab(13); "▣ 서로 다른 검사기기 방법에 대한 Data 비교 ▣ "
Printer.FontSize = 10
Printer.Print ""
Printer.Print Tab(13); sHead; 'Tab(150); "Page : " & i & "/" & PageCount
Printer.FontSize = 10
Printer.Print ""


'Call vasTest.OwnerPrintDraw(Printer.hDC, 550, (Printer.Height / 4) - 300, (Printer.Width - 300), (Printer.Height / 2) + 300, 1)
Call vasTest.OwnerPrintDraw(Printer.hDC, 1500, 1200, (Printer.Width / 2) - 300, Printer.Height - 300, 1)

px = Printer.TwipsPerPixelX
py = Printer.TwipsPerPixelY
w = Printer.Width
h = Printer.Height
gap = 100 / px
t = 10 * gap
b = ((h / 2) / px)
r = (w / px) - (gap * 5)
'l = 3 * gap
l = (h / px) / 2 + 10 * gap * 3

't = b + (10 * gap)
t = 10 * gap * 2
b = (h / px) - (10 * gap * 4)

r = (w / px) - (gap * 20)
ChartFX1.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0

Printer.CurrentY = 10300
Printer.Print Tab(62); "IJUPPH Clinical Pathology            확인 :        년     월      일            부서 책임자 :                             "
Printer.Print ""
Printer.Print Tab(62); "                                                                                                 담당  staff :                             "
Printer.EndDoc

Exit Sub

PrtErr:
    MsgBox "프린터 오류!"
    Exit Sub
End Sub

Sub Print_L_1()
Dim l, r, t, b As Integer
Dim px, py As Integer
Dim w, h, gap As Integer
Dim PageCount As Integer
Dim i         As Integer
Dim sHead       As String


'On Error GoTo PrtErr

sHead = "  검사일자 : " & dtpExamDate

Call vasTest.OwnerPrintPageCount(Printer.hDC, 550, 1200, (Printer.Width / 2) - 300, Printer.Height - 300, PageCount)

Printer.Orientation = 2

Printer.FontSize = 13
Printer.Print ""
'Printer.Print Tab(15); "▣ Proficiency Testing Worksheet ▣ "
Printer.Print Tab(15); "▣ 서로 다른 검사기기 방법에 대한 Data 비교 ▣ "
Printer.FontSize = 10
Printer.Print ""
Printer.Print Tab(15); sHead; 'Tab(150); "Page : " & i & "/" & PageCount
Printer.FontSize = 9
Printer.Print ""


'Call vasTest.OwnerPrintDraw(Printer.hDC, 550, (Printer.Height / 4) - 300, (Printer.Width - 300), (Printer.Height / 2) + 300, 1)
Call vasTest.OwnerPrintDraw(Printer.hDC, 1500, 1200, (Printer.Width / 2) - 300, Printer.Height - 300, 1)

px = Printer.TwipsPerPixelX
py = Printer.TwipsPerPixelY
w = Printer.Width
h = Printer.Height
gap = 100 / px
t = 10 * gap
b = ((h / 2) / px)
r = (w / px) - (gap * 5)
'l = 3 * gap
l = (h / px) / 2 + 10 * gap * 3

't = b + (10 * gap)
t = 10 * gap * 2
b = (h / px) - (10 * gap * 4)

r = (w / px) - (gap * 20)
ChartFX1.Paint Printer.hDC, l, t, r, b, CPAINT_PRINT, 0

Printer.EndDoc

End Sub

Private Sub cmdExcelLoad_Click()
    frmExcelLoad.Show
End Sub

Private Sub cmdPrint1_Click()
    Dim sHead, sFoot As String
    'Dim Width1, Width2
    
    On Error GoTo PrtErr
    
    'Width1 = vasTest.ColWidth(2)
    vasTest.ColWidth(-1) = 11
    vasTest.RowHeight(vasTest.DataRowCnt - 2) = 0
    vasTest.RowHeight(vasTest.DataRowCnt - 3) = 0
    vasTest.RowHeight(vasTest.DataRowCnt - 7) = 0
    sHead = "/fn""궁서체"" /fz""15"" /fb1 /fi0 /fu0 " & "/c" & "▣  서로 다른 검사기기 방법에 대한 Data 비교  ▣" & "/n/n " & _
                "/fn""굴림체"" /fz""11"" /fb0 /fi0 /fu0 " & "/r" & "  검사일자 : " & dtpExamDate.Value & "          /n"
    'sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "부산백병원 진단검사의학과"
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & "IJUPPH Clinical Pathology" & "/c" & "확인 :        년     월      일 " & "/r" & "부서 책임자 :                             " & "/n/n" & _
            "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/r" & "담당  staff :                             "
    vasTest.PrintOrientation = 2
    vasTest.PrintAbortMsg = "인쇄중 입니다 ..."
    vasTest.PrintJobName = "XE2100 - Data 비교"
    vasTest.PrintHeader = sHead
    vasTest.PrintFooter = sFoot
    vasTest.PrintMarginTop = 600
    vasTest.PrintMarginBottom = 200
    vasTest.PrintMarginLeft = 720
    vasTest.PrintMarginRight = 0
    
    vasTest.PrintColor = False
    vasTest.PrintGrid = True
    
    'Set printing range
    vasTest.Row = 1
    vasTest.Row2 = vasTest.DataRowCnt
    vasTest.Col = 1
    vasTest.Col2 = vasTest.MaxCols
    vasTest.PrintType = PrintTypeCellRange
    
    'vasTest.PrintType = 0  'SS_PRINT_ALL(default)
    
    
    vasTest.PrintShadows = True

    vasTest.Action = 13 'SS_ACTION_PRINT
    
    vasTest.ColWidth(-1) = 10
    vasTest.RowHeight(vasTest.DataRowCnt - 2) = 10.5
    vasTest.RowHeight(vasTest.DataRowCnt - 3) = 10.5
    vasTest.RowHeight(vasTest.DataRowCnt - 7) = 10.5
        
    Exit Sub
PrtErr:
    MsgBox "프린터 오류!"
    Exit Sub
End Sub

Private Sub cmdPrint2_Click()
    Print_L '가로
End Sub

Private Sub cmdReCal_Click()
    Dim lRow As Long
    Dim lCol As Long
    
    Dim lsEquip As String
     
    Dim i As Integer
    
    Dim Sta_Cnt()
    Dim Sta_Sum()
    Dim Sta_Sum2()
    Dim Sta_Sum1()
    Dim Sta_Min()
    Dim Sta_Max()
    Dim Sta_SD()
    Dim Sta_CV()
    Dim a, b, r
    
    Dim pro_rs As ADODB.Recordset
    
    On Error GoTo ErrHandle
        
'    lCol = 0
'    For i = 1 To vasExam.DataRowCnt
'        vasExam.Row = i
'        vasExam.Col = 1
'        If vasExam.Value = 1 Then
'            lCol = lCol + 1
'            SetText vasTest, Trim(GetText(vasExam, i, 2)), 1, lCol
'            SetText vasTest, Trim(GetText(vasExam, i, 3)), 0, lCol
'            lCol = lCol + 1
'            SetText vasTest, Trim(GetText(vasExam, i, 2)), 1, lCol
'            'SetText vasTest, Trim(GetText(vasExam, i, 3)), 0, lCol
'        End If
'    Next i
'
'    vasTest.MaxCols = lCol
'    'vasTest.RowHeight(1) = 0
'
'    For lRow = 2 To vasTemp.DataRowCnt
'        vasTest.SetText 0, lRow, Trim(GetText(vasTemp, lRow, 1))
'
'        SQL = "Select barcode, equipcode, result " & vbCrLf & _
'              "from pat_res " & vbCrLf & _
'              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'              "  and equipno = '" & gEquip & "' " & vbCrLf & _
'              "  and examtype = '" & Trim(GetText(vasTemp, 1, 1)) & "' " & vbCrLf & _
'              "  and barcode = '" & Trim(GetText(vasTemp, lRow, 1)) & "'"
'        Set pro_rs = db_select_rs(gLocal, SQL)
'
'        Do While Not pro_rs.EOF
'            For lCol = 1 To vasTest.MaxCols Step 2
'                Debug.Print CStr(CCur(Trim(GetText(vasTest, 1, lCol))))
'                Debug.Print Trim(pro_rs.Fields.Item(1).Value)
'                If CStr(CCur(Trim(GetText(vasTest, 1, lCol)))) = Trim(pro_rs.Fields.Item(1).Value) Then
'                    SetText vasTest, Trim(pro_rs.Fields.Item(2).Value), lRow + 1, lCol
'                    Exit For
'                End If
'            Next lCol
'            pro_rs.MoveNext
'        Loop
'        pro_rs.Close
'        Set pro_rs = Nothing
'
'        SQL = "Select barcode, equipcode, result " & vbCrLf & _
'              "from pat_res " & vbCrLf & _
'              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'              "  and equipno = '" & gEquip & "' " & vbCrLf & _
'              "  and examtype = '" & Trim(GetText(vasTemp, 1, 2)) & "' " & vbCrLf & _
'              "  and barcode = '" & Trim(GetText(vasTemp, lRow, 2)) & "'"
'        Set pro_rs = db_select_rs(gLocal, SQL)
'
'        Do While Not pro_rs.EOF
'            For lCol = 2 To vasTest.MaxCols Step 2
'                If CStr(CCur(Trim(GetText(vasTest, 1, lCol)))) = Trim(pro_rs.Fields.Item(1).Value) Then
'                    SetText vasTest, Trim(pro_rs.Fields.Item(2).Value), lRow + 1, lCol
'                    Exit For
'                End If
'            Next lCol
'            pro_rs.MoveNext
'        Loop
'
'        pro_rs.Close
'        Set pro_rs = Nothing
'    Next lRow
        
    ReDim Sta_Cnt(1 To vasTest.MaxCols)
    ReDim Sta_Sum(1 To vasTest.MaxCols)     '∑x
    ReDim Sta_Sum1(1 To vasTest.MaxCols)    '∑x * y
    ReDim Sta_Sum2(1 To vasTest.MaxCols)    '∑x^2
    ReDim Sta_Min(1 To vasTest.MaxCols)
    ReDim Sta_Max(1 To vasTest.MaxCols)
    ReDim Sta_SD(1 To vasTest.MaxCols)
    ReDim Sta_CV(1 To vasTest.MaxCols)
    
    For lCol = 1 To vasTest.MaxCols
'        If lCol Mod 2 = 0 Then
'            vasTest.SetText lCol, 1, Trim(GetText(vasTemp, 1, 2))
'        Else
'            vasTest.SetText lCol, 1, Trim(GetText(vasTemp, 1, 2))
'        End If
        
        Sta_Cnt(lCol) = 0
        Sta_Sum(lCol) = 0
        Sta_Sum1(lCol) = 0
        Sta_Sum2(lCol) = 0
        Sta_Min(lCol) = 999
        Sta_Max(lCol) = 0
        Sta_SD(lCol) = 0
        Sta_CV(lCol) = 0
    Next lCol
    
    
    For lRow = 2 To vasTest.DataRowCnt
        'SetText vasTest, lRow - 1, lRow, 0
        For lCol = 1 To vasTest.MaxCols
            If IsNumeric(Trim(GetText(vasTest, lRow, lCol))) Then
                Sta_Cnt(lCol) = Sta_Cnt(lCol) + 1
                Sta_Sum(lCol) = Sta_Sum(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol)))
                Sta_Sum2(lCol) = Sta_Sum2(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol))) ^ 2
                If Sta_Min(lCol) > CCur(Trim(GetText(vasTest, lRow, lCol))) Then
                    Sta_Min(lCol) = CCur(Trim(GetText(vasTest, lRow, lCol)))
                End If
                
                If Sta_Max(lCol) < CCur(Trim(GetText(vasTest, lRow, lCol))) Then
                    Sta_Max(lCol) = CCur(Trim(GetText(vasTest, lRow, lCol)))
                End If
                
                If lCol Mod 2 = 0 Then
                    Sta_Sum1(lCol) = Sta_Sum1(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol))) * CCur(Trim(GetText(vasTest, lRow, lCol - 1)))
                End If
                
            End If
        Next lCol
    Next lRow
    
    i = vasTest.DataRowCnt + 1
    For lCol = 1 To vasTest.MaxCols
        vasTest.SetText lCol, i, Sta_Min(lCol) & "~" & Sta_Max(lCol)
        vasTest.SetText lCol, i + 1, Format(Sta_Sum(lCol) / Sta_Cnt(lCol), "#0.00")
        vasTest.SetText lCol, i + 7, Sta_Cnt(lCol)
    Next lCol
    
    For lRow = 2 To vasTest.DataRowCnt
        SetText vasTest, lRow - 1, lRow, 0
        For lCol = 1 To vasTest.MaxCols
            If IsNumeric(Trim(GetText(vasTest, lRow, lCol))) Then
                Sta_SD(lCol) = Sta_SD(lCol) + (CCur(Trim(GetText(vasTest, lRow, lCol))) - Sta_Sum(lCol) / Sta_Cnt(lCol)) ^ 2
            End If
        Next lCol
    Next lRow
    For lCol = 1 To vasTest.MaxCols
        vasTest.SetText lCol, i + 2, Format(Sqr(Sta_SD(lCol) / Sta_Cnt(lCol)), "#0.00")
        vasTest.SetText lCol, i + 3, Format((Sqr(Sta_SD(lCol) / Sta_Cnt(lCol)) / Sta_Sum(lCol) / Sta_Cnt(lCol) * 100), "#0.00")
        
        If lCol Mod 2 = 0 Then
            '홀수 : Y, 짝수 : X
            a = Format((Sta_Cnt(lCol) * Sta_Sum1(lCol) - Sta_Sum(lCol - 1) * Sta_Sum(lCol)) / (Sta_Cnt(lCol) * Sta_Sum2(lCol) - (Sta_Sum(lCol) ^ 2)), "#0.0000")
            r = Format((Sta_Cnt(lCol) * Sta_Sum1(lCol) - Sta_Sum(lCol - 1) * Sta_Sum(lCol)) / Sqr((Sta_Cnt(lCol) * Sta_Sum2(lCol) - (Sta_Sum(lCol) ^ 2)) * (Sta_Cnt(lCol) * Sta_Sum2(lCol - 1) - (Sta_Sum(lCol - 1) ^ 2))), "#0.0000")
            b = Format(Sta_Sum(lCol - 1) / Sta_Cnt(lCol - 1) - a * Sta_Sum(lCol) / Sta_Cnt(lCol), "#0.0000")
            vasTest.SetText lCol, i + 4, a
            vasTest.SetText lCol, i + 5, b
            vasTest.SetText lCol, i + 6, r
        End If
    Next lCol
    
    
    vasTest.SetText 0, i, "Range"
    vasTest.SetText 0, i + 1, "MEAN"
    vasTest.SetText 0, i + 2, "SD"
    vasTest.SetText 0, i + 3, "CV"
    vasTest.SetText 0, i + 4, "Slope"
    vasTest.SetText 0, i + 5, "Y Intercept"
    vasTest.SetText 0, i + 6, "R"
    vasTest.SetText 0, i + 7, "Cnt"
    vasTest.RowHeight(i + 7) = 0
    
    vasTest.SetText 0, 1, "기종"
    vasTest.SetText 0, 0, "항목"
    
    For lRow = i + 4 To i + 7
        For lCol = 1 To vasTest.DataColCnt
            If lCol Mod 2 = 1 Then
                vasTest.SetCellBorder lCol, lRow, lCol, lRow, 2, RGB(255, 255, 255), 6
            Else
                vasTest.SetCellBorder lCol, lRow, lCol, lRow, 1, RGB(255, 255, 255), 6
            End If
        Next lCol
    Next lRow
    
    Exit Sub
    
ErrHandle:
    'Resume Next
    Exit Sub

End Sub

Private Sub cmdSch_Click()
    Dim lRow As Long
    Dim lCol As Long
    
    Dim lsEquip As String
     
    Dim i As Integer
    
    Dim Sta_Cnt()
    Dim Sta_Sum()
    Dim Sta_Sum2()
    Dim Sta_Sum1()
    Dim Sta_Min()
    Dim Sta_Max()
    Dim Sta_SD()
    Dim Sta_CV()
    Dim a, b, r
    
    Dim pro_rs As ADODB.Recordset
    
    On Error GoTo ErrHandle
        
    lCol = 0
    For i = 1 To vasExam.DataRowCnt
        vasExam.Row = i
        vasExam.Col = 1
        If vasExam.Value = 1 Then
            lCol = lCol + 1
            SetText vasTest, Trim(GetText(vasExam, i, 2)), 1, lCol
            SetText vasTest, Trim(GetText(vasExam, i, 3)), 0, lCol
            lCol = lCol + 1
            SetText vasTest, Trim(GetText(vasExam, i, 2)), 1, lCol
            'SetText vasTest, Trim(GetText(vasExam, i, 3)), 0, lCol
        End If
    Next i
    
    vasTest.MaxCols = lCol
    'vasTest.RowHeight(1) = 0
    
    For lRow = 2 To vasTemp.DataRowCnt
        vasTest.SetText 0, lRow, Trim(GetText(vasTemp, lRow, 1))
        
        SQL = "Select barcode, equipcode, result " & vbCrLf & _
              "from pat_res " & vbCrLf & _
              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
              "  and equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examtype = '" & Trim(GetText(vasTemp, 1, 1)) & "' " & vbCrLf & _
              "  and barcode = '" & Trim(GetText(vasTemp, lRow, 1)) & "'"
        Set pro_rs = db_select_rs(gLocal, SQL)
           
        Do While Not pro_rs.EOF
            For lCol = 1 To vasTest.MaxCols Step 2
                Debug.Print CStr(CCur(Trim(GetText(vasTest, 1, lCol))))
                Debug.Print Trim(pro_rs.Fields.Item(1).Value)
                If CStr(CCur(Trim(GetText(vasTest, 1, lCol)))) = CStr(CCur(Trim(pro_rs.Fields.Item(1).Value))) Then
                    SetText vasTest, Trim(pro_rs.Fields.Item(2).Value), lRow, lCol
                    Exit For
                End If
            Next lCol
            pro_rs.MoveNext
        Loop
        pro_rs.Close
        Set pro_rs = Nothing
        
        SQL = "Select barcode, equipcode, result " & vbCrLf & _
              "from pat_res " & vbCrLf & _
              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
              "  and equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examtype = '" & Trim(GetText(vasTemp, 1, 2)) & "' " & vbCrLf & _
              "  and barcode = '" & Trim(GetText(vasTemp, lRow, 2)) & "'"
        Set pro_rs = db_select_rs(gLocal, SQL)
           
        Do While Not pro_rs.EOF
            For lCol = 2 To vasTest.MaxCols Step 2
                If CStr(CCur(Trim(GetText(vasTest, 1, lCol)))) = CStr(CCur(Trim(pro_rs.Fields.Item(1).Value))) Then
                    SetText vasTest, Trim(pro_rs.Fields.Item(2).Value), lRow, lCol
                    Exit For
                End If
            Next lCol
            pro_rs.MoveNext
        Loop
        
        pro_rs.Close
        Set pro_rs = Nothing
    Next lRow
        
    ReDim Sta_Cnt(1 To vasTest.MaxCols)
    ReDim Sta_Sum(1 To vasTest.MaxCols)     '∑x
    ReDim Sta_Sum1(1 To vasTest.MaxCols)    '∑x * y
    ReDim Sta_Sum2(1 To vasTest.MaxCols)    '∑x^2
    ReDim Sta_Min(1 To vasTest.MaxCols)
    ReDim Sta_Max(1 To vasTest.MaxCols)
    ReDim Sta_SD(1 To vasTest.MaxCols)
    ReDim Sta_CV(1 To vasTest.MaxCols)
    
    For lCol = 1 To vasTest.MaxCols
        If lCol Mod 2 = 0 Then
            vasTest.SetText lCol, 1, Trim(GetText(vasTemp, 1, 2))
        Else
            vasTest.SetText lCol, 1, Trim(GetText(vasTemp, 1, 1))
        End If
        
        Sta_Cnt(lCol) = 0
        Sta_Sum(lCol) = 0
        Sta_Sum1(lCol) = 0
        Sta_Sum2(lCol) = 0
        Sta_Min(lCol) = 999
        Sta_Max(lCol) = 0
        Sta_SD(lCol) = 0
        Sta_CV(lCol) = 0
    Next lCol
    
    
    For lRow = 2 To vasTest.DataRowCnt
        'SetText vasTest, lRow - 1, lRow, 0
        For lCol = 1 To vasTest.MaxCols
            If IsNumeric(Trim(GetText(vasTest, lRow, lCol))) Then
                Sta_Cnt(lCol) = Sta_Cnt(lCol) + 1
                Sta_Sum(lCol) = Sta_Sum(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol)))
                Sta_Sum2(lCol) = Sta_Sum2(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol))) ^ 2
                If Sta_Min(lCol) > CCur(Trim(GetText(vasTest, lRow, lCol))) Then
                    Sta_Min(lCol) = CCur(Trim(GetText(vasTest, lRow, lCol)))
                End If
                
                If Sta_Max(lCol) < CCur(Trim(GetText(vasTest, lRow, lCol))) Then
                    Sta_Max(lCol) = CCur(Trim(GetText(vasTest, lRow, lCol)))
                End If
                
                If lCol Mod 2 = 0 Then
                    Sta_Sum1(lCol) = Sta_Sum1(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol))) * CCur(Trim(GetText(vasTest, lRow, lCol - 1)))
                End If
                
            End If
        Next lCol
    Next lRow
    
    i = vasTest.DataRowCnt + 1
    For lCol = 1 To vasTest.MaxCols
        vasTest.SetText lCol, i, Sta_Min(lCol) & "~" & Sta_Max(lCol)
        vasTest.SetText lCol, i + 1, Format(Sta_Sum(lCol) / Sta_Cnt(lCol), "#0.00")
        vasTest.SetText lCol, i + 7, Sta_Cnt(lCol)
    Next lCol
    
    For lRow = 2 To vasTest.DataRowCnt
        'SetText vasTest, lRow - 1, lRow, 0
        For lCol = 1 To vasTest.MaxCols
            If IsNumeric(Trim(GetText(vasTest, lRow, lCol))) Then
                Sta_SD(lCol) = Sta_SD(lCol) + (CCur(Trim(GetText(vasTest, lRow, lCol))) - Sta_Sum(lCol) / Sta_Cnt(lCol)) ^ 2
            End If
        Next lCol
    Next lRow
    For lCol = 1 To vasTest.MaxCols
        vasTest.SetText lCol, i + 2, Format(Sqr(Sta_SD(lCol) / Sta_Cnt(lCol)), "#0.00")
        vasTest.SetText lCol, i + 3, Format((Sqr(Sta_SD(lCol) / Sta_Cnt(lCol)) / Sta_Sum(lCol) / Sta_Cnt(lCol) * 100), "#0.00")
        
        If lCol Mod 2 = 0 Then
            '홀수 : Y, 짝수 : X
            a = Format((Sta_Cnt(lCol) * Sta_Sum1(lCol) - Sta_Sum(lCol - 1) * Sta_Sum(lCol)) / (Sta_Cnt(lCol) * Sta_Sum2(lCol) - (Sta_Sum(lCol) ^ 2)), "#0.0000")
            r = Format((Sta_Cnt(lCol) * Sta_Sum1(lCol) - Sta_Sum(lCol - 1) * Sta_Sum(lCol)) / Sqr((Sta_Cnt(lCol) * Sta_Sum2(lCol) - (Sta_Sum(lCol) ^ 2)) * (Sta_Cnt(lCol) * Sta_Sum2(lCol - 1) - (Sta_Sum(lCol - 1) ^ 2))), "#0.0000")
            b = Format(Sta_Sum(lCol - 1) / Sta_Cnt(lCol - 1) - a * Sta_Sum(lCol) / Sta_Cnt(lCol), "#0.0000")
            vasTest.SetText lCol, i + 4, a
            vasTest.SetText lCol, i + 5, b
            vasTest.SetText lCol, i + 6, r
        End If
    Next lCol
    
    
    vasTest.SetText 0, i, "Range"
    vasTest.SetText 0, i + 1, "MEAN"
    vasTest.SetText 0, i + 2, "SD"
    vasTest.SetText 0, i + 3, "CV"
    vasTest.SetText 0, i + 4, "Slope"
    vasTest.SetText 0, i + 5, "Y Intercept"
    vasTest.SetText 0, i + 6, "R"
    vasTest.SetText 0, i + 7, "Cnt"
    vasTest.RowHeight(i + 7) = 0
    
    vasTest.SetText 0, 1, "기종"
    vasTest.SetText 0, 0, "항목"
    
    For lRow = i + 4 To i + 7
        For lCol = 1 To vasTest.DataColCnt
            If lCol Mod 2 = 1 Then
                vasTest.SetCellBorder lCol, lRow, lCol, lRow, 2, RGB(255, 255, 255), 6
            Else
                vasTest.SetCellBorder lCol, lRow, lCol, lRow, 1, RGB(255, 255, 255), 6
            End If
        Next lCol
    Next lRow
    
    Exit Sub
    
ErrHandle:
    'Resume Next
    Exit Sub

End Sub

Private Sub cmdSearch_Click()
    Dim lRow As Long
    Dim lCol As Long
    
    Dim lsEquip As String
     
    Dim i As Integer
    
    Dim Sta_Cnt()
    Dim Sta_Sum()
    Dim Sta_Sum2()
    Dim Sta_Sum1()
    Dim Sta_Min()
    Dim Sta_Max()
    Dim Sta_SD()
    Dim Sta_CV()
    Dim a, b, r
    
    Dim pro_rs As ADODB.Recordset
    
    On Error GoTo ErrHandle
        
    lCol = 0
    For i = 1 To vasExam.DataRowCnt
        vasExam.Row = i
        vasExam.Col = 1
        If vasExam.Value = 1 Then
            lCol = lCol + 1
            SetText vasTest, Trim(GetText(vasExam, i, 2)), 1, lCol
            SetText vasTest, Trim(GetText(vasExam, i, 3)), 0, lCol
            lCol = lCol + 1
            SetText vasTest, Trim(GetText(vasExam, i, 2)), 1, lCol
            'SetText vasTest, Trim(GetText(vasExam, i, 3)), 0, lCol
        End If
    Next i
    
    vasTest.MaxCols = lCol
    'vasTest.RowHeight(1) = 0
    
    For lRow = 1 To vasTemp.DataRowCnt
        SQL = "Select barcode, equipcode, result " & vbCrLf & _
              "from pat_res " & vbCrLf & _
              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
              "  and equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examtype = 'IPU1' " & vbCrLf & _
              "  and barcode = '" & Trim(GetText(vasTemp, lRow, 1)) & "'"
        Set pro_rs = db_select_rs(gLocal, SQL)
           
        Do While Not pro_rs.EOF
            For lCol = 1 To vasTest.MaxCols Step 2
                Debug.Print CStr(CCur(Trim(GetText(vasTest, 1, lCol))))
                Debug.Print Trim(pro_rs.Fields.Item(1).Value)
                If CStr(CCur(Trim(GetText(vasTest, 1, lCol)))) = Trim(pro_rs.Fields.Item(1).Value) Then
                    SetText vasTest, Trim(pro_rs.Fields.Item(2).Value), lRow + 1, lCol
                    Exit For
                End If
            Next lCol
            pro_rs.MoveNext
        Loop
        pro_rs.Close
        Set pro_rs = Nothing
        
        SQL = "Select barcode, equipcode, result " & vbCrLf & _
              "from pat_res " & vbCrLf & _
              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
              "  and equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examtype = 'IPU2' " & vbCrLf & _
              "  and barcode = '" & Trim(GetText(vasTemp, lRow, 2)) & "'"
        Set pro_rs = db_select_rs(gLocal, SQL)
           
        Do While Not pro_rs.EOF
            For lCol = 2 To vasTest.MaxCols Step 2
                If CStr(CCur(Trim(GetText(vasTest, 1, lCol)))) = Trim(pro_rs.Fields.Item(1).Value) Then
                    SetText vasTest, Trim(pro_rs.Fields.Item(2).Value), lRow + 1, lCol
                    Exit For
                End If
            Next lCol
            pro_rs.MoveNext
        Loop
        
        pro_rs.Close
        Set pro_rs = Nothing
    Next lRow
        
    ReDim Sta_Cnt(1 To vasTest.MaxCols)
    ReDim Sta_Sum(1 To vasTest.MaxCols)     '∑x
    ReDim Sta_Sum1(1 To vasTest.MaxCols)    '∑x * y
    ReDim Sta_Sum2(1 To vasTest.MaxCols)    '∑x^2
    ReDim Sta_Min(1 To vasTest.MaxCols)
    ReDim Sta_Max(1 To vasTest.MaxCols)
    ReDim Sta_SD(1 To vasTest.MaxCols)
    ReDim Sta_CV(1 To vasTest.MaxCols)
    
    For lCol = 1 To vasTest.MaxCols
        If lCol Mod 2 = 0 Then
            vasTest.SetText lCol, 1, "IPU2"
        Else
            vasTest.SetText lCol, 1, "IPU1"
        End If
        
        Sta_Cnt(lCol) = 0
        Sta_Sum(lCol) = 0
        Sta_Sum1(lCol) = 0
        Sta_Sum2(lCol) = 0
        Sta_Min(lCol) = 999
        Sta_Max(lCol) = 0
        Sta_SD(lCol) = 0
        Sta_CV(lCol) = 0
    Next lCol
    
    
    For lRow = 2 To vasTest.DataRowCnt
        SetText vasTest, lRow - 1, lRow, 0
        For lCol = 1 To vasTest.MaxCols
            If IsNumeric(Trim(GetText(vasTest, lRow, lCol))) Then
                Sta_Cnt(lCol) = Sta_Cnt(lCol) + 1
                Sta_Sum(lCol) = Sta_Sum(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol)))
                Sta_Sum2(lCol) = Sta_Sum2(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol))) ^ 2
                If Sta_Min(lCol) > CCur(Trim(GetText(vasTest, lRow, lCol))) Then
                    Sta_Min(lCol) = CCur(Trim(GetText(vasTest, lRow, lCol)))
                End If
                
                If Sta_Max(lCol) < CCur(Trim(GetText(vasTest, lRow, lCol))) Then
                    Sta_Max(lCol) = CCur(Trim(GetText(vasTest, lRow, lCol)))
                End If
                
                If lCol Mod 2 = 0 Then
                    Sta_Sum1(lCol) = Sta_Sum1(lCol) + CCur(Trim(GetText(vasTest, lRow, lCol))) * CCur(Trim(GetText(vasTest, lRow, lCol - 1)))
                End If
                
            End If
        Next lCol
    Next lRow
    
    i = vasTest.DataRowCnt + 1
    For lCol = 1 To vasTest.MaxCols
        vasTest.SetText lCol, i, Sta_Min(lCol) & "~" & Sta_Max(lCol)
        vasTest.SetText lCol, i + 1, Format(Sta_Sum(lCol) / Sta_Cnt(lCol), "#0.00")
        vasTest.SetText lCol, i + 7, Sta_Cnt(lCol)
    Next lCol
    
    For lRow = 2 To vasTest.DataRowCnt
        SetText vasTest, lRow - 1, lRow, 0
        For lCol = 1 To vasTest.MaxCols
            If IsNumeric(Trim(GetText(vasTest, lRow, lCol))) Then
                Sta_SD(lCol) = Sta_SD(lCol) + (CCur(Trim(GetText(vasTest, lRow, lCol))) - Sta_Sum(lCol) / Sta_Cnt(lCol)) ^ 2
            End If
        Next lCol
    Next lRow
    For lCol = 1 To vasTest.MaxCols
        vasTest.SetText lCol, i + 2, Format(Sqr(Sta_SD(lCol) / Sta_Cnt(lCol)), "#0.00")
        vasTest.SetText lCol, i + 3, Format((Sqr(Sta_SD(lCol) / Sta_Cnt(lCol)) / Sta_Sum(lCol) / Sta_Cnt(lCol) * 100), "#0.00")
        
        If lCol Mod 2 = 0 Then
            '홀수 : Y, 짝수 : X
            a = Format((Sta_Cnt(lCol) * Sta_Sum1(lCol) - Sta_Sum(lCol - 1) * Sta_Sum(lCol)) / (Sta_Cnt(lCol) * Sta_Sum2(lCol) - (Sta_Sum(lCol) ^ 2)), "#0.0000")
            r = Format((Sta_Cnt(lCol) * Sta_Sum1(lCol) - Sta_Sum(lCol - 1) * Sta_Sum(lCol)) / Sqr((Sta_Cnt(lCol) * Sta_Sum2(lCol) - (Sta_Sum(lCol) ^ 2)) * (Sta_Cnt(lCol) * Sta_Sum2(lCol - 1) - (Sta_Sum(lCol - 1) ^ 2))), "#0.0000")
            b = Format(Sta_Sum(lCol - 1) / Sta_Cnt(lCol - 1) - a * Sta_Sum(lCol) / Sta_Cnt(lCol), "#0.0000")
            vasTest.SetText lCol, i + 4, a
            vasTest.SetText lCol, i + 5, b
            vasTest.SetText lCol, i + 6, r
        End If
    Next lCol
    
    
    vasTest.SetText 0, i, "Range"
    vasTest.SetText 0, i + 1, "MEAN"
    vasTest.SetText 0, i + 2, "SD"
    vasTest.SetText 0, i + 3, "CV"
    vasTest.SetText 0, i + 4, "Slope"
    vasTest.SetText 0, i + 5, "Y Intercept"
    vasTest.SetText 0, i + 6, "R"
    vasTest.SetText 0, i + 7, "Cnt"
    vasTest.RowHeight(i + 7) = 0
    
    vasTest.SetText 0, 1, "기종"
    vasTest.SetText 0, 0, "항목"
    
    For lRow = i + 4 To i + 7
        For lCol = 1 To vasTest.DataColCnt
            If lCol Mod 2 = 1 Then
                vasTest.SetCellBorder lCol, lRow, lCol, lRow, 2, RGB(255, 255, 255), 6
            Else
                vasTest.SetCellBorder lCol, lRow, lCol, lRow, 1, RGB(255, 255, 255), 6
            End If
        Next lCol
    Next lRow
    
    Exit Sub
    
ErrHandle:
    'Resume Next
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
    Dim lCol As Long
    Dim lRow As Long
    Dim i As Long
    
    If cboEquip.ListIndex = 0 Then
        lCol = 1
    Else
        lCol = 2
    End If
    
    Dim lsEquip As String
    
    Select Case cboEquip.ListIndex
    Case 0
        lsEquip = "XE"
    Case 1
        lsEquip = "IPU1"
    Case 2
        lsEquip = "IPU2"
    End Select
    
    If vasTemp.DataColCnt > 0 Then
        lCol = vasTemp.DataColCnt + 1
    Else
        lCol = 1
    End If
    
    vasTemp.Row = 1
    vasTemp.Col = lCol
    vasTemp.Row2 = vasTemp.MaxRows
    vasTemp.Col2 = lCol
    vasTemp.BlockMode = True
    vasTemp.Action = 3
    vasTemp.BlockMode = False
    
    lRow = 1
    vasTemp.SetText lCol, 1, lsEquip
    
    For i = 1 To vasList.DataRowCnt
        vasList.Row = i
        vasList.Col = 1
        If vasList.Value = 1 Then
            lRow = lRow + 1
            
            SetText vasTemp, Trim(GetText(vasList, i, 2)), lRow, lCol
            
            'SetText vasTest, Trim(GetText(vasList, i, 2)), lRow + 1, lCol
        End If
    Next i
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdSelect1_Click()
    Dim lCol As Long
    Dim lRow As Long
    Dim i As Long
    
    If cboEquip.ListIndex = 0 Then
        lCol = 1
    Else
        lCol = 2
    End If
    
    ClearSpread vasTemp
    
'    vasTemp.Row = 1
'    vasTemp.Col = lCol
'    vasTemp.Row2 = vasTemp.MaxRows
'    vasTemp.Col2 = lCol
'    vasTemp.BlockMode = True
'    vasTemp.Action = 3
'    vasTemp.BlockMode = False
    
    lRow = 1
    SetText vasTemp, Trim(GetText(vasList, 1, 3)), lRow, 1
    SetText vasTemp, Trim(GetText(vasList, 1, 4)), lRow, 2
    
    For i = 1 To vasList.DataRowCnt
        vasList.Row = i
        vasList.Col = 1
        If vasList.Value = 1 Then
            lRow = lRow + 1
            
            SetText vasTemp, Trim(GetText(vasList, i, 2)), lRow, 1
            SetText vasTemp, Trim(GetText(vasList, i, 2)), lRow, 2
            
            'SetText vasTest, Trim(GetText(vasList, i, 2)), lRow + 1, lCol
        End If
    Next i
End Sub

Private Sub Form_Load()

    dtpExamDate.Value = Format(CDate(GetDateFull), "yyyy-mm-dd")
    
    ClearSpread vasExam
    
    SQL = "SELECT distinct EquipCode, ExamName, seqno" & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          " Order by seqno "
    db_select_Vas gLocal, SQL, vasExam, 1, 2
    vasExam.MaxRows = vasExam.DataRowCnt
End Sub

Private Sub vasList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    If BlockRow > BlockRow2 Then
        iRow1 = BlockRow2
        iRow2 = BlockRow
    Else
        iRow1 = BlockRow
        iRow2 = BlockRow2
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    iRow1 = Row
    iCol1 = Col
    iRow2 = Row
    iCol2 = Col
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        Select Case Col
        Case 2
            vasSort vasList, 2
        Case 3, 4
            vasSort vasList, 3, 4
        End Select
    End If
End Sub

Private Sub vasList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim i As Long
    
    If iRow1 < 1 Or iRow2 < 1 Then Exit Sub
        
    For i = iRow1 To iRow2
        vasList.Row = i
        vasList.Col = 1
        vasList.Value = 1
    Next i
End Sub

Private Sub vasTest_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lCol As Long
    Dim lRow As Long
    Dim XMin, XMax, YMin, YMax
    Dim X, Y
    Dim i
    
    On Error GoTo ErrHandle
    
    If Col < 1 Or Col > vasTest.DataColCnt Then Exit Sub
    
    If Col Mod 2 = 0 Then
        lCol = Col
    Else
        lCol = Col + 1
    End If
    
    ClearGraph ChartFX1
    
    ClearSpread vasTemp1
    For lRow = 2 To vasTest.DataRowCnt - 8
        vasTemp1.SetText lCol - 1, lRow, Trim(GetText(vasTest, lRow, lCol - 1))
        vasTemp1.SetText lCol, lRow, Trim(GetText(vasTest, lRow, lCol))
    Next lRow
    vasSort vasTemp1, lCol
    
    InsertRow vasTemp1, 1
        
    'Exit Sub
    
'    YMin = Trim(GetText(vasTest, vasTest.DataRowCnt - 7, lCol - 1))
'    i = InStr(1, YMin, "~")
'    YMax = Mid(YMin, i + 1)
'    YMin = Left(YMin, i - 1)
'    XMin = Trim(GetText(vasTest, vasTest.DataRowCnt - 7, lCol))
'    i = InStr(1, XMin, "~")
'    XMax = Mid(XMin, i + 1)
'    XMin = Left(XMin, i - 1)
    
    YMin = Trim(GetText(vasTest, vasTest.DataRowCnt - 7, lCol))
    i = InStr(1, YMin, "~")
    YMax = Mid(YMin, i + 1)
    YMin = Left(YMin, i - 1)
    XMin = Trim(GetText(vasTest, vasTest.DataRowCnt - 7, lCol - 1))
    i = InStr(1, XMin, "~")
    XMax = Mid(XMin, i + 1)
    XMin = Left(XMin, i - 1)
    
    ChartFX1.OpenDataEx COD_VALUES, 2, Trim(GetText(vasTest, vasTest.DataRowCnt, lCol))
    ChartFX1.OpenDataEx COD_XVALUES, 2, Trim(GetText(vasTest, vasTest.DataRowCnt, lCol))
    ChartFX1.Title(CHART_TOPTIT) = Trim(GetText(vasTest, 0, lCol - 1))
    
    ChartFX1.Axis(AXIS_Y).Max = CCur(YMax) + 1
    ChartFX1.Axis(AXIS_Y).Min = CCur(YMin) - 1
        
    ChartFX1.Axis(AXIS_X).Max = CCur(XMax)
    ChartFX1.Axis(AXIS_X).Min = CCur(XMin)
    
    For lRow = 2 To vasTest.DataRowCnt - 8
'        'First let's set the Y coordinates using the YValue Property
'        ChartFX1.Series(0).Yvalue(lRow - 1) = Trim(GetText(vasTest, lRow, lCol - 1))
'        'Now we send the X coordinates using the XValue Property
'        ChartFX1.Series(0).Xvalue(lRow - 1) = Trim(GetText(vasTest, lRow, lCol))
        
        'First let's set the Y coordinates using the YValue Property
        ChartFX1.Series(0).Yvalue(lRow - 1) = Trim(GetText(vasTest, lRow, lCol))
        'Now we send the X coordinates using the XValue Property
        ChartFX1.Series(0).Xvalue(lRow - 1) = Trim(GetText(vasTest, lRow, lCol - 1))
        
        X = Trim(GetText(vasTemp1, lRow, lCol))
        Y = Trim(GetText(vasTest, vasTest.DataRowCnt - 3, lCol)) * X + Trim(GetText(vasTest, vasTest.DataRowCnt - 2, lCol))
        ChartFX1.Series(1).Yvalue(lRow - 1) = Y
        ChartFX1.Series(1).Xvalue(lRow - 1) = X
        
    Next lRow
    
    ChartFX1.CloseData COD_XVALUES
    ChartFX1.CloseData COD_VALUES

    Exit Sub
ErrHandle:
    MsgBox "조회 도중 오류 발생!"
    Exit Sub
End Sub
