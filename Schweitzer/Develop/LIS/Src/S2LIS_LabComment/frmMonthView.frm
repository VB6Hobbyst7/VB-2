VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmMonthView 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "월별 보고현황"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDoct 
      Height          =   300
      Left            =   6585
      Style           =   2  '드롭다운 목록
      TabIndex        =   16
      Top             =   1110
      Width           =   1755
   End
   Begin VB.CommandButton cmdMonthCnt 
      BackColor       =   &H00FEFCFE&
      Caption         =   "월별집계"
      Height          =   495
      Left            =   8595
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   4950
      Width           =   1260
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   8610
      TabIndex        =   9
      Top             =   1440
      Width           =   1260
   End
   Begin VB.ListBox lstPtList 
      Appearance      =   0  '평면
      BackColor       =   &H00F1F5F5&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3990
      Left            =   4500
      TabIndex        =   7
      Top             =   1455
      Width           =   3810
   End
   Begin MedControls1.LisLabel lblTotCnt 
      Height          =   225
      Left            =   2985
      TabIndex        =   6
      Top             =   1155
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   397
      BackColor       =   12637910
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   2
      Caption         =   "100"
   End
   Begin VB.HScrollBar spnMonth 
      Height          =   255
      Left            =   3945
      Max             =   2
      TabIndex        =   4
      Top             =   1140
      Width           =   480
   End
   Begin FPSpread.vaSpread tblMonth 
      Height          =   4020
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
      _Version        =   196608
      _ExtentX        =   7435
      _ExtentY        =   7091
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      MaxCols         =   7
      MaxRows         =   12
      OperationMode   =   1
      Protect         =   0   'False
      ScrollBars      =   0
      ShadowColor     =   12648447
      ShadowDark      =   12648447
      ShadowText      =   0
      SpreadDesigner  =   "frmMonthView.frx":0000
      VisibleCols     =   7
      VisibleRows     =   10
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료(&X)"
      Height          =   495
      Left            =   8610
      TabIndex        =   1
      Top             =   1995
      Width           =   1260
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   285
      Left            =   2055
      TabIndex        =   5
      Top             =   1125
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   503
      BackColor       =   12637910
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "Total :          명"
      LeftGab         =   100
   End
   Begin VB.Frame fraMonthCnt 
      BackColor       =   &H00DBE6E6&
      Height          =   2115
      Left            =   240
      TabIndex        =   11
      Top             =   5925
      Visible         =   0   'False
      Width           =   9675
      Begin FPSpread.vaSpread tblMonthCnt 
         Height          =   1065
         Left            =   150
         TabIndex        =   12
         Top             =   840
         Width           =   9375
         _Version        =   196608
         _ExtentX        =   16536
         _ExtentY        =   1879
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         MaxRows         =   1
         OperationMode   =   1
         ScrollBars      =   0
         ShadowColor     =   16775924
         ShadowDark      =   16775924
         SpreadDesigner  =   "frmMonthView.frx":127B
      End
      Begin MedControls1.LisLabel lblYearCnt 
         Height          =   225
         Left            =   2895
         TabIndex        =   13
         Top             =   435
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         BackColor       =   12637910
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   2
         Caption         =   "100"
      End
      Begin MedControls1.LisLabel lbllabel 
         Height          =   285
         Left            =   1965
         TabIndex        =   14
         Top             =   405
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         BackColor       =   12637910
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "Total :          명"
         LeftGab         =   100
      End
      Begin VB.Label lblYear 
         BackStyle       =   0  '투명
         Caption         =   "1998년 12월"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0725F&
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   420
         Width           =   1485
      End
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  '투명
      Caption         =   "1998년 12월"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0725F&
      Height          =   255
      Left            =   4530
      TabIndex        =   8
      Top             =   1155
      Width           =   2460
   End
   Begin VB.Label lblMonth 
      BackStyle       =   0  '투명
      Caption         =   "1998년 12월"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0725F&
      Height          =   255
      Left            =   285
      TabIndex        =   3
      Top             =   1140
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   405
      Picture         =   "frmMonthView.frx":17D8
      Top             =   285
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "월별현황"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   390
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   630
      Index           =   0
      Left            =   195
      Shape           =   4  '둥근 사각형
      Top             =   210
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  '단색
      Height          =   630
      Index           =   1
      Left            =   255
      Shape           =   4  '둥근 사각형
      Top             =   270
      Width           =   2265
   End
End
Attribute VB_Name = "frmMonthView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PtntCnt(1 To 31) As Integer
Private Today As Date
Private OldRow As Integer
Private OldCol As Integer
Private OldBkColor As Long

Private Sub cboDoct_Click()
    Call cmdRefresh_Click
End Sub

Private Sub cmdExit_Click()
    'Set Me = Nothing
    Unload Me
    Set frmMonthView = Nothing
End Sub

Private Sub cmdMonthCnt_Click()
    Dim SqlStmt     As String
    Dim rs          As Recordset
    Dim strDoct     As String
    Dim strAllFg    As String
    
    strDoct = medGetP(cboDoct.Text, 1, Space(5))
    strAllFg = medGetP(cboDoct.Text, 2, Space(5))
    
    lblYear.Caption = Year(Today) & " 년"
    lblYearCnt.Caption = ""
    
    SqlStmt = " select rptdt, count(*) as YearCnt from " & T_LAB501 & _
              " where " & DBW("rptdt like ", Format(Today, "YYYY%"))
              
    If strAllFg <> "" Then
        SqlStmt = SqlStmt & _
              " and   " & DBW("rptid  = ", strDoct)
    End If

    SqlStmt = SqlStmt & _
              " and   " & DBW("donefg = ", enStsCd.StsCd_LIS_Accession) & _
              " group by rptdt "
              
'        " and   " & DBW("rptid  = ", objDoctor.DoctId)
'    Set Rs = OpenRecordSet(SqlStmt)
    Set rs = New Recordset
    rs.Open SqlStmt, DBConn
    
    lblYearCnt.Caption = ""
    
    tblMonthCnt.Row = 1: tblMonthCnt.Row2 = 1
    tblMonthCnt.Col = 1: tblMonthCnt.Col2 = tblMonthCnt.MaxCols
    tblMonthCnt.BlockMode = True
    tblMonthCnt.Value = ""
    tblMonthCnt.BlockMode = False
    
    tblMonthCnt.Row = 1
    While (Not rs.EOF)
        tblMonthCnt.Col = Val(Mid("" & rs.Fields("RptDt").Value, 5, 2))
        tblMonthCnt.Value = Val(tblMonthCnt.Value) + Val("" & rs.Fields("YearCnt").Value)
        lblYearCnt.Caption = Val(lblYearCnt.Caption) + Val("" & rs.Fields("YearCnt").Value)
        rs.MoveNext
    Wend
    
'    Rs.RsClose
    Set rs = Nothing
    
    fraMonthCnt.Visible = True
End Sub

Private Sub cmdRefresh_Click()
    
    Call GetCalendar(Today)
    Call GetCount(Today)
    Call cmdMonthCnt_Click
    
End Sub

Private Sub Form_Load()
    Dim strFileNm As String
    Dim strTitle As String
    Dim WeekDayKor As String
    
    Me.Left = 4845
    Me.Top = 0
    
    '-- 전문의(Supervisor) Set
    cboDoct.Clear
    
    cboDoct.AddItem "[전체]"
    
    Call SetSupervisor
    
    If cboDoct.ListCount > 0 Then
        cboDoct.ListIndex = 0
    End If
    
    Erase PtntCnt
    
    Today = Format(Now, "YY-MM-DD")
    
    Call GetCalendar(Today)
    Call DisplayDate(Today)
    Call FindDay(Today)
    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", _
                          "화요일", "수요일", "목요일", "금요일", "토요일")
    
    spnMonth.Value = 1
    
End Sub

Private Sub SetSupervisor()
    Dim SSQL        As String
    Dim rs          As New ADODB.Recordset
    Dim strID       As String
    Dim strNm       As String
    
    '-- Group ID 는 일단 Fix ㅡ.ㅡ ...
    SSQL = " select a.empid, b.empnm " & _
           "   from " & T_COM010 & " a, " & T_COM006 & " b " & _
           "  where a.groupid = 'G004' " & _
           "    and b.empid = a.empid " & _
           "  order by a.empid "
                          
    rs.Open SSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.BOF = False Then
        
        Do Until rs.EOF = True
            
            strID = rs.Fields("empid").Value & ""
            strNm = rs.Fields("empnm").Value & ""
            
            cboDoct.AddItem strID & Space(5) & strNm
            
            rs.MoveNext
        Loop
    End If
                  
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub GetCount(ByVal pToday As Date)

    Dim SqlStmt As String
    Dim rs As Recordset
    Dim ThisDay As Integer
    Dim i As Integer
    Dim strDoct     As String
    Dim strAllFg    As String
    
    strDoct = medGetP(cboDoct.Text, 1, Space(5))
    strAllFg = medGetP(cboDoct.Text, 2, Space(5))
    
    SqlStmt = " select rptdt, count(*) as PtCnt from " & T_LAB501 & _
              " where " & DBW("rptdt like ", Format(pToday, "YYYYMM%"))
              
    If strAllFg <> "" Then
        SqlStmt = SqlStmt & _
              " and   " & DBW("rptid  = ", strDoct)
    End If
    
    SqlStmt = SqlStmt & _
              " and   " & DBW("donefg = ", enStsCd.StsCd_LIS_Accession) & _
              " group by rptdt"
              
'    " and   " & DBW("rptid  = ", objDoctor.DoctId) & _

'    Set Rs = OpenRecordSet(SqlStmt)
    Set rs = New Recordset
    rs.Open SqlStmt, DBConn
    
    lstPtList.Clear
    
    Erase PtntCnt

    While (Not rs.EOF)
        
        ThisDay = Val(Mid("" & rs.Fields("RptDt").Value, 7, 2))
        PtntCnt(ThisDay) = Val("" & rs.Fields("PtCnt").Value)
        
        rs.MoveNext
        
    Wend
    
'    Rs.RsClose
    Set rs = Nothing
    
    lblTotCnt.Caption = ""
    For i = 1 To UBound(PtntCnt)
        lblTotCnt.Caption = Val(lblTotCnt.Caption) + PtntCnt(i)
    Next
    
End Sub

Private Sub lstPtList_DblClick()
    With frmReport
        .Show
        .ZOrder 0
        DoEvents
        .ptid = Trim(medGetP(lstPtList.List(lstPtList.ListIndex), 1, vbTab))
        .BedinDt = Trim(medGetP(lstPtList.List(lstPtList.ListIndex), 4, vbTab))
        Call .StartQuery
        .cmdSave.Enabled = False
    End With
End Sub

Private Sub spnMonth_Change()

    With tblMonth
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .BackColor = vbWhite
        .BlockMode = False
    End With
   
    If spnMonth.Value = 2 Then
        Call spnMonth_SpinUp
    ElseIf spnMonth.Value = 0 Then
        Call spnMonth_SpinDown
    End If
    spnMonth.Value = 1
    lblDate.Caption = ""
   
    Call FindDay(Today)
    fraMonthCnt.Visible = False
   
End Sub

Private Sub spnMonth_SpinUp()
    
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")

    ThisMonth = Month(Today)
    Today = DateAdd("m", 1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    
    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If

End Sub

Private Sub spnMonth_SpinDown()
    
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")

    ThisMonth = Month(Today)
    Today = DateAdd("m", -1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    
    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If
 
End Sub

'Private Sub InitRtn()
'Dim strFileNm As String
'Dim InputData As String

    'strFileNm = App.Path & "\" & gUser & ".txt"
    'If SSTab1.Tab = 2 Then
    '    cmdAdd.Visible = False
    '    txtNote.Text = ""
    '    If Dir(strFileNm) <> "" Then
    '        Open strFileNm For Input As #1
    '        Do While Not EOF(1)   ' 파일의 끝을 확인합니다.
    '            Line Input #1, InputData   ' 데이터 행을 읽어 들입니다.
    '            txtNote.Text = txtNote.Text & InputData & Chr(13) & Chr(10)
    '        Loop
    '        Close #1
    '    End If
    'Else
    '    cmdAdd.Visible = True
    '    If SSTab1.Tab = 1 Then
    '        Call GetSchedule(calCalendar.Value)
    '    End If
    '    If PreviousTab = 2 Then
    '        Open strFileNm For Output As #1
    '        Print #1, txtNote.Text
    '        Close #1
    '    End If
    'End If

'End Sub

Private Sub tblMonth_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim tmpDay As String
    Dim lngRow As Long

    If Row = 0 Then Exit Sub
    
    If OldRow > 0 And OldCol > 0 Then
        tblMonth.Col = OldCol
        tblMonth.Row = OldRow
        tblMonth.BackColor = OldBkColor
        tblMonth.Row = OldRow + 1
        tblMonth.BackColor = OldBkColor
    End If
    
    If Row Mod 2 = 0 Then
        lngRow = Row - 1
    Else
        lngRow = Row
    End If
    
    tblMonth.Col = Col
    
    tblMonth.Row = lngRow
    
    OldRow = lngRow
    OldCol = Col
    OldBkColor = tblMonth.BackColor
    
    tblMonth.BackColor = &HC0FFFF       'dcm_LIGHTYELLOW
    
    tblMonth.Row = lngRow + 1
    tblMonth.BackColor = &HC0FFFF
    
    tblMonth.Row = lngRow
    tmpDay = tblMonth.Value
    lblDate.Caption = lblMonth.Caption & " " & tmpDay & "일"
    
    Call GetPtList(Format(lblDate.Caption, CS_DateDbFormat))

End Sub

Private Sub tblMonth_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim tmpDay As String

    'If Row Mod 2 = 0 Then
    '    tblMonth.Row = Row - 1
    'Else
    '    tblMonth.Row = Row
    'End If
    'tblMonth.Col = Col
    'tmpDay = tblMonth.Value
    'calCalendar.Day = Val(tmpDay)
    'Call calCalendar_Click
    
    'cmdAdd_Click
End Sub


Private Sub GetCalendar(ByVal datToday As Date)
Dim tmpDate As Date
Dim ThisMonth As Integer
Dim ThisDay As Integer
Dim FirstWeekDay As Integer
Dim i As Integer
    
'    With tblMonth
'        .Row = 1: .Row2 = .MaxRows
'        .Col = 1: .Col2 = .MaxCols
'        .BlockMode = True
'        .BackColor = vbWhite
'        .BlockMode = False
'    End With
    
    Call GetCount(datToday)
    
    tblMonth.Row = -1
    tblMonth.Col = -1
    tblMonth.BlockMode = True
    tblMonth.Action = 12   'Text Clear
    tblMonth.BlockMode = False
    
    tmpDate = CDate(Format(datToday, "YYYY-MM-") & "01")
    ThisMonth = Month(tmpDate)
    FirstWeekDay = Weekday(tmpDate)
    
    Do While Month(tmpDate) = ThisMonth
        ThisDay = Day(tmpDate)
        tblMonth.Row = (((ThisDay + FirstWeekDay - 2) \ 7) * 2) + 1
        tblMonth.Col = (ThisDay + FirstWeekDay - 2) Mod 7 + 1
        If tblMonth.Col = 1 Then
            tblMonth.ForeColor = &HFF&
        Else
            tblMonth.ForeColor = &H0&
        End If
        'tblMonth.BackColor = &HFFFFFF
        tblMonth.TypeHAlign = TypeHAlignLeft
        tblMonth.TypeVAlign = TypeVAlignTop
        tblMonth.Value = ThisDay
        tblMonth.Row = tblMonth.Row + 1
        tblMonth.TypeHAlign = TypeHAlignCenter
        tblMonth.TypeVAlign = TypeVAlignBottom
        'tblMonth.BackColor = &HFFFFFF
        
        If PtntCnt(ThisDay) <> 0 Then
            tblMonth.Text = PtntCnt(ThisDay)
        End If
        
        tmpDate = DateAdd("d", 1, tmpDate)
    Loop
    
    OldRow = -1
    OldCol = -1
    
End Sub

Private Sub DisplayDate(ByVal datToday As Date)
Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(datToday), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")
    lblMonth.Caption = Format(datToday, "YYYY년 MM월")
    lblDate.Caption = Format(datToday, "YYYY년 MM월 DD일 ") & WeekDayKor
    
End Sub

'Sub FindDay(ByVal OldDate As Date, ByVal NewDate As Date)
Private Sub FindDay(ByVal NewDate As Date)
Dim tmpDate As Date
Dim ThisDay As Integer
Dim FirstWeekDay As Integer

    If Month(NewDate) <> Month(Now) Then Exit Sub
    
    tmpDate = CDate(Format(NewDate, "YYYY-MM-") & "01")
    FirstWeekDay = Weekday(tmpDate)
    ThisDay = Day(NewDate)
    
    tblMonth.Col = (ThisDay + FirstWeekDay - 2) Mod 7 + 1
    tblMonth.Row = (((ThisDay + FirstWeekDay - 2) \ 7) * 2) + 1
    tblMonth.BackColor = &HFFFFC0
    tblMonth.Row = tblMonth.Row + 1
    tblMonth.BackColor = &HFFFFC0

End Sub

Private Sub GetPtList(ByVal pDate As String)
    Dim SqlStmt As String
    Dim rs As Recordset
    Dim strTmp As String
    Dim strDoct     As String
    Dim strAllFg    As String
    
    strDoct = medGetP(cboDoct.Text, 1, Space(5))
    strAllFg = medGetP(cboDoct.Text, 2, Space(5))
    
    SqlStmt = " select a.*, b." & F_PTNM & " as PtNm " & _
              " from   " & T_LAB501 & " a, " & T_HIS001 & " b " & _
              " where  " & DBW("a.rptdt = ", pDate)

    If strAllFg <> "" Then
        SqlStmt = SqlStmt & _
              " and   " & DBW("a.rptid  = ", strDoct)
    End If
    
    SqlStmt = SqlStmt & _
              " and    " & DBW("a.donefg = ", enStsCd.StsCd_LIS_Accession) & _
              " and   b." & F_PTID & " = a.ptid " & _
              " order by a.ptid"
              
'              " and    " & DBW("a.rptid = ", objDoctor.DoctId) & _
'    Set Rs = OpenRecordSet(SqlStmt)
    Set rs = New Recordset
    rs.Open SqlStmt, DBConn
    
    lstPtList.Clear
    While (Not rs.EOF)
        strTmp = Format(Trim("" & rs.Fields("PtId").Value), "@@@@@@@@@@") & vbTab
        strTmp = strTmp & Trim("" & rs.Fields("PtNm").Value) & vbTab
        strTmp = strTmp & Trim("" & rs.Fields("WardId").Value) & "-"
        strTmp = strTmp & Trim("" & rs.Fields("HosilId").Value) & vbTab
        strTmp = strTmp & Trim("" & rs.Fields("bedindt").Value) & vbTab
        lstPtList.AddItem strTmp
        
        rs.MoveNext
    Wend
    
'    Rs.RsClose
    Set rs = Nothing
    
End Sub
