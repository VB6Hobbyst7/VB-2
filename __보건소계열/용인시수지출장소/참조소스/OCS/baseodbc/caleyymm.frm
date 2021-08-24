VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form FrmCalendarYYMM 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "YYMM"
   ClientHeight    =   720
   ClientLeft      =   4245
   ClientTop       =   4800
   ClientWidth     =   3480
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   720
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand CmdOk 
      Height          =   390
      Left            =   2700
      TabIndex        =   6
      Top             =   150
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   16711680
      Font3D          =   1
   End
   Begin VB.ComboBox ComboMM 
      Height          =   360
      Left            =   1545
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   150
      Width           =   660
   End
   Begin VB.ComboBox ComboYY 
      Height          =   360
      Left            =   120
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   1020
   End
   Begin VB.PictureBox SS 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   3420
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3480
      Begin FPSpread.vaSpread SS1 
         Height          =   390
         Left            =   0
         OleObjectBlob   =   "Caleyymm.frx":0000
         TabIndex        =   5
         Top             =   0
         Width           =   3465
      End
   End
   Begin VB.Label Label 
      Caption         =   "월"
      Height          =   270
      Index           =   1
      Left            =   2250
      TabIndex        =   4
      Top             =   225
      Width           =   195
   End
   Begin VB.Label Label 
      Caption         =   "년"
      Height          =   270
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   225
      Width           =   195
   End
End
Attribute VB_Name = "FrmCalendarYYMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim FnActiveCol             As Integer
Dim FnActiveRow             As Integer

Dim FnOldBackColor          As Long
Dim FdSelectedDate          As Date

Private Const START_YEAR = 1900
Private Const LAST_YEAR = 2100

Private Sub Display_Month(ByVal ArgDate As Date)

    Dim strFirstDate        As String
    Dim nStartCol           As Integer
    Dim nBeforeMonthLast    As Integer
    Dim nCurrentMonthLast   As Integer
    Dim i                   As Integer
    
    ComboYY.ListIndex = Val(Format$(ArgDate, "YYYY")) - START_YEAR
    ComboMM.ListIndex = Val(Format$(ArgDate, "MM")) - 1
    strFirstDate = Format$(ArgDate, "YYYY-MM-01")
    nBeforeMonthLast = Val(Format$(DateAdd("d", -1, strFirstDate), "DD"))
    nCurrentMonthLast = Val(Format$(DateAdd("d", -1, DateAdd("m", 1, strFirstDate)), "DD"))
    nStartCol = Format(strFirstDate, "w")
    If nStartCol = 1 Then nStartCol = 8
    
    SS1.Redraw = False
    Call Spread_Clear
    GoSub Spread_Display
    SS1.Redraw = True
    
Exit Sub

'/--------------------------------------------------------------------/

Spread_Display:

    ' 전월 Display
    SS1.Row = 1
    SS1.Col = 1
    For i = 1 To nStartCol - 1
        SS1.Text = Trim(Str(nBeforeMonthLast - nStartCol + 1 + i))
        SS1.BackColor = RGB(255, 255, 192)
        SS1.ForeColor = RGB(192, 192, 192)
        If SS1.Col = 7 Then
            If SS1.Row < 6 Then
                SS1.Row = SS1.Row + 1
                SS1.Col = 1
            Else
                Exit For
            End If
        Else
            SS1.Col = SS1.Col + 1
        End If
    Next i
    
    ' 당월 Display
    For i = 1 To nCurrentMonthLast
        SS1.Text = Trim(Str(i))
        ' 당일
        'If CVDate(Format$(ArgDate, "YYYY-MM-") & Format$(I, "00")) = Date Then SS.BackColor = RGB(0, 255, 0)
        ' 공휴일
        'If Is_HolyDay(CVDate(Format$(ArgDate, "YYYY-MM-") & Format$(I, "00"))) Then SS.ForeColor = RGB(255, 0, 0)
        If SS1.Col = 7 Then
            If SS1.Row < 6 Then
                SS1.Row = SS1.Row + 1
                SS1.Col = 1
            Else
                Exit For
            End If
        Else
            SS1.Col = SS1.Col + 1
        End If
    Next i
    
    ' 다음월 Display
    For i = 1 To 14
        SS1.Text = Trim(Str(i))
        SS1.BackColor = RGB(255, 255, 192)
        SS1.ForeColor = RGB(192, 192, 192)
        If SS1.Col = 7 Then
            If SS1.Row < 6 Then
                SS1.Row = SS1.Row + 1
                SS1.Col = 1
            Else
                Exit For
            End If
        Else
            SS1.Col = SS1.Col + 1
        End If
    Next i

    Return

End Sub

Private Function Is_HolyDay(ByVal ArgDate As Date) As Integer

    Is_HolyDay = True
    
    Select Case Format(ArgDate, "MMDD")
        Case "0101" To "0102"
        Case "0301"
        Case "0405"
        Case "0505"
        Case "0606"
        Case "0717"
        Case "0815"
        Case "1003"
        Case "1225"
        Case Else
            Is_HolyDay = False
    End Select
    
End Function


Private Sub Select_Current_Date(ByVal ArgDate As Date)

    Dim nCol            As Integer
    Dim nRow            As Integer
    Dim nTemp           As Integer
    Dim strFirstDate    As String
    
    strFirstDate = Format(ArgDate, "YYYY-MM") & "-01"
    nTemp = Val(Format(strFirstDate, "w"))
    If nTemp = 1 Then nTemp = 8
    nTemp = nTemp + Val(Format(ArgDate, "DD")) - 1
    nRow = (nTemp + 6) \ 7
    nCol = Val(Format(ArgDate, "w"))
    SS1.Col = nCol
    SS1.Row = nRow
    SS1.Action = 0       'SS_ACTION_ACTIVE_CELL
    Call Spread_Select(nRow, nCol)

End Sub

Private Sub Spread_Clear()

    If SS1.Redraw = True Then
        SS1.Redraw = False
        GoSub Clearing
        SS1.Redraw = True
    Else
        GoSub Clearing
    End If
    
Exit Sub

'/-----------------------------------------------------/

Clearing:

    SS1.Col = 1:     SS1.Col2 = SS1.MaxCols
    SS1.Row = 1:     SS1.Row2 = SS1.MaxRows
    SS1.BlockMode = True
    SS1.Text = ""
    SS1.BackColor = RGB(255, 255, 0)
    SS1.ForeColor = RGB(0, 0, 0)
    SS1.BlockMode = False
    
    SS1.Col = 1
    SS1.Row = -1
    SS1.ForeColor = RGB(255, 0, 0)
    
    SS1.Col = 7
    SS1.Row = -1
    SS1.ForeColor = RGB(0, 0, 255)
    
    Return

End Sub


Private Sub Spread_Initialize()

    Dim i           As Integer
    Dim StrYY       As String * 4
    Dim StrMM       As String * 2
    
    ComboYY.Clear
    For i = START_YEAR To LAST_YEAR
        RSet StrYY = Trim(Str(i))
        ComboYY.AddItem StrYY
    Next i
    
    ComboMM.Clear
    For i = 1 To 12
        RSet StrMM = Trim(Str(i))
        ComboMM.AddItem StrMM
    Next i
    
    SS1.Row = 0
    SS1.Col = 1:     SS1.Text = "Sun"
    SS1.Col = 2:     SS1.Text = "Mon"
    SS1.Col = 3:     SS1.Text = "Tue"
    SS1.Col = 4:     SS1.Text = "Wed"
    SS1.Col = 5:     SS1.Text = "Thu"
    SS1.Col = 6:     SS1.Text = "Fri"
    SS1.Col = 7:     SS1.Text = "Sat"

End Sub


Private Sub Spread_Select(ByVal Row As Integer, ByVal Col As Integer)
    
    Dim strFirstDate    As String
    
    SS1.Col = Col:     FnActiveCol = Col
    SS1.Row = Row:     FnActiveRow = Row
    FnOldBackColor = SS1.BackColor
    SS1.BackColor = RGB(255, 255, 255)
    
    strFirstDate = Format(FdSelectedDate, "YYYY-MM") & "-01"
    FdSelectedDate = DateAdd("d", Val(SS1.Text) - 1, strFirstDate)
    Me.Caption = Format(FdSelectedDate, "YYYY 년  MM 월")   'Long Date

End Sub


Private Sub Spread_UnSelect()

    SS1.Col = FnActiveCol
    SS1.Row = FnActiveRow
    SS1.BackColor = FnOldBackColor

End Sub


Private Sub CmdOk_Click()

    Me.Tag = Format$(Left$(Trim(ComboYY.Text), 4) & "-" & Left$(Trim(ComboMM.Text), 2), "YYYY-MM-DD")
    
    'SendKeys "{tab}"
    'Me.Hide
    Unload Me
    
    SendKeys "{tab}"
    
End Sub

Private Sub ComboMM_Click()

    Dim StrDate     As String
    
    Do      ' 3월31일이 선택 됐을때 2월로 변환했을 경우 2월31일이 아닌 2월28(29)일로 선택
        StrDate = Format(FdSelectedDate, "YYYY-") & Format(ComboMM.Text, "00") & Format(FdSelectedDate, "-DD")
        FdSelectedDate = FdSelectedDate - 1
    Loop Until IsDate(StrDate)
    
    FdSelectedDate = CVDate(StrDate)
    
    Call Display_Month(FdSelectedDate)
    Call Select_Current_Date(FdSelectedDate)
    
    SendKeys "{Tab}"

End Sub

Private Sub ComboMM_DropDown()

    DoEvents
    
End Sub


Private Sub ComboMM_KeyPress(KeyAscii As Integer)

    SendKeys "{Tab}"
    
End Sub

Private Sub ComboYY_Click()

    Dim StrDate     As String
    
    Do      ' 2월29일이 선택 됐을때 평년로 변환했을 경우 2월29일이 아닌 2월28일로 선택
        StrDate = Format(ComboYY.Text, "###0") & Format(FdSelectedDate, "-MM-DD")
        FdSelectedDate = FdSelectedDate - 1
    Loop Until IsDate(StrDate)
    
    FdSelectedDate = CVDate(StrDate)
    
    Call Display_Month(FdSelectedDate)
    Call Select_Current_Date(FdSelectedDate)
    
    SendKeys "{Tab}"

End Sub


Private Sub ComboYY_DropDown()

    DoEvents
    
End Sub


Private Sub ComboYY_KeyPress(KeyAscii As Integer)

    SendKeys "{Tab}"

End Sub

Private Sub Form_Activate()

    If IsDate(Me.Tag) Then
        FdSelectedDate = CVDate(Me.Tag)
    Else
        FdSelectedDate = Date
    End If
    Call Display_Month(FdSelectedDate)
    Call Select_Current_Date(FdSelectedDate)

End Sub

Private Sub Form_Load()
    
    Call Spread_Initialize
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = True
    Me.Hide
    
End Sub

Private Sub SS1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If (Col < 1) Or (Row < 1) Then Exit Sub

    Me.Tag = Format(FdSelectedDate, "YYYY-MM-DD")
    
    'SendKeys "{tab}"
    'Me.Hide
    Unload Me
    
    SendKeys "{tab}"
    
End Sub

Private Sub SS1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim nCol            As Integer
    Dim nRow            As Integer
    
    If KeyCode = 37 And FnActiveCol = 1 Then
        KeyCode = 0
        If FnActiveRow = 1 Then Exit Sub
        nCol = 7
        nRow = FnActiveRow - 1
        Call SS1_LeaveCell(FnActiveCol + 0, FnActiveRow + 0, nCol + 0, nRow + 0, False)
        SS1.Action = 0       'SS_ACTION_ACTIVE_CELL
    ElseIf KeyCode = 39 And FnActiveCol = 7 Then
        KeyCode = 0
        If FnActiveRow = 7 Then Exit Sub
        nCol = 1
        nRow = FnActiveRow + 1
        Call SS1_LeaveCell(FnActiveCol + 0, FnActiveRow + 0, nCol + 0, nRow + 0, False)
        SS1.Action = 0       'SS_ACTION_ACTIVE_CELL
    End If


End Sub


Private Sub SS1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.Tag = Format(FdSelectedDate, "YYYY-MM-DD")
        Me.Hide
    ElseIf KeyAscii = 27 Then
        Me.Hide
    End If


End Sub


Private Sub SS1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    If (Col < 1) Or (Row < 1) Then Exit Sub
    If (NewCol < 1) Or (NewRow < 1) Then Exit Sub
    
    SS1.Col = NewCol
    SS1.Row = NewRow
    If SS1.BackColor = RGB(255, 255, 192) Then
        If NewRow < 2 Then
            If FdSelectedDate < CVDate(Format(START_YEAR, "###0") & "-02-01") Then
                Cancel = True
                Exit Sub
            End If
            FdSelectedDate = DateAdd("m", -1, FdSelectedDate)
            FdSelectedDate = Format(FdSelectedDate, "YYYY-MM") & "-" & Format(Val(SS1.Text), "00")
            Call Display_Month(FdSelectedDate)
            Call Select_Current_Date(FdSelectedDate)
        Else
            If FdSelectedDate >= CVDate(Format(LAST_YEAR, "###0") & "-12-01") Then
                Cancel = True
                Exit Sub
            End If
            FdSelectedDate = DateAdd("m", 1, FdSelectedDate)
            FdSelectedDate = Format(FdSelectedDate, "YYYY-MM") & "-" & Format(Val(SS1.Text), "00")
            Call Display_Month(FdSelectedDate)
            Call Select_Current_Date(FdSelectedDate)
        End If
    Else
        Call Spread_UnSelect
        Call Spread_Select(NewRow, NewCol)
    End If
    
End Sub


