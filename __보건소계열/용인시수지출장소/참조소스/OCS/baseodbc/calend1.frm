VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form FrmCalendar 
   BorderStyle     =   3  '고정 대화 상자
   Caption         =   "Calendar"
   ClientHeight    =   2010
   ClientLeft      =   4260
   ClientTop       =   3420
   ClientWidth     =   3480
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2010
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread SS 
      Height          =   1695
      Left            =   0
      OleObjectBlob   =   "CALEND1.frx":0000
      TabIndex        =   0
      Top             =   315
      Width           =   3480
   End
   Begin VB.ComboBox ComboMM 
      Height          =   315
      Left            =   1320
      Style           =   2  '늘어진 목록
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   585
   End
   Begin VB.ComboBox ComboYY 
      Height          =   315
      Left            =   120
      Style           =   2  '늘어진 목록
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label 
      Caption         =   "월"
      Height          =   195
      Index           =   1
      Left            =   1950
      TabIndex        =   4
      Top             =   60
      Width           =   195
   End
   Begin VB.Label Label 
      Caption         =   "년"
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   60
      Width           =   195
   End
End
Attribute VB_Name = "FrmCalendar"
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
    Dim I                   As Integer
    
    ComboYY.ListIndex = Val(Format$(ArgDate, "YYYY")) - START_YEAR
    ComboMM.ListIndex = Val(Format$(ArgDate, "MM")) - 1
    strFirstDate = Format$(ArgDate, "YYYY-MM-01")
    nBeforeMonthLast = Val(Format$(DateAdd("d", -1, strFirstDate), "DD"))
    nCurrentMonthLast = Val(Format$(DateAdd("d", -1, DateAdd("m", 1, strFirstDate)), "DD"))
    nStartCol = Format(strFirstDate, "w")
    If nStartCol = 1 Then nStartCol = 8
    
    Ss.ReDraw = False
    Call Spread_Clear
    GoSub Spread_Display
    Ss.ReDraw = True
    
Exit Sub

'/--------------------------------------------------------------------/

Spread_Display:

    ' 전월 Display
    Ss.Row = 1
    Ss.Col = 1
    For I = 1 To nStartCol - 1
        Ss.Text = Trim(Str(nBeforeMonthLast - nStartCol + 1 + I))
        Ss.BackColor = RGB(255, 255, 192)
        Ss.ForeColor = RGB(192, 192, 192)
        If Ss.Col = 7 Then
            If Ss.Row < 6 Then
                Ss.Row = Ss.Row + 1
                Ss.Col = 1
            Else
                Exit For
            End If
        Else
            Ss.Col = Ss.Col + 1
        End If
    Next I
    
    ' 당월 Display
    For I = 1 To nCurrentMonthLast
        Ss.Text = Trim(Str(I))
        ' 당일
        'If CVDate(Format$(ArgDate, "YYYY-MM-") & Format$(I, "00")) = Date Then SS.BackColor = RGB(0, 255, 0)
        ' 공휴일
        'If Is_HolyDay(CVDate(Format$(ArgDate, "YYYY-MM-") & Format$(I, "00"))) Then SS.ForeColor = RGB(255, 0, 0)
        If Ss.Col = 7 Then
            If Ss.Row < 6 Then
                Ss.Row = Ss.Row + 1
                Ss.Col = 1
            Else
                Exit For
            End If
        Else
            Ss.Col = Ss.Col + 1
        End If
    Next I
    
    ' 다음월 Display
    For I = 1 To 14
        Ss.Text = Trim(Str(I))
        Ss.BackColor = RGB(255, 255, 192)
        Ss.ForeColor = RGB(192, 192, 192)
        If Ss.Col = 7 Then
            If Ss.Row < 6 Then
                Ss.Row = Ss.Row + 1
                Ss.Col = 1
            Else
                Exit For
            End If
        Else
            Ss.Col = Ss.Col + 1
        End If
    Next I

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
    Ss.Col = nCol
    Ss.Row = nRow
    Ss.Action = 0       'SS_ACTION_ACTIVE_CELL
    Call Spread_Select(nRow, nCol)

End Sub

Private Sub Spread_Clear()

    If Ss.ReDraw = True Then
        Ss.ReDraw = False
        GoSub Clearing
        Ss.ReDraw = True
    Else
        GoSub Clearing
    End If
    
Exit Sub

'/-----------------------------------------------------/

Clearing:

    Ss.Col = 1:     Ss.Col2 = Ss.MaxCols
    Ss.Row = 1:     Ss.Row2 = Ss.MaxRows
    Ss.BlockMode = True
    Ss.Text = ""
    Ss.BackColor = RGB(255, 255, 0)
    Ss.ForeColor = RGB(0, 0, 0)
    Ss.BlockMode = False
    
    Ss.Col = 1
    Ss.Row = -1
    Ss.ForeColor = RGB(255, 0, 0)
    
    Ss.Col = 7
    Ss.Row = -1
    Ss.ForeColor = RGB(0, 0, 255)
    
    Return

End Sub


Private Sub Spread_Initialize()

    Dim I           As Integer
    Dim StrYY       As String * 4
    Dim StrMM       As String * 2
    
    ComboYY.Clear
    For I = START_YEAR To LAST_YEAR
        RSet StrYY = Trim(Str(I))
        ComboYY.AddItem StrYY
    Next I
    
    ComboMM.Clear
    For I = 1 To 12
        RSet StrMM = Trim(Str(I))
        ComboMM.AddItem StrMM
    Next I
    
    Ss.Row = 0
    Ss.Col = 1:     Ss.Text = "Sun"
    Ss.Col = 2:     Ss.Text = "Mon"
    Ss.Col = 3:     Ss.Text = "Tue"
    Ss.Col = 4:     Ss.Text = "Wed"
    Ss.Col = 5:     Ss.Text = "Thu"
    Ss.Col = 6:     Ss.Text = "Fri"
    Ss.Col = 7:     Ss.Text = "Sat"

End Sub


Private Sub Spread_Select(ByVal Row As Integer, ByVal Col As Integer)
    
    Dim strFirstDate    As String
    
    Ss.Col = Col:     FnActiveCol = Col
    Ss.Row = Row:     FnActiveRow = Row
    FnOldBackColor = Ss.BackColor
    Ss.BackColor = RGB(255, 255, 255)
    
    strFirstDate = Format(FdSelectedDate, "YYYY-MM") & "-01"
    FdSelectedDate = DateAdd("d", Val(Ss.Text) - 1, strFirstDate)
    Me.Caption = Format(FdSelectedDate, "Long Date")

End Sub


Private Sub Spread_UnSelect()

    Ss.Col = FnActiveCol
    Ss.Row = FnActiveRow
    Ss.BackColor = FnOldBackColor

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

Private Sub SS_DblClick(ByVal Col As Long, ByVal Row As Long)

    If (Col < 1) Or (Row < 1) Then Exit Sub

    Me.Tag = Format(FdSelectedDate, "YYYY-MM-DD")
    Me.Hide
    
End Sub

Private Sub SS_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim nCol            As Integer
    Dim nRow            As Integer
    
    If KeyCode = 37 And FnActiveCol = 1 Then
        KeyCode = 0
        If FnActiveRow = 1 Then Exit Sub
        nCol = 7
        nRow = FnActiveRow - 1
        Call SS_LeaveCell(FnActiveCol + 0, FnActiveRow + 0, nCol + 0, nRow + 0, False)
        Ss.Action = 0       'SS_ACTION_ACTIVE_CELL
    ElseIf KeyCode = 39 And FnActiveCol = 7 Then
        KeyCode = 0
        If FnActiveRow = 7 Then Exit Sub
        nCol = 1
        nRow = FnActiveRow + 1
        Call SS_LeaveCell(FnActiveCol + 0, FnActiveRow + 0, nCol + 0, nRow + 0, False)
        Ss.Action = 0       'SS_ACTION_ACTIVE_CELL
    End If

End Sub


Private Sub SS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.Tag = Format(FdSelectedDate, "YYYY-MM-DD")
        Me.Hide
    ElseIf KeyAscii = 27 Then
        Me.Hide
    End If

End Sub

Private Sub SS_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    If (Col < 1) Or (Row < 1) Then Exit Sub
    If (NewCol < 1) Or (NewRow < 1) Then Exit Sub
    
    Ss.Col = NewCol
    Ss.Row = NewRow
    If Ss.BackColor = RGB(255, 255, 192) Then
        If NewRow < 2 Then
            If FdSelectedDate < CVDate(Format(START_YEAR, "###0") & "-02-01") Then
                Cancel = True
                Exit Sub
            End If
            FdSelectedDate = DateAdd("m", -1, FdSelectedDate)
            FdSelectedDate = Format(FdSelectedDate, "YYYY-MM") & "-" & Format(Val(Ss.Text), "00")
            Call Display_Month(FdSelectedDate)
            Call Select_Current_Date(FdSelectedDate)
        Else
            If FdSelectedDate >= CVDate(Format(LAST_YEAR, "###0") & "-12-01") Then
                Cancel = True
                Exit Sub
            End If
            FdSelectedDate = DateAdd("m", 1, FdSelectedDate)
            FdSelectedDate = Format(FdSelectedDate, "YYYY-MM") & "-" & Format(Val(Ss.Text), "00")
            Call Display_Month(FdSelectedDate)
            Call Select_Current_Date(FdSelectedDate)
        End If
    Else
        Call Spread_UnSelect
        Call Spread_Select(NewRow, NewCol)
    End If
    
End Sub


