VERSION 4.00
Begin VB.Form FrmCalendar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar"
   ClientHeight    =   2010
   ClientLeft      =   4260
   ClientTop       =   3420
   ClientWidth     =   3480
   BeginProperty Font 
      name            =   "굴림"
      charset         =   1
      weight          =   400
      size            =   9.75
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   2415
   Left            =   4200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Top             =   3075
   Width           =   3600
   Begin VB.ComboBox ComboMM 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   585
   End
   Begin VB.ComboBox ComboYY 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   795
   End
   Begin VBX.SpreadSheet SS 
      AutoSize        =   -1  'True
      DisplayRowHeaders=   0   'False
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "굴림"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1695
      InterfaceDesigner=   "CALENDAR.frx":0000
      Left            =   0
      MaxCols         =   7
      MaxRows         =   6
      ScrollBars      =   0  'None
      SelectBlockOptions=   0
      TabIndex        =   0
      Top             =   315
      UserResize      =   0
      Width           =   3480
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
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'/----------------------------------/
'/ Visual Basic 4.0 - 16bit 용 달력 /
'/         만든이 : 채화용          /
'/           Version 1.1            /
'/----------------------------------/

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
    
    SS.ReDraw = False
    Call Spread_Clear
    GoSub Spread_Display
    SS.ReDraw = True
    
Exit Sub

'/--------------------------------------------------------------------/

Spread_Display:

    ' 전월 Display
    SS.Row = 1
    SS.Col = 1
    For I = 1 To nStartCol - 1
        SS.Text = Trim(Str(nBeforeMonthLast - nStartCol + 1 + I))
        SS.BackColor = RGB(255, 255, 192)
        SS.ForeColor = RGB(192, 192, 192)
        If SS.Col = 7 Then
            If SS.Row < 6 Then
                SS.Row = SS.Row + 1
                SS.Col = 1
            Else
                Exit For
            End If
        Else
            SS.Col = SS.Col + 1
        End If
    Next I
    
    ' 당월 Display
    For I = 1 To nCurrentMonthLast
        SS.Text = Trim(Str(I))
        ' 당일
        'If CVDate(Format$(ArgDate, "YYYY-MM-") & Format$(I, "00")) = Date Then SS.BackColor = RGB(0, 255, 0)
        ' 공휴일
        'If Is_HolyDay(CVDate(Format$(ArgDate, "YYYY-MM-") & Format$(I, "00"))) Then SS.ForeColor = RGB(255, 0, 0)
        If SS.Col = 7 Then
            If SS.Row < 6 Then
                SS.Row = SS.Row + 1
                SS.Col = 1
            Else
                Exit For
            End If
        Else
            SS.Col = SS.Col + 1
        End If
    Next I
    
    ' 다음월 Display
    For I = 1 To 14
        SS.Text = Trim(Str(I))
        SS.BackColor = RGB(255, 255, 192)
        SS.ForeColor = RGB(192, 192, 192)
        If SS.Col = 7 Then
            If SS.Row < 6 Then
                SS.Row = SS.Row + 1
                SS.Col = 1
            Else
                Exit For
            End If
        Else
            SS.Col = SS.Col + 1
        End If
    Next I
    
    If Format(ArgDate, "YYYYMM") > Trim(Str(START_YEAR)) & "01" Then
        SS.Col = 1:          SS.Row = 1:          SS.Text = " ◀ "
    End If
    If Format(ArgDate, "YYYYMM") < Trim(Str(LAST_YEAR)) & "12" Then
        SS.Col = SS.MaxCols: SS.Row = SS.MaxRows: SS.Text = " ▶ "
    End If

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
    SS.Col = nCol
    SS.Row = nRow
    SS.Action = 0       'SS_ACTION_ACTIVE_CELL
    Call Spread_Select(nRow, nCol)

End Sub

Private Sub Spread_Clear()

    If SS.ReDraw = True Then
        SS.ReDraw = False
        GoSub Clearing
        SS.ReDraw = True
    Else
        GoSub Clearing
    End If
    
Exit Sub

'/-----------------------------------------------------/

Clearing:

    SS.Col = 1:     SS.Col2 = SS.MaxCols
    SS.Row = 1:     SS.Row2 = SS.MaxRows
    SS.BlockMode = True
    SS.Text = ""
    SS.BackColor = RGB(255, 255, 0)
    SS.ForeColor = RGB(0, 0, 0)
    SS.BlockMode = False
    
    SS.Col = 1
    SS.Row = -1
    SS.ForeColor = RGB(255, 0, 0)
    
    SS.Col = 7
    SS.Row = -1
    SS.ForeColor = RGB(0, 0, 255)
    
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
    
    SS.Row = 0
    SS.Col = 1:     SS.Text = "Sun"
    SS.Col = 2:     SS.Text = "Mon"
    SS.Col = 3:     SS.Text = "Tue"
    SS.Col = 4:     SS.Text = "Wed"
    SS.Col = 5:     SS.Text = "Thu"
    SS.Col = 6:     SS.Text = "Fri"
    SS.Col = 7:     SS.Text = "Sat"

End Sub


Private Sub Spread_Select(ByVal Row As Integer, ByVal Col As Integer)
    
    Dim strFirstDate    As String
    
    SS.Col = Col:     FnActiveCol = Col
    SS.Row = Row:     FnActiveRow = Row
    FnOldBackColor = SS.BackColor
    SS.BackColor = RGB(255, 255, 255)
    
    strFirstDate = Format(FdSelectedDate, "YYYY-MM") & "-01"
    FdSelectedDate = DateAdd("d", Val(SS.Text) - 1, strFirstDate)
    Me.Caption = Format(FdSelectedDate, "Long Date")

End Sub


Private Sub Spread_UnSelect()

    SS.Col = FnActiveCol
    SS.Row = FnActiveRow
    SS.BackColor = FnOldBackColor

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
        If Year(FdSelectedDate) < START_YEAR Or Year(FdSelectedDate) > LAST_YEAR Then
            FdSelectedDate = Date
        End If
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

Private Sub SS_DblClick(Col As Long, Row As Long)

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
        SS.Action = 0       'SS_ACTION_ACTIVE_CELL
    ElseIf KeyCode = 39 And FnActiveCol = 7 Then
        KeyCode = 0
        If FnActiveRow = 7 Then Exit Sub
        nCol = 1
        nRow = FnActiveRow + 1
        Call SS_LeaveCell(FnActiveCol + 0, FnActiveRow + 0, nCol + 0, nRow + 0, False)
        SS.Action = 0       'SS_ACTION_ACTIVE_CELL
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

Private Sub SS_LeaveCell(Col As Long, Row As Long, NewCol As Long, NewRow As Long, Cancel As Integer)

    If (Col < 1) Or (Row < 1) Then Exit Sub
    If (NewCol < 1) Or (NewRow < 1) Then Exit Sub
    
    SS.Col = NewCol
    SS.Row = NewRow
    If SS.BackColor = RGB(255, 255, 192) Then
        If NewRow < 2 Then
            If FdSelectedDate < CVDate(Format(START_YEAR, "###0") & "-02-01") Then
                Cancel = True
                Exit Sub
            End If
            FdSelectedDate = DateAdd("m", -1, FdSelectedDate)
            If IsNumeric(SS.Text) Then  ' 화살표(전월) 제외
                FdSelectedDate = Format(FdSelectedDate, "YYYY-MM") & "-" & Format(Val(SS.Text), "00")
            End If
            Call Display_Month(FdSelectedDate)
            Call Select_Current_Date(FdSelectedDate)
        Else
            If FdSelectedDate >= CVDate(Format(LAST_YEAR, "###0") & "-12-01") Then
                Cancel = True
                Exit Sub
            End If
            FdSelectedDate = DateAdd("m", 1, FdSelectedDate)
            If IsNumeric(SS.Text) Then  ' 화살표(다음월) 제외
                FdSelectedDate = Format(FdSelectedDate, "YYYY-MM") & "-" & Format(Val(SS.Text), "00")
            End If
            Call Display_Month(FdSelectedDate)
            Call Select_Current_Date(FdSelectedDate)
        End If
    Else
        Call Spread_UnSelect
        Call Spread_Select(NewRow, NewCol)
    End If
    
End Sub

