VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmCalendar 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Calendar"
   ClientHeight    =   3120
   ClientLeft      =   3096
   ClientTop       =   1848
   ClientWidth     =   3276
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmcal32.frx":0000
   ScaleHeight     =   260
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancel 
      Caption         =   "취 소"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2190
      TabIndex        =   5
      Top             =   2760
      Width           =   1035
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "확 인"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1140
      TabIndex        =   4
      Top             =   2760
      Width           =   1035
   End
   Begin Threed.SSPanel PanelYear 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3165
      _Version        =   65536
      _ExtentX        =   5583
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "1998 년"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.6
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.Label LabelYear 
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2865
         TabIndex        =   7
         Top             =   15
         Width           =   285
      End
      Begin VB.Label LabelYear 
         Caption         =   "◀"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   15
         TabIndex        =   6
         Top             =   15
         Width           =   345
      End
   End
   Begin Threed.SSPanel Panel 
      Align           =   1  '위 맞춤
      Height          =   2772
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3276
      _Version        =   65536
      _ExtentX        =   5778
      _ExtentY        =   4890
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread SScal 
         Height          =   1980
         Left            =   60
         TabIndex        =   3
         Top             =   750
         Width           =   3165
         _Version        =   196608
         _ExtentX        =   5583
         _ExtentY        =   3493
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   7
         MaxRows         =   6
         OperationMode   =   1
         ScrollBars      =   0
         SelectBlockOptions=   0
         ShadowText      =   0
         SpreadDesigner  =   "frmcal32.frx":0442
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin Threed.SSPanel PanelMonth 
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   390
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "09 월 (September)"
         ForeColor       =   8388608
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.6
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.Label LabelMonth 
            Caption         =   "▶"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2865
            TabIndex        =   9
            Top             =   15
            Width           =   285
         End
         Begin VB.Label LabelMonth 
            Caption         =   "◀"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   15
            TabIndex        =   8
            Top             =   15
            Width           =   345
         End
      End
   End
   Begin VB.Menu mnuYear 
      Caption         =   "&Year"
      Visible         =   0   'False
      Begin VB.Menu mnuYears 
         Caption         =   "1900 년대"
         Index           =   0
         Begin VB.Menu mnuYears00 
            Caption         =   "1900"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1910 년대"
         Index           =   1
         Begin VB.Menu mnuYears01 
            Caption         =   "1910"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1920 년대"
         Index           =   2
         Begin VB.Menu mnuYears02 
            Caption         =   "1920"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1930 년대"
         Index           =   3
         Begin VB.Menu mnuYears03 
            Caption         =   "1930"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1940 년대"
         Index           =   4
         Begin VB.Menu mnuYears04 
            Caption         =   "1940"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1950 년대"
         Index           =   5
         Begin VB.Menu mnuYears05 
            Caption         =   "1950"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1960 년대"
         Index           =   6
         Begin VB.Menu mnuYears06 
            Caption         =   "1960"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1970 년대"
         Index           =   7
         Begin VB.Menu mnuYears07 
            Caption         =   "1970"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1980 년대"
         Index           =   8
         Begin VB.Menu mnuYears08 
            Caption         =   "1980"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "1990 년대"
         Index           =   9
         Begin VB.Menu mnuYears09 
            Caption         =   "1990"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "2000 년대"
         Index           =   10
         Begin VB.Menu mnuYears10 
            Caption         =   "2000"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "2010 년대"
         Index           =   11
         Begin VB.Menu mnuYears11 
            Caption         =   "2010"
            Index           =   0
         End
      End
      Begin VB.Menu mnuYears 
         Caption         =   "2020 년대"
         Index           =   12
         Begin VB.Menu mnuYears12 
            Caption         =   "2020"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuMonth 
      Caption         =   "&Month"
      Visible         =   0   'False
      Begin VB.Menu mnuMonths 
         Caption         =   "01 월"
         Index           =   1
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "02 월"
         Index           =   2
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "03 월"
         Index           =   3
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "04 월"
         Index           =   4
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "05 월"
         Index           =   5
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "06 월"
         Index           =   6
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "07 월"
         Index           =   7
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "08 월"
         Index           =   8
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "09 월"
         Index           =   9
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "10 월"
         Index           =   10
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "11 월"
         Index           =   11
      End
      Begin VB.Menu mnuMonths 
         Caption         =   "12 월"
         Index           =   12
      End
   End
End
Attribute VB_Name = "FrmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
'Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long
'
'Type PointAPI
'     X As Long
'     Y As Long
'End Type
'Global ReturnPos    As PointAPI
'위에 정의 한것들을 모두 BAS File(Module)에 정의를 하세요
'이 Form을 사용하기위해서는 꼭 필요합니다.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'달력을 띠우기위해서는 필요한 Control 에서
'       Call FrmCalendar.Calendar_Show(Object_Name)
'위와 같이 선언하면 됩니다.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'

Public GstrCalendarDate     As String
Public GnCurrScreenX        As Long
Public GnCurrScreenY        As Long
Public GnControlWidth       As Long
Public GnControlHeight      As Long

Dim strChoiceDay            As String

Dim nYear                   As Integer
Dim nMon                    As Integer
Dim nDay                    As Integer
Dim nCurrCol                As Integer
Dim nCurrRow                As Integer

Dim naDays(12)              As Integer
Dim saMonths(12)            As String
Dim saText()                As String
Dim saMenuYear(10)          As String


Public Sub Calendar_Show(ArgControl As Object)
    Dim nCursor             As Long
    
    If TypeOf ArgControl Is vaSpread Then
        ArgControl.Col = ArgControl.ActiveCol       'Spread Sheet일경우
        ArgControl.Row = ArgControl.ActiveRow       '해당 셀에서 날자 형식 Move
        GstrCalendarDate = ArgControl.Caption
    Else
        GstrCalendarDate = ArgControl.Caption       '해당 TextBox에서 날자형식 Move
    End If
    
    nCursor = GetCursorPos(ReturnPos)               '화면의 Cursor 위치 파악
    
    GnCurrScreenY = ReturnPos.y * 15                '800 * 600 Mode 기준
    GnCurrScreenX = ReturnPos.x * 15                'ScreenMode가 빠뀌면 변경요망
    GnControlWidth = ArgControl.Width
    GnControlHeight = ArgControl.Height
    
    FrmCalendar.Show 1
    
    If GstrCalendarDate > "" Then                   '날자를 선택했을경우
        If TypeOf ArgControl Is vaSpread Then
            ArgControl.Col = ArgControl.ActiveCol   'Spread Sheet일경우
            ArgControl.Row = ArgControl.ActiveRow   '해당 셀에 Move
            ArgControl.Caption = GstrCalendarDate
        Else
            ArgControl.Caption = GstrCalendarDate      '해당 TextBox에 Move
        End If
    End If
    
End Sub

Private Sub Date_Convert_DDMONYY()
    Dim strYy               As String
    Dim strMM               As String
    Dim strDD               As String
    
    strDD = Mid$(GstrCalendarDate, 1, 2)
    strMM = Mid$(GstrCalendarDate, 4, 3)
    strYy = Mid$(GstrCalendarDate, 8, 2)
    
    Select Case UCase(strMM)
        Case "JAN": strMM = "01"
        Case "FEB": strMM = "02"
        Case "MAR": strMM = "03"
        Case "APR": strMM = "04"
        Case "MAY": strMM = "05"
        Case "JUN": strMM = "06"
        Case "JUL": strMM = "07"
        Case "AUG": strMM = "08"
        Case "SEP": strMM = "09"
        Case "OCT": strMM = "10"
        Case "NOV": strMM = "11"
        Case "DEC": strMM = "12"
    End Select
    
    Select Case strYy
        Case "00" To "30":  strYy = "20" & strYy
        Case Else:          strYy = "19" & strYy
    End Select
    
    GstrCalendarDate = strYy & "-" & strMM & "-" & strDD
    
End Sub


Private Sub Date_Convert_YYMMDD()
    Dim strYy               As String
    Dim strMM               As String
    Dim strDD               As String
    
    If Mid$(GstrCalendarDate, 3, 1) = "-" Then
        Select Case Mid$(GstrCalendarDate, 1, 2)
            Case "00" To "30":  GstrCalendarDate = "20" & GstrCalendarDate
            Case Else:          GstrCalendarDate = "19" & GstrCalendarDate
        End Select
    Else
        If Len(GstrCalendarDate) = 6 Then
            strYy = Mid$(GstrCalendarDate, 1, 2)
            strMM = Mid$(GstrCalendarDate, 3, 2)
            strDD = Mid$(GstrCalendarDate, 5, 2)
            Select Case strYy
                Case "00" To "30":  strYy = "20" & strYy
                Case Else:          strYy = "19" & strYy
            End Select
        Else
            strYy = Mid$(GstrCalendarDate, 1, 4)
            strMM = Mid$(GstrCalendarDate, 5, 2)
            strDD = Mid$(GstrCalendarDate, 7, 2)
        End If
        
        GstrCalendarDate = strYy & "-" & strMM & "-" & strDD
    End If
    
End Sub
Private Sub Display_Calendar(ArgYear%, ArgMonth%, ArgDay%)
    Dim nCol%
    Dim d$
    Dim i%, j%, k%
    
    
    naDays(2) = 28
    If ((ArgYear% Mod 4) = 0) Then
        naDays(2) = 29
        If ((ArgYear% Mod 100) = 0) Then
            naDays(2) = 28
            If ((ArgYear% Mod 400) = 0) Then
                naDays(2) = 29
            End If
        End If
    End If
    
    PanelYear.Caption = ArgYear% & " 년"
    PanelMonth.Caption = Format$(ArgMonth%, "00") & " 월  (" & saMonths(ArgMonth%) & ")"
    strChoiceDay = Format(ArgDay%, "00")
    If ArgDay% > naDays(ArgMonth%) Then strChoiceDay = "00"
    
    SScal.ReDraw = False
    d = Format$(ArgYear%, "0000") & "-" & Format$(ArgMonth%, "00") & "-" & "01"
    nCol = Weekday(d)
    
    ReDim saText(6, 7)
    
    k = 0
    For i = 1 To 6
        For j = nCol To 7
            k = k + 1
            If k <= naDays(ArgMonth%) Then saText(i, j) = CStr(k)
        Next j
        nCol = 1
    Next i
    
    nCurrCol = 0
    nCurrRow = 0
    
    For i = 1 To 6
        For j = 1 To 7
            SScal.Row = i
            SScal.Col = j
            SScal.TypeButtonText = saText(i, j)
            
            If Val(saText(i, j)) > 0 And Val(saText(i, j)) = ArgDay% Then
                SScal.TypeButtonColor = RGB(128, 255, 255)
                nCurrCol = j
                nCurrRow = i
            End If
        Next j
    Next i
    
    SScal.ReDraw = True
    
    If strChoiceDay < "01" Or strChoiceDay > "31" Then
        CmdOK.Enabled = False
    End If
    
End Sub

Private Sub CmdCancel_Click()

    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    GstrCalendarDate = ""
    Me.Hide

End Sub



Private Sub CmdOK_Click()
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    GstrCalendarDate = Mid$(PanelYear.Caption, 1, 4) & "-" & _
                       Mid$(PanelMonth.Caption, 1, 2) & "-" & _
                       strChoiceDay
    
    Me.Hide
    
End Sub



Private Sub Form_Activate()
    Dim nTop                As Long
    Dim nLeft               As Long
    
    If GnCurrScreenX = 0 Or GnCurrScreenY = 0 Then
        Me.Top = (Screen.Height - Me.Height) / 2 + 200
        Me.Left = (Screen.Width - Me.Width) / 2
    Else
        nTop = GnCurrScreenY
        nLeft = GnCurrScreenX
        
        If (Screen.Height - GnCurrScreenY) < Me.Height Then
           'nTop = GnCurrScreenY - Me.Height
            nTop = 8800 - Me.Height
            If nTop < 0 Then nTop = 0
        End If
        
        If (Screen.Width - GnCurrScreenX) < Me.Width Then
            nLeft = GnCurrScreenX - Me.Width
            If nLeft < 0 Then nLeft = 0
        End If
        
        Me.Top = nTop
        Me.Left = nLeft
    End If
    
    
    If Len(GstrCalendarDate) = 6 Then Call Date_Convert_YYMMDD
    If Len(GstrCalendarDate) = 8 Then Call Date_Convert_YYMMDD
    If Len(GstrCalendarDate) = 9 Then Call Date_Convert_DDMONYY
    
    If Not IsDate(GstrCalendarDate) Then
        GstrCalendarDate = Format(Now, "yyyy-mm-dd")
    End If
    
    nYear = Year(GstrCalendarDate)
    nMon = Month(GstrCalendarDate)
    nDay = Day(GstrCalendarDate)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub Form_Load()
    Dim i%, j%
    Dim nTop                As Long
    Dim nLeft               As Long
    
    For i% = 1 To 12
        Select Case i%
            Case 4, 6, 9, 11:   naDays(i%) = 30
            Case 2:             naDays(i%) = 28
            Case Else:          naDays(i%) = 31
        End Select
    Next i%
    
    For j% = 1 To 9
        Load mnuYears00(j%)
        Load mnuYears01(j%)
        Load mnuYears02(j%)
        Load mnuYears03(j%)
        Load mnuYears04(j%)
        Load mnuYears05(j%)
        Load mnuYears06(j%)
        Load mnuYears07(j%)
        Load mnuYears08(j%)
        Load mnuYears09(j%)
        Load mnuYears10(j%)
        Load mnuYears11(j%)
        Load mnuYears12(j%)
    Next j%
    
    saMonths(1) = "January"
    saMonths(2) = "February"
    saMonths(3) = "March"
    saMonths(4) = "April"
    saMonths(5) = "May"
    saMonths(6) = "June"
    saMonths(7) = "July"
    saMonths(8) = "August"
    saMonths(9) = "September"
    saMonths(10) = "October"
    saMonths(11) = "November"
    saMonths(12) = "December"
    
    SScal.CursorStyle = SS_CURSOR_STYLE_ARROW
    
    For i = 1 To 6
        SScal.Row = i
        SScal.Col = 1: SScal.TypeButtonTextColor = RGB(255, 0, 0)
        SScal.Col = 7: SScal.TypeButtonTextColor = RGB(0, 0, 255)
    Next i
    
    
    If GnCurrScreenX = 0 Or GnCurrScreenY = 0 Then
        Me.Top = (Screen.Height - Me.Height) / 2 + 200
        Me.Left = (Screen.Width - Me.Width) / 2
    Else
        nTop = GnCurrScreenY
        nLeft = GnCurrScreenX
        
        If (Screen.Height - GnCurrScreenY) < Me.Height Then
           'nTop = GnCurrScreenY - Me.Height
            nTop = 8800 - Me.Height
            If nTop < 0 Then nTop = 0
        End If
        
        If (Screen.Width - GnCurrScreenX) < Me.Width Then
            nLeft = GnCurrScreenX - Me.Width
            If nLeft < 0 Then nLeft = 0
        End If
        
        Me.Top = nTop
        Me.Left = nLeft
    End If
    
End Sub


Private Sub LabelMonth_Click(Index As Integer)
    
    nYear = Mid$(PanelYear.Caption, 1, 4)
    nMon = Mid$(PanelMonth.Caption, 1, 2)
    nDay = Val(strChoiceDay)
    
    Select Case Index
        Case 0: nMon = nMon - 1
        Case 1: nMon = nMon + 1
    End Select
    
    If nMon < 1 Then nMon = 12: nYear = nYear - 1
    If nMon > 12 Then nMon = 1: nYear = nYear + 1
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub LabelYear_Click(Index As Integer)
    
    nYear = Mid$(PanelYear.Caption, 1, 4)
    nMon = Mid$(PanelMonth.Caption, 1, 2)
    nDay = Val(strChoiceDay)
    
    Select Case Index
        Case 0: nYear = nYear - 1
        Case 1: nYear = nYear + 1
    End Select
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuMonths_Click(Index As Integer)
    
    nYear = Val(Mid$(PanelYear.Caption, 1, 4))
    nMon = Val(Mid$(mnuMonths(Index).Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears00_Click(Index As Integer)
    
    nYear = Val(mnuYears00(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears01_Click(Index As Integer)
    
    nYear = Val(mnuYears01(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears02_Click(Index As Integer)
    
    nYear = Val(mnuYears02(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears03_Click(Index As Integer)
    
    nYear = Val(mnuYears03(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears04_Click(Index As Integer)
    
    nYear = Val(mnuYears04(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears05_Click(Index As Integer)
    
    nYear = Val(mnuYears05(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears06_Click(Index As Integer)
    
    nYear = Val(mnuYears06(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears07_Click(Index As Integer)
    
    nYear = Val(mnuYears07(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears08_Click(Index As Integer)
    
    nYear = Val(mnuYears08(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears09_Click(Index As Integer)
    
    nYear = Val(mnuYears09(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears10_Click(Index As Integer)
    
    nYear = Val(mnuYears10(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears11_Click(Index As Integer)
    
    nYear = Val(mnuYears11(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub mnuYears12_Click(Index As Integer)
    
    nYear = Val(mnuYears12(Index).Caption)
    nMon = Val(Mid$(PanelMonth.Caption, 1, 2))
    nDay = Val(strChoiceDay)
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    Call Display_Calendar(nYear, nMon, nDay)
    
End Sub

Private Sub PanelMonth_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        PopupMenu mnuMonth
    End If
    
End Sub

Private Sub PanelYear_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i%, j%
    Dim sNowYear            As String
    Dim nNowYear            As Integer
    
    If Button = vbRightButton Then
        sNowYear = Format(Now, "yyyy")
        nNowYear = Val(Mid$(sNowYear, 1, 3))
        
        For i = 0 To 12
            mnuYears(i).Caption = (nNowYear * 10) + ((i * 10) - 90) & " 년대"
        Next i
        
        For j = 0 To 9
            mnuYears00(j).Caption = Mid$(mnuYears(0).Caption, 1, 3) & j
            mnuYears01(j).Caption = Mid$(mnuYears(1).Caption, 1, 3) & j
            mnuYears02(j).Caption = Mid$(mnuYears(2).Caption, 1, 3) & j
            mnuYears03(j).Caption = Mid$(mnuYears(3).Caption, 1, 3) & j
            mnuYears04(j).Caption = Mid$(mnuYears(4).Caption, 1, 3) & j
            mnuYears05(j).Caption = Mid$(mnuYears(5).Caption, 1, 3) & j
            mnuYears06(j).Caption = Mid$(mnuYears(6).Caption, 1, 3) & j
            mnuYears07(j).Caption = Mid$(mnuYears(7).Caption, 1, 3) & j
            mnuYears08(j).Caption = Mid$(mnuYears(8).Caption, 1, 3) & j
            mnuYears09(j).Caption = Mid$(mnuYears(9).Caption, 1, 3) & j
            mnuYears10(j).Caption = Mid$(mnuYears(10).Caption, 1, 3) & j
            mnuYears11(j).Caption = Mid$(mnuYears(11).Caption, 1, 3) & j
            mnuYears12(j).Caption = Mid$(mnuYears(12).Caption, 1, 3) & j
        Next j
        
        PopupMenu mnuYear
    End If
    
End Sub

Private Sub SScal_Click(ByVal Col As Long, ByVal Row As Long)
    
    SScal.Col = Col
    SScal.Row = Row
    
    If Row < 1 Then Exit Sub
    If Trim$(SScal.TypeButtonText) = "" Then Exit Sub
    
    CmdOK.Enabled = True
    strChoiceDay = Format$(SScal.TypeButtonText, "00")
    
    SScal.ReDraw = False
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    SScal.Col = Col
    SScal.Row = Row
    SScal.TypeButtonColor = RGB(128, 255, 255)
    SScal.ReDraw = True
        
    nCurrCol = Col
    nCurrRow = Row
    
End Sub


Private Sub SScal_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    SScal.Col = Col
    SScal.Row = Row
    
    If Row < 1 Then Exit Sub
    If Trim$(SScal.TypeButtonText) = "" Then Exit Sub
    
    CmdOK.Enabled = True
    strChoiceDay = Format$(SScal.TypeButtonText, "00")
    
    SScal.ReDraw = False
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    SScal.Col = Col
    SScal.Row = Row
    SScal.TypeButtonColor = RGB(128, 255, 255)
    SScal.ReDraw = True
        
    nCurrCol = Col
    nCurrRow = Row
    
    SScal.Col = nCurrCol
    SScal.Row = nCurrRow
    SScal.TypeButtonColor = RGB(192, 192, 192)
    
    GstrCalendarDate = Mid$(PanelYear.Caption, 1, 4) & "-" & _
                       Mid$(PanelMonth.Caption, 1, 2) & "-" & _
                       strChoiceDay
    
    Me.Hide
    
End Sub


