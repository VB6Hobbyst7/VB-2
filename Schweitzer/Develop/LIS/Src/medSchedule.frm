VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Begin VB.Form medSchedule 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "  개인 스케쥴 관리"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   Icon            =   "medSchedule.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin DRcontrol1.DrFrame fraInput 
      Height          =   3075
      Left            =   2430
      TabIndex        =   10
      Top             =   2460
      Visible         =   0   'False
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   5424
      Title           =   "Schedule 작성"
      TitlePos        =   1
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   285
         Left            =   690
         TabIndex        =   18
         Top             =   480
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         Format          =   66584578
         CurrentDate     =   37964
      End
      Begin VB.TextBox txtDescription 
         Height          =   1650
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "medSchedule.frx":09EA
         Top             =   855
         Width           =   5505
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00DBE6E6&
         Caption         =   "확인"
         Height          =   375
         Left            =   4650
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   2610
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "취소"
         Height          =   375
         Left            =   3540
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   2610
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   ":"
         Height          =   165
         Left            =   1410
         TabIndex        =   15
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  '투명
         Caption         =   "시간"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   510
         Width           =   360
      End
   End
   Begin VB.TextBox txtSchdule 
      Appearance      =   0  '평면
      BackColor       =   &H00D1ECFC&
      Height          =   2745
      Left            =   4665
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "medSchedule.frx":0A0C
      Top             =   4920
      Width           =   5535
   End
   Begin VB.HScrollBar spnDay 
      Height          =   315
      Left            =   6975
      Max             =   2
      TabIndex        =   9
      Top             =   840
      Width           =   480
   End
   Begin VB.HScrollBar spnMonth 
      Height          =   315
      Left            =   1815
      Max             =   2
      TabIndex        =   8
      Top             =   810
      Width           =   480
   End
   Begin FPSpread.vaSpread tblMonth 
      Height          =   3420
      Left            =   225
      TabIndex        =   2
      Top             =   1155
      Width           =   4215
      _Version        =   196608
      _ExtentX        =   7435
      _ExtentY        =   6032
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      AutoSize        =   -1  'True
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
      Protect         =   0   'False
      ScrollBars      =   0
      ShadowColor     =   12648447
      ShadowDark      =   12648447
      ShadowText      =   0
      SpreadDesigner  =   "medSchedule.frx":0A12
      VisibleCols     =   7
      VisibleRows     =   10
   End
   Begin FPSpread.vaSpread tblDay 
      Height          =   3420
      Left            =   4620
      TabIndex        =   4
      Top             =   1155
      Width           =   5565
      _Version        =   196608
      _ExtentX        =   9816
      _ExtentY        =   6033
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   10
      OperationMode   =   1
      Protect         =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "medSchedule.frx":1E06
      UserResize      =   0
      VisibleCols     =   2
      VisibleRows     =   10
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   8880
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   7725
      Width           =   1320
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  '평면
      BackColor       =   &H00D1ECFC&
      Height          =   2760
      Left            =   255
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "medSchedule.frx":221F
      Top             =   4920
      Width           =   4200
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '투명
      Caption         =   "Today's Schedule"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4650
      TabIndex        =   17
      Top             =   4635
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '투명
      Caption         =   "Today's Memo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   270
      TabIndex        =   7
      Top             =   4635
      Width           =   1905
   End
   Begin VB.Label lblDay 
      BackColor       =   &H00DBE6E6&
      Caption         =   "1998년 12월 28일 월요일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4620
      TabIndex        =   5
      Top             =   855
      Width           =   2310
   End
   Begin VB.Label lblMonth 
      BackColor       =   &H00DBE6E6&
      Caption         =   "1998년 12월"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   405
      Picture         =   "medSchedule.frx":2225
      Top             =   180
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "스케쥴"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   405
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1050
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
      Top             =   105
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  '단색
      Height          =   630
      Index           =   1
      Left            =   225
      Shape           =   4  '둥근 사각형
      Top             =   150
      Width           =   2295
   End
End
Attribute VB_Name = "medSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ScheduleCnt(1 To 31)    As Integer
'Private WithEvents mnuPopup     As menu
'Private WithEvents mnuAdd       As menu
'Private WithEvents mnuDelete    As menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_ADD& = 1
Private Const MENU_DEL& = 2

Private Today                   As Date
Private tmpDate1
Private RealDate

Private Sub cmdCancel_Click()
    fraInput.Visible = False
End Sub

Private Sub cmdExit_Click()
'    Set mnuPopup = Nothing
'    Set mnuAdd = Nothing
'    Set mnuDelete = Nothing
    Call SetNote
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim tmpStr  As String
    Dim i       As Integer

    If tblDay.MaxRows < tblDay.DataRowCnt Then
        tblDay.MaxRows = tblDay.MaxRows + 1
    End If
    tblDay.Row = tblDay.DataRowCnt + 1

    tblDay.Col = 1
    tblDay.Text = Format(dtpTime.Value, "HH:MM")
    tblDay.Col = 2
    tblDay.Value = txtDescription.Text

    ScheduleCnt(Day(Today)) = tblDay.DataRowCnt
    
    Call SetSchedule(Today)
    Call SetNote
    
    Call GetCalendar(Today)
    Call DisplayDate(Today)
    Call FindDay(Today)
    Call Day_Schedule(Today)
    
    fraInput.Visible = False
End Sub

Private Sub Form_Load()
    Dim strFileNm   As String
    Dim strTitle    As String
    Dim WeekDayKor  As String


    Me.Top = medMain.Top + (medMain.Height - Me.Height) / 2
    Me.Left = medMain.Left + (medMain.Width - Me.Width) / 2
    
    dtpTime.Value = GetSystemdate
    
    Call medAlwaysOn(Me, 1)
    Erase ScheduleCnt

    Today = Format(GetSystemdate, "YY-MM-DD")
    RealDate = Today
    Call GetCalendar(Today)
    Call DisplayDate(Today)
    Call FindDay(Today)
    Call Day_Schedule(Today)
    
    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", _
                          "화요일", "수요일", "목요일", "금요일", "토요일")
    
    spnMonth.Value = 1
    spnDay.Value = 1
    txtSchdule.Text = ""

End Sub



Private Sub spnDay_SpinDown()
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")
    
    
    ThisMonth = Month(Today)
    Today = DateAdd("y", -1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    lblDay = Format(Today, "YYYY년 MM월 DD일 ") & WeekDayKor
    Call Day_Schedule(Today)
    

    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If

End Sub

Private Sub spnDay_SpinUp()
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")
    
    ThisMonth = Month(Today)
    Today = DateAdd("y", 1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    lblDay = Format(Today, "YYYY년 MM월 DD일 ") & WeekDayKor
    Call Day_Schedule(Today)
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
    lblDay = Format(Today, "YYYY년 MM월 DD일 ") & WeekDayKor

    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If

End Sub

Private Sub spnMonth_SpinUp()
    Dim ThisMonth As Integer
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(Today), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")

    ThisMonth = Month(Today)
    Today = DateAdd("m", 1, Today)
    lblMonth = Format(Today, "YYYY년 MM월")
    lblDay = Format(Today, "YYYY년 MM월 DD일 ") & WeekDayKor

    If ThisMonth <> Month(Today) Then
        Call GetCalendar(Today)
    End If
    

End Sub


'Private Sub mnuAdd_Click()
'    txtDescription.Text = ""
'    fraInput.Visible = True
'    fraInput.ZOrder 0
'
'End Sub

'Private Sub mnuDelete_Click()
'    tblDay.Col = -1
'    tblDay.Action = ActionDeleteRow
'    ScheduleCnt(Day(Today)) = tblDay.DataRowCnt
'
'    Call SetSchedule(Today)
'    Call SetNote
'End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_ADD
            txtDescription.Text = ""
            fraInput.Visible = True
            fraInput.ZOrder 0
        Case MENU_DEL
            tblDay.Col = -1
            tblDay.Action = ActionDeleteRow
            ScheduleCnt(Day(Today)) = tblDay.DataRowCnt
            
            Call SetSchedule(Today)
            Call SetNote
    End Select
End Sub

Private Sub spnDay_Change()
   If spnDay.Value = 2 Then
      Call spnDay_SpinUp
   ElseIf spnDay.Value = 0 Then
      Call spnDay_SpinDown
   End If
   spnDay.Value = 1

End Sub

Private Sub spnMonth_Change()
   If spnMonth.Value = 2 Then
      Call spnMonth_SpinUp
   ElseIf spnMonth.Value = 0 Then
      Call spnMonth_SpinDown
   End If
   spnMonth.Value = 1
   txtNote.Text = ""
   txtSchdule.Text = ""
End Sub

Private Sub tblDay_Click(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub
    txtSchdule.Text = ""
    With tblDay
        .Row = Row
        .Col = 1
        If .Value = "" Then Exit Sub
        txtSchdule.Text = .Value
        .Col = 2: txtSchdule.Text = txtSchdule.Text & vbCRLF & .Value
    End With
End Sub

Private Sub tblDay_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Set objPop = Nothing
    Set objPop = New clsPopupMenu
    
    With objPop
        .AddMenu MENU_ADD, "SCHEDULE 추가"
        .AddMenu MENU_DEL, "SCHEDULE 삭제"
        
        .PopupMenus Me.hwnd
    End With
    
    Set objPop = Nothing
    
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuAdd = frmControls.mnuSub
'    Set mnuDelete = frmControls.mnuSub1
'
'    frmControls.mnuSub2.Visible = False
'    mnuAdd.Caption = "Schedule 추가"
'    mnuDelete.Caption = "Schedule 삭제"""
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuAdd = Nothing
'    Set mnuDelete = Nothing
End Sub

Private Sub tblMonth_Click(ByVal Col As Long, ByVal Row As Long)
    Dim tmpDay
    If Row = 0 Then Exit Sub
    If Row Mod 2 = 1 Then Exit Sub
    
    
    If Row Mod 2 = 0 Then
        tblMonth.Row = Row - 1
    Else
        tblMonth.Row = Row
    End If
    tblMonth.Col = Col
    
    tmpDay = Format(Today, "MM") & " " & tblMonth.Value & "," & Format(Today, "YYYY")
    tmpDate1 = CDate(tmpDay)
    
    tmpDay = tblMonth.Value
    lblDay.Caption = Format(tmpDate1, "Long Date")
    txtSchdule.Text = ""
    medClearTable tblDay
    Call Day_Schedule(tmpDate1)
End Sub
Private Sub Day_Schedule(ByVal datToday As Date)
    Dim tmpFileNm As String
    Dim InputData As String
    Dim i         As Integer
    Dim vbCRLF As String
    Dim vbTRS As String

    vbCRLF = Chr(13) & Chr(10)
    vbTRS = Chr(2)
    Call medClearTable(tblDay)
    txtDescription.Text = ""
    txtSchdule.Text = ""
    txtNote.Text = ""
    
    tmpFileNm = App.Path & "\" & ObjSysInfo.EmpId & Format(datToday, "YYMMDD") & ".txt"
    
    If Dir(tmpFileNm) <> "" Then
        i = 1
        Open tmpFileNm For Input As #1
        Do While Not EOF(1)   ' 파일의 끝을 확인합니다.
            Line Input #1, InputData   ' 데이터 행을 읽어 들입니다.
            tblDay.Row = i
            tblDay.Col = 1
            tblDay.Value = medGetKey(InputData, Chr(1))
            tblDay.Col = 2
            tblDay.Value = medTR(medGetKey(InputData, Chr(1)), vbTRS, vbCRLF)
            tblDay.TypeHAlign = TypeHAlignLeft
            i = i + 1
        Loop
        Close #1
        
    End If
    
    tmpFileNm = App.Path & "\" & ObjSysInfo.EmpId & Format(datToday, "YYYYMMDD") & ".txt"
    
    txtNote.Text = ""
    If Dir(tmpFileNm) <> "" Then
        Open tmpFileNm For Input As #1
        Do While Not EOF(1)   ' 파일의 끝을 확인합니다.
            Line Input #1, InputData   ' 데이터 행을 읽어 들입니다.
            txtNote.Text = txtNote.Text & InputData & Chr(13) & Chr(10)
        Loop
        Close #1
    End If
End Sub

Private Sub GetCalendar(ByVal datToday As Date)
    Dim tmpDate         As Date
    Dim ThisMonth       As Integer
    Dim ThisDay         As Integer
    Dim FirstWeekDay    As Integer
    Dim i               As Integer
    Dim strTmpFile      As String
    

    Call medClearTable(tblDay)
    txtSchdule.Text = "": txtDescription.Text = ""

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
        
        tblMonth.BackColor = &HFFFFFF
        tblMonth.Value = ThisDay
        tblMonth.Row = tblMonth.Row + 1
        If ScheduleCnt(ThisDay) <> 0 Then
            tblMonth.Text = ScheduleCnt(ThisDay)
        End If
        strTmpFile = App.Path & "\" & ObjSysInfo.EmpId & Mid(Format(Today, "YYYYMM"), 3) & Format(ThisDay, "00") & ".txt"
        If Dir(strTmpFile) <> "" Then
            tblMonth.Value = "Y": tblMonth.ForeColor = DCM_LightRed: tblMonth.FontBold = True
        Else
            tblMonth.Value = "": tblMonth.ForeColor = vbBlack
        End If
        tblMonth.BackColor = &HFFFFFF
        

        tmpDate = DateAdd("d", 1, tmpDate)
    Loop
    Call FindDay(Today)
End Sub

Private Sub DisplayDate(ByVal datToday As Date)
    Dim WeekDayKor As String

    WeekDayKor = Choose(Weekday(datToday), "일요일", "월요일", "화요일", _
                        "수요일", "목요일", "금요일", "토요일")
    lblMonth = Format(datToday, "YYYY년 MM월")
    lblDay = Format(datToday, "YYYY년 MM월 DD일 ") & WeekDayKor

End Sub


Private Sub SetSchedule(ByVal datToday As Date)
    Dim tmpFileNm   As String
    Dim strInput    As String
    Dim i           As Integer
    Dim vbCRLF      As String
    Dim vbTRS       As String

    vbCRLF = Chr(13) & Chr(10)
    vbTRS = Chr(2)

    tmpFileNm = App.Path & "\" & ObjSysInfo.EmpId & Format(tmpDate1, "YYMMDD") & ".txt"
    Open tmpFileNm For Output As #1
    For i = 1 To tblDay.DataRowCnt
        tblDay.Row = i
        tblDay.Col = 1
        strInput = tblDay.Value
        tblDay.Col = 2
        strInput = strInput & Chr(1) & medTR(tblDay.Value, vbCRLF, vbTRS)
        Print #1, strInput
    Next i
    Close #1

End Sub

Private Sub SetNote()
    Dim tmpFileNm As String
    
    If txtNote.Text = "" Then Exit Sub
    
    tmpFileNm = App.Path & "\" & ObjSysInfo.EmpId & Format(tmpDate1, "YYYYMMDD") & ".txt"

    Open tmpFileNm For Output As #1
    Print #1, txtNote.Text
    Close #1

End Sub
Private Sub FindDay(ByVal NewDate As Date)
    Dim ii As Integer
    Dim jj As Long
    
    With tblMonth
        For ii = 1 To 9 Step 2
            .Row = ii
            For jj = 1 To 7
                .Col = jj
                If Format(Today, "yyMM") & Format(.Value, "00") = Format(RealDate, "yymmdd") Then
                    .BackColor = &HFFFFC0
                End If
            Next jj
        Next ii
    
    End With
End Sub



'*-----------------------------------------------------------------
'*  1. 기능 : 문자열 내의 특정 string을 다른 string으로 대치한다.
'*  2. 관련변수 :
'*  3. Parameter : strOrigin - 대상 문자열
'*                       transFROM - 바꿀 문자열
'*                       transTo - 새 문자열
'*-----------------------------------------------------------------
Function medTR(ByVal strOrigin As String, ByVal transFROM As String, transTo As String)
Dim i As Integer
Dim intLen As Integer

    i = 1
    intLen = Len(transFROM)
    Do While i <= Len(strOrigin)
        If Mid(strOrigin, i, intLen) = transFROM Then
            strOrigin = Mid(strOrigin, 1, i - 1) & transTo & Mid(strOrigin, i + intLen)
        End If
        i = i + 1
    Loop
    medTR = strOrigin

End Function

'*-----------------------------------------------------------------
'*  1. 기능 : BLBX의 Delimiter로 구분된 첫번째 String을 읽어오고
'*               나머지 문자열은 그대로 남긴다.
'*  2. 관련변수 :
'*  3. Parameter : strText - Delimiter로 묶여있는 대상 문자열
'*  4. ReturnValue : 선택된 문자열
'*                   strText 자신은 Shift가 이루어진다.
'*-----------------------------------------------------------------
Public Function medGetKey(ByRef strText As String, ByVal strDeli As String) As String
Dim CNTA, CNTB As Integer

    medGetKey = "": CNTA = 0: CNTB = 0

    CNTA = InStr(1, strText, strDeli)
    If CNTA = 0 Then
        medGetKey = strText
        strText = ""
        Exit Function
    End If

    medGetKey = Mid$(strText, 1, CNTA - 1)
    strText = Mid$(strText, CNTA + 1)

End Function
