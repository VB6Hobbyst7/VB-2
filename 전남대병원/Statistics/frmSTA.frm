VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSTA 
   Appearance      =   0  '평면
   BackColor       =   &H00F8E4D8&
   Caption         =   "AMR 현황"
   ClientHeight    =   13305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19035
   Icon            =   "frmSTA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   13305
   ScaleWidth      =   19035
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdEvent 
      Caption         =   "이벤트 기록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      TabIndex        =   18
      Top             =   90
      Width           =   2145
   End
   Begin VB.Frame fraEvent 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  '없음
      Height          =   10485
      Left            =   2820
      TabIndex        =   8
      Top             =   1410
      Visible         =   0   'False
      Width           =   12915
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9630
         TabIndex        =   16
         Top             =   570
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8370
         TabIndex        =   15
         Top             =   570
         Width           =   1245
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12060
         TabIndex        =   14
         Top             =   570
         Width           =   465
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11580
         TabIndex        =   13
         Top             =   570
         Width           =   465
      End
      Begin VB.TextBox txtFont 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10920
         TabIndex        =   12
         Text            =   "12"
         Top             =   570
         Width           =   615
      End
      Begin VB.TextBox txtEventLog 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   9315
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1110
         Width           =   12855
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   345
         Left            =   4830
         TabIndex        =   11
         Top             =   570
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364739
         CurrentDate     =   40238
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "장비 이벤트 기록"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2010
         TabIndex        =   9
         Top             =   540
         Width           =   3405
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   30
         Picture         =   "frmSTA.frx":144A
         Top             =   30
         Width           =   12900
      End
   End
   Begin VB.Frame fraSTA 
      BackColor       =   &H00F8E4D8&
      Height          =   12375
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   18915
      Begin VB.CommandButton cmdExcel 
         Caption         =   "엑셀"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   17250
         TabIndex        =   19
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSTAPrint 
         Caption         =   "출력"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   15900
         TabIndex        =   17
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   14550
         TabIndex        =   7
         Top             =   240
         Width           =   1305
      End
      Begin FPSpreadADO.fpSpread spdMaster 
         Height          =   11445
         Left            =   120
         TabIndex        =   1
         Top             =   780
         Width           =   2580
         _Version        =   524288
         _ExtentX        =   4551
         _ExtentY        =   20188
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   14737632
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   2
         MaxRows         =   499
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmSTA.frx":2B8D
         UserResize      =   0
      End
      Begin FPSpreadADO.fpSpread spdStaList 
         Height          =   11445
         Left            =   2760
         TabIndex        =   2
         Top             =   780
         Width           =   16020
         _Version        =   524288
         _ExtentX        =   28258
         _ExtentY        =   20188
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   14737632
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   2
         MaxRows         =   36
         Protect         =   0   'False
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmSTA.frx":3094
         UserResize      =   0
         VisibleCols     =   2
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin MSComCtl2.DTPicker dtpYear 
         Height          =   465
         Left            =   12120
         TabIndex        =   5
         Top             =   240
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   820
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364739
         CurrentDate     =   40238
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회년도"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   10410
         TabIndex        =   6
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "검사항목"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   300
         Width           =   1785
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   180
         Picture         =   "frmSTA.frx":3E03
         Top             =   330
         Width           =   150
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   2820
         Picture         =   "frmSTA.frx":41ED
         Top             =   330
         Width           =   150
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "AMR 현황"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3150
         TabIndex        =   3
         Top             =   300
         Width           =   2805
      End
   End
End
Attribute VB_Name = "frmSTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- AMR 현황 조회

Option Explicit

Private Sub cmdDown_Click()

    txtFont.Text = txtFont.Text - 1
    txtEventLog.FontSize = txtFont.Text

End Sub

Private Sub cmdEvent_Click()
    
    If fraEvent.Visible = True Then
        fraEvent.Visible = False
    Else
        fraEvent.Visible = True
        fraEvent.ZOrder 0
        Call dtpToday_Click
        
    End If
    
End Sub

Private Sub cmdExcel_Click()
    
    Call spdStaList.ExportExcelBook(App.Path & "\AMR.xls", App.Path & "\export.txt")

    MsgBox "엑셀 출력완료" & vbNewLine & "위치 : " & App.Path & "\AMR.xls", vbOKOnly + vbInformation, Me.Caption
    
End Sub

Private Sub cmdPrint_Click()
    Dim prtText As String
    
    prtText = "장비 이벤트 기록   작성일 : " & dtpToday.Value & vbNewLine & vbNewLine
    prtText = prtText & txtEventLog.Text & vbNewLine & vbNewLine
    
'    Printer.CurrentX = X 'FONT_WIDTH * X
'    Printer.CurrentY = y 'FONT_HEIGHT * Y
    Printer.FontSize = 12
    Printer.FontName = 12
    Printer.FontBold = True
    
    Printer.Print prtText
    
    Printer.EndDoc

End Sub

Private Sub cmdSave_Click()
    
    If GetINUP(dtpToday.Value) = "IN" Then
              SQL = "INSERT INTO TB_EVENTLOG (EVTDATE,EVTLOG) "
        SQL = SQL & " Values ('" & dtpToday.Value & "','" & txtEventLog.Text & "') "
    ElseIf GetINUP(dtpToday.Value) = "UP" Then
              SQL = "UPDATE TB_EVENTLOG SET "
        SQL = SQL & " EVTLOG = '" & txtEventLog.Text & "'"
        SQL = SQL & " WHERE EVTDATE = '" & dtpToday.Value & "'"
    End If
    
    Set cmdSQL = New ADODB.Command
    Set cmdSQL.ActiveConnection = Cn_Ser
    
    cmdSQL.CommandText = SQL
    Set RS_Ser = cmdSQL.Execute
    
End Sub

Private Sub cmdSearch_Click()

    Call GetAMRList(Year(dtpYear))
    
End Sub

Private Sub cmdSTAPrint_Click()

'    If optPrint(0).Value = True Then
        spdStaList.PrintOrientation = PrintOrientationLandscape '가로출력
        spdStaList.Action = 13
'    Else
'        spdStaList.PrintOrientation = PrintOrientationPortrait '세로출력
'        spdStaList.Action = 13
'    End If
        
    MsgBox "결과 출력완료", vbOKOnly + vbInformation, Me.Caption
    
    
End Sub

Private Sub cmdUp_Click()
    
    txtFont.Text = txtFont.Text + 1
    txtEventLog.FontSize = txtFont.Text

End Sub

Private Sub dtpToday_Change()

    Call GetEventLog(dtpToday.Value)

End Sub

Private Sub dtpToday_Click()

    Call GetEventLog(dtpToday.Value)
    
End Sub


Private Function GetINUP(ByVal pEvtDate As String) As String
    Dim intCnt      As Integer
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim strAbbr     As String
    Dim strExamCode As String
    Dim varTmp      As Variant
    
    GetINUP = ""
    
    intRow = 1
    intCol = 3
    
          SQL = "Select EvtLog " & vbCr
    SQL = SQL & "  From TB_EventLog" & vbCr
    SQL = SQL & " Where EvtDate = '" & pEvtDate & "'"
 
    Cn_Ser.CursorLocation = adUseClient
    Set RS_Ser = Cn_Ser.Execute(SQL)
    
    If Not RS_Ser.EOF = True And Not RS_Ser.BOF = True Then
        GetINUP = "UP"
    Else
        GetINUP = "IN"
    End If
    
    RS_Ser.Close
    
End Function

Private Sub GetEventLog(ByVal pEvtDate As String)
              
    txtEventLog.Text = ""
          
          SQL = "Select EvtLog " & vbCr
    SQL = SQL & "  From TB_EventLog" & vbCr
    SQL = SQL & " Where EvtDate = '" & pEvtDate & "'"
 
    Cn_Ser.CursorLocation = adUseClient
    Set RS_Ser = Cn_Ser.Execute(SQL)
    
    Do Until RS_Ser.EOF
        txtEventLog.Text = Trim(RS_Ser.Fields("EvtLog"))
        txtEventLog.FontSize = txtFont.Text
        RS_Ser.MoveNext
    Loop
    
    RS_Ser.Close
    
End Sub


Private Sub Form_Load()
    
'    Me.Width = 22800
'    Me.Height = 13365
    
    Call FrmInitial
    
    If Connect_PRServer = False Then
        MsgBox "서버에 연결되지 않았습니다.", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    'Test ListUp
    Call GetEQPList
        
End Sub

Private Sub GetEQPList()
    Dim intCnt      As Integer
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim strAbbr     As String
    Dim strExamCode As String
    Dim varTmp      As Variant
    
    intRow = 1
    intCol = 3
          SQL = "Select DISTINCT ABBR1,EXAMCODE " & vbCr
    SQL = SQL & "  From TB_CODE" & vbCr
    Cn_Ser.CursorLocation = adUseClient
    Set RS_Ser = Cn_Ser.Execute(SQL)
    With spdMaster
        Do Until RS_Ser.EOF
            If intCnt = 0 Then
                .MaxRows = intRow
                Call .SetText(1, intRow, Trim(RS_Ser.Fields("ABBR1")))
                '-- AMR현황
                spdStaList.MaxCols = intCol
                Call spdStaList.SetText(intCol, 0, Trim(RS_Ser.Fields("ABBR1")))
            Else
                If strAbbr = Trim(RS_Ser.Fields("ABBR1")) Then
                    Call .GetText(2, intRow, varTmp)
                    Call .SetText(2, intRow, Trim(RS_Ser.Fields("EXAMCODE")) & "','" & varTmp)  '쿼리 IN 문
                Else
                    intRow = intRow + 1
                    .MaxRows = intRow
                    Call .SetText(1, intRow, Trim(RS_Ser.Fields("ABBR1")))
                    Call .SetText(2, intRow, Trim(RS_Ser.Fields("EXAMCODE")))
                    '-- AMR현황
                    intCol = intCol + 1
                    spdStaList.MaxCols = intCol
                    spdStaList.CellType = CellTypeStaticText
                    spdStaList.TypeEditCharSet = TypeEditCharSetASCII
                    spdStaList.TypeEditCharCase = TypeEditCharCaseSetNone
                    spdStaList.TypeHAlign = TypeHAlignCenter
                    spdStaList.TypeVAlign = TypeVAlignCenter
                    Call spdStaList.SetText(intCol, 0, Trim(RS_Ser.Fields("ABBR1")))
                End If
            End If
            strAbbr = Trim(RS_Ser.Fields("ABBR1"))
            intCnt = intCnt + 1
            RS_Ser.MoveNext
        Loop
        RS_Ser.Close
    End With
End Sub

Private Sub GetAMRList(ByVal pYear As String)
    Dim intCnt          As Integer
    Dim intRow          As Integer
    Dim intCol          As Integer
    Dim intStRow        As Integer
    Dim varTmp          As Variant
    Dim strExamCode()   As String
    Dim lngTot          As Long
    Dim lngAmr          As Long
        
    Screen.MousePointer = 11
    
    intRow = 1
    intStRow = 1
    
    '-- 전체 검사항목의 검사코드들을 배열변수에 저장
    With spdMaster
        ReDim Preserve strExamCode(.MaxRows)
        For intRow = 1 To .MaxRows
            Call .GetText(2, intRow, varTmp)
            strExamCode(intRow) = varTmp
        Next
    End With
    
    SQL = ""
    With spdStaList
        .ReDraw = False
        '-- 전체건수
        For intCol = 3 To .MaxCols
            SQL = ""
            For intCnt = 1 To 12
                SQL = SQL & "Select " & intCnt & " as MON, COUNT(*) as CNT " & vbCr
                SQL = SQL & "  From CNUH..TB_RESULT" & vbCr
                SQL = SQL & " Where SUBSTRING(EXAMDATE,1,4) = '" & pYear & "'" & vbCr
                SQL = SQL & "   And SUBSTRING(EXAMDATE,5,2) = '" & Format(intCnt, "00") & "'" & vbCr
                SQL = SQL & "   and EXAMCODE IN ('" & strExamCode(intCol - 2) & "')" & vbCr
                If intCnt = 12 Then
                    SQL = SQL & " ORDER BY MON " & vbCr
                Else
                    SQL = SQL & " UNION ALL " & vbCr
                End If
            Next
            'Call SetSQLData("장비총건수조회", SQL)
            Cn_Ser.CursorLocation = adUseClient
            Set RS_Ser = Cn_Ser.Execute(SQL)
            Do Until RS_Ser.EOF
                Select Case CStr(RS_Ser.Fields("MON"))
                    Case "1":  Call .SetText(intCol, intStRow, Trim(RS_Ser.Fields("CNT")))
                    Case "2":  Call .SetText(intCol, intStRow + 3, Trim(RS_Ser.Fields("CNT")))
                    Case "3":  Call .SetText(intCol, intStRow + 6, Trim(RS_Ser.Fields("CNT")))
                    Case "4":  Call .SetText(intCol, intStRow + 9, Trim(RS_Ser.Fields("CNT")))
                    Case "5":  Call .SetText(intCol, intStRow + 12, Trim(RS_Ser.Fields("CNT")))
                    Case "6":  Call .SetText(intCol, intStRow + 15, Trim(RS_Ser.Fields("CNT")))
                    Case "7":  Call .SetText(intCol, intStRow + 18, Trim(RS_Ser.Fields("CNT")))
                    Case "8":  Call .SetText(intCol, intStRow + 21, Trim(RS_Ser.Fields("CNT")))
                    Case "9":  Call .SetText(intCol, intStRow + 24, Trim(RS_Ser.Fields("CNT")))
                    Case "10": Call .SetText(intCol, intStRow + 27, Trim(RS_Ser.Fields("CNT")))
                    Case "11": Call .SetText(intCol, intStRow + 30, Trim(RS_Ser.Fields("CNT")))
                    Case "12": Call .SetText(intCol, intStRow + 33, Trim(RS_Ser.Fields("CNT")))
                End Select
                RS_Ser.MoveNext
            Loop
            RS_Ser.Close
        Next
        
        '-- 초과건수
        intStRow = 2
        For intCol = 3 To .MaxCols
            SQL = ""
            For intCnt = 1 To 12
                SQL = SQL & "Select " & intCnt & " as MON, COUNT(*) as CNT " & vbCr
                SQL = SQL & "  From CNUH..TB_RESULT" & vbCr
                SQL = SQL & " Where SUBSTRING(EXAMDATE,1,4) = '" & pYear & "'" & vbCr
                SQL = SQL & "   And SUBSTRING(EXAMDATE,5,2) = '" & Format(intCnt, "00") & "'" & vbCr
                SQL = SQL & "   and EXAMCODE IN ('" & strExamCode(intCol - 2) & "')" & vbCr
                SQL = SQL & "   and FLAG = 'A'" & vbCr
                If intCnt = 12 Then
                    SQL = SQL & " ORDER BY MON " & vbCr
                Else
                    SQL = SQL & " UNION ALL " & vbCr
                End If
            Next
            'Call SetSQLData("장비초과건수조회", SQL)
            Cn_Ser.CursorLocation = adUseClient
            Set RS_Ser = Cn_Ser.Execute(SQL)
            Do Until RS_Ser.EOF
                Select Case CStr(RS_Ser.Fields("MON"))
                    Case "1":   Call .SetText(intCol, intStRow, Trim(RS_Ser.Fields("CNT")))
                    Case "2":   Call .SetText(intCol, intStRow + 3, Trim(RS_Ser.Fields("CNT")))
                    Case "3":   Call .SetText(intCol, intStRow + 6, Trim(RS_Ser.Fields("CNT")))
                    Case "4":   Call .SetText(intCol, intStRow + 9, Trim(RS_Ser.Fields("CNT")))
                    Case "5":   Call .SetText(intCol, intStRow + 12, Trim(RS_Ser.Fields("CNT")))
                    Case "6":   Call .SetText(intCol, intStRow + 15, Trim(RS_Ser.Fields("CNT")))
                    Case "7":   Call .SetText(intCol, intStRow + 18, Trim(RS_Ser.Fields("CNT")))
                    Case "8":   Call .SetText(intCol, intStRow + 21, Trim(RS_Ser.Fields("CNT")))
                    Case "9":   Call .SetText(intCol, intStRow + 24, Trim(RS_Ser.Fields("CNT")))
                    Case "10":  Call .SetText(intCol, intStRow + 27, Trim(RS_Ser.Fields("CNT")))
                    Case "11":  Call .SetText(intCol, intStRow + 30, Trim(RS_Ser.Fields("CNT")))
                    Case "12":  Call .SetText(intCol, intStRow + 33, Trim(RS_Ser.Fields("CNT")))
                End Select
                RS_Ser.MoveNext
            Loop
            RS_Ser.Close
        Next
        '-- 초과건수 %
        'intStRow = 3
        For intCol = 3 To .MaxCols
            intStRow = 3
            For intCnt = 1 To 12
                Select Case CStr(intCnt)
                    Case "1":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "2":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "3":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "4":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "5":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "6":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "7":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "8":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "9":   Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "10":  Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "11":  Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                    Case "12":  Call .GetText(intCol, intStRow - 2, varTmp): lngTot = varTmp:       Call .GetText(intCol, intStRow - 1, varTmp): lngAmr = varTmp
                End Select
                If lngTot > 0 And lngAmr > 0 Then
                    Call .SetText(intCol, intStRow, Format((lngAmr / lngTot) * 100, "#0.00"))
                Else
                    Call .SetText(intCol, intStRow, "0")
                End If
                intStRow = intStRow + 3
            Next
            'Call SetSQLData("장비초과건수조회", SQL)
            Cn_Ser.CursorLocation = adUseClient
            Set RS_Ser = Cn_Ser.Execute(SQL)
        Next
    End With
    
    Screen.MousePointer = 0
        
End Sub

Private Sub FrmInitial()
    Dim DB_Tmp As String * 100

    DB_Tmp = ""
    
    spdMaster.MaxRows = 0
    
    dtpYear.Value = Date
    dtpToday.Value = Date
    
    DB_Tmp = ""
    Call GetPrivateProfileString("STA", "WIDTH", "", DB_Tmp, 100, App.Path & "\STA.ini")
    gWIDTH = Trim(DB_Tmp)

    DB_Tmp = ""
    Call GetPrivateProfileString("STA", "IP", "", DB_Tmp, 100, App.Path & "\STA.ini")
    gIP = Trim(DB_Tmp)

    DB_Tmp = ""
    Call GetPrivateProfileString("STA", "DB", "", DB_Tmp, 100, App.Path & "\STA.ini")
    gDB = Trim(DB_Tmp)

    DB_Tmp = ""
    Call GetPrivateProfileString("STA", "ID", "", DB_Tmp, 100, App.Path & "\STA.ini")
    gID = Trim(DB_Tmp)

    DB_Tmp = ""
    Call GetPrivateProfileString("STA", "PW", "", DB_Tmp, 100, App.Path & "\STA.ini")
    gPW = Trim(DB_Tmp)

End Sub

'-- 서버연결
Public Function Connect_PRServer() As Boolean

On Error GoTo errFind
    Connect_PRServer = False
    Set Cn_Ser = New ADODB.Connection
    With Cn_Ser
        .ConnectionTimeout = 25
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = gIP
        .Properties("Initial Catalog").Value = gDB
        .Properties("User ID").Value = gID
        .Properties("Password").Value = gPW
        .Open
    End With
    Connect_PRServer = True
    Exit Function

errFind:
    Connect_PRServer = False
End Function

Private Sub txtEventLog_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If txtEventLog.Text <> "" Then
            Call cmdSave_Click
        End If
    End If
    
End Sub
