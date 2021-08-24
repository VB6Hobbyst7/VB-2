VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "일자별마감현황"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   14880
   WindowState     =   2  '최대화
   Begin XLibrary_XGroupBox.XGroupBox grpMain 
      Height          =   9675
      Left            =   30
      Top             =   30
      Width           =   14790
      _ExtentX        =   26088
      _ExtentY        =   17066
      BackColor       =   16777215
      BorderColor     =   10070188
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TextPosition    =   0
      TextCustomMargin=   4
      GroupBoxStyle   =   0
      TextBarColor1   =   12757903
      TextBarStyle    =   3
      TextBarColor2   =   11767328
      TextBarSymbol   =   0   'False
      TextBarSymbolColor=   16777215
      TextBarHeightMargin=   10
      MouseCursor     =   0
      TextBarMouseCursor=   0
      IconandTextMargin=   4
      BodyColor       =   16777215
      Enabled         =   -1  'True
      Begin TeeChart.TChart tchChart 
         Height          =   3825
         Left            =   90
         TabIndex        =   2
         Top             =   5760
         Width           =   14640
         Base64          =   $"PIS301.frx":0000
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   12270
         TabIndex        =   5
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "조 회"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS301.frx":073A
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13500
         TabIndex        =   4
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "닫 기"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS301.frx":1EFC
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox1 
         Height          =   495
         Left            =   3000
         Top             =   90
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   873
         BackColor       =   16777215
         BorderColor     =   10070188
         BorderRoundNum  =   0
         BorderStyle     =   1
         TextColor       =   0
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         TextPosition    =   0
         TextCustomMargin=   4
         GroupBoxStyle   =   0
         TextBarColor1   =   12757903
         TextBarStyle    =   3
         TextBarColor2   =   11767328
         TextBarSymbol   =   0   'False
         TextBarSymbolColor=   16777215
         TextBarHeightMargin=   10
         MouseCursor     =   0
         TextBarMouseCursor=   0
         IconandTextMargin=   4
         BodyColor       =   16777215
         Enabled         =   -1  'True
         Begin XLibrary_XLabel.XLabel lblCount 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   90
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   556
            BackColor       =   16777215
            Text            =   "마감일자"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   0
            IconAndTextMargin=   8
            TextAlign       =   2
            TextAlignMargin =   0
            Focus           =   0   'False
            MouseCursor     =   0
            ToolTipIcon     =   0
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   -1
            ToolTipBackColor=   14811135
            ToolTipForeColor=   0
            ToolTipOpacity  =   100
            ToolTipStyle    =   0
            ToolTipCentered =   0   'False
            ToolTipTitleText=   "Title"
            ToolTipBodyText =   ""
            TextBackColor1  =   -2147483629
            TextBackColor2  =   -2147483629
            TextBackMargin  =   4
            TextBackStyle   =   0
            Enabled         =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   5025
         Left            =   90
         TabIndex        =   1
         Tag             =   "20001"
         Top             =   660
         Width           =   14640
         _Version        =   524288
         _ExtentX        =   25823
         _ExtentY        =   8864
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   7
         MaxRows         =   6
         Protect         =   0   'False
         ScrollBars      =   0
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS301.frx":36BE
         VisibleCols     =   3
         VisibleRows     =   6
         CellNoteIndicatorColor=   16576
      End
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox3 
         Height          =   495
         Left            =   90
         Top             =   90
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   873
         BackColor       =   16311512
         BorderColor     =   10070188
         BorderRoundNum  =   0
         BorderStyle     =   1
         TextColor       =   0
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         TextPosition    =   0
         TextCustomMargin=   4
         GroupBoxStyle   =   0
         TextBarColor1   =   12757903
         TextBarStyle    =   3
         TextBarColor2   =   11767328
         TextBarSymbol   =   0   'False
         TextBarSymbolColor=   16777215
         TextBarHeightMargin=   10
         MouseCursor     =   0
         TextBarMouseCursor=   0
         IconandTextMargin=   4
         BodyColor       =   16777215
         Enabled         =   -1  'True
         Begin TDBDate6Ctl.TDBDate dtpDt 
            Height          =   315
            Left            =   1620
            TabIndex        =   6
            Top             =   90
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   556
            Calendar        =   "PIS301.frx":3E58
            Caption         =   "PIS301.frx":3F3F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS301.frx":3FA2
            Keys            =   "PIS301.frx":3FC0
            Spin            =   "PIS301.frx":401E
            AlignHorizontal =   0
            AlignVertical   =   2
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "yyyy-mm"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   -2147483640
            Format          =   "yyyy-mm"
            HighlightText   =   0
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   5
            MarginRight     =   5
            MarginTop       =   1
            MaxDate         =   2958465
            MinDate         =   -657434
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "2015-07"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   42200
            CenturyMode     =   2
         End
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   210
            TabIndex        =   0
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "마감년월"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   0
            IconAndTextMargin=   8
            TextAlign       =   0
            TextAlignMargin =   0
            Focus           =   0   'False
            MouseCursor     =   0
            ToolTipIcon     =   0
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   -1
            ToolTipBackColor=   14811135
            ToolTipForeColor=   0
            ToolTipOpacity  =   100
            ToolTipStyle    =   0
            ToolTipCentered =   0   'False
            ToolTipTitleText=   "Title"
            ToolTipBodyText =   ""
            TextBackColor1  =   -2147483629
            TextBackColor2  =   -2147483629
            TextBackMargin  =   4
            TextBackStyle   =   0
            Enabled         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "PIS301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdFind_Click()
Dim cPis311 As clsPis311
Dim sDate As String, sLastDt As String, sDayCnt As Integer, sCnt As Integer, sRow As Integer, sCol As Integer
Dim sSetText As String, sSum As Long, sStartDt As String, sTotal As Long, sTestCnt(5) As Long

    tchChart.Header.Text.Clear
    tchChart.Series(0).Clear
    
    With spList
        .Row = 1:           .Col = 1
        .Row2 = .MaxRows:   .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With

    sDate = Format(dtpDt.Value, "yyyy-MM") & "-01"
    sLastDt = DateAdd("d", -1, DateAdd("m", 1, sDate))
    sDayCnt = DateDiff("d", sDate, sLastDt) + 1
    sStartDt = sDate
    
    Set cPis311 = New clsPis311
    
    sCol = Weekday(sDate)
    sRow = 1
    For sCnt = 1 To sDayCnt
        spList.Row = sRow
        spList.Col = sCol
        
        sDate = DateAdd("d", sCnt - 1, sStartDt)
        
        sSetText = sCnt & "일 "
        gSql = "SELECT SUM(TESTCNT) AS TESTCNT, SUM(FREECNT) AS FREECNT, SUM(QCCNT) AS QCCNT, SUM(RETESTCNT) AS RETESTCNT, SUM(MANUCNT) AS MANUCNT" & vbNewLine & _
               " FROM S2PIS311 WHERE WORKDT='" & Format(sDate, "yyyyMMdd") & "' GROUP BY WORKDT"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    sSum = Val("" & .Fields("TESTCNT").Value) + Val("" & .Fields("FREECNT").Value) _
                           + Val("" & .Fields("QCCNT").Value) + Val("" & .Fields("RETESTCNT").Value) + Val("" & .Fields("MANUCNT").Value)
                    
                    sSetText = sSetText & vbNewLine & "일반:" & pfNumString(.Fields("TESTCNT").Value)
                    sSetText = sSetText & vbNewLine & "무료:" & pfNumString(.Fields("FREECNT").Value)
                    sSetText = sSetText & vbNewLine & "전체:" & pfNumString(sSum)
                    
                    tchChart.Series(0).Add sSum, sCnt & "일", clTeeColor
                    
                    sTotal = sTotal + sSum
                    
                    sTestCnt(1) = sTestCnt(1) + Val("" & .Fields("TESTCNT").Value)
                    sTestCnt(2) = sTestCnt(2) + Val("" & .Fields("FREECNT").Value)
                    sTestCnt(3) = sTestCnt(3) + Val("" & .Fields("QCCNT").Value)
                    sTestCnt(4) = sTestCnt(4) + Val("" & .Fields("RETESTCNT").Value)
                    sTestCnt(5) = sTestCnt(5) + Val("" & .Fields("MANUCNT").Value)
                Else
                    tchChart.Series(0).Add 0, sCnt & "일", clTeeColor
                End If
                .Close
            End If
        End With
        
        spList.SetText sCol, sRow, sSetText
        
        sCol = sCol + 1
        If sCol > 7 Then
            sRow = sRow + 1
            sCol = 1
        End If
    Next sCnt
    
    tchChart.Header.Text.Add Format(dtpDt.Value, "yyyy") & "년 " & Format(dtpDt.Value, "M") & "월 검사건수현황 (" & Format(sTotal, "#,##0") & "건)"
    lblCount.Text = "■ 총건수: " & Format(sTotal, "#,##0")
    lblCount.Text = lblCount.Text & " ■ 일반: " & Format(sTestCnt(1), "#,##0")
    lblCount.Text = lblCount.Text & " ■ 무료: " & Format(sTestCnt(2), "#,##0")
    lblCount.Text = lblCount.Text & " ■ Q.C: " & Format(sTestCnt(3), "#,##0")
    lblCount.Text = lblCount.Text & " ■ 재검: " & Format(sTestCnt(4), "#,##0")
    lblCount.Text = lblCount.Text & " ■ 수기: " & Format(sTestCnt(5), "#,##0")

End Sub

Private Function pfNumString(ByVal brCnt As Long) As String
Dim sReturn As String

    sReturn = Format(brCnt, "#,##0")
    sReturn = Space(8 - Len(sReturn)) & sReturn & " "
    pfNumString = sReturn
    
End Function

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    dtpDt.Value = gfSystemDate
    tchChart.Header.Text.Clear
    lblCount.Text = ""
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    On Error Resume Next
    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2
    grpMain.Height = Me.ScaleHeight - 50
    tchChart.Height = (grpMain.Height - tchChart.Top) - 50
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub


