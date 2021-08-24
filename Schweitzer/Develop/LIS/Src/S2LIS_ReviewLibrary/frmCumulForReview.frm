VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCumulForReview 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "누적결과 조회 - 환자 ID : 00000001 , 환자명 : 홍길동"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   14520
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E4F3F8&
      Caption         =   "<< (&P)"
      Height          =   465
      Index           =   0
      Left            =   7065
      Style           =   1  '그래픽
      TabIndex        =   13
      Top             =   60
      Width           =   1380
   End
   Begin VB.CommandButton cmdPrintGrp 
      BackColor       =   &H00DBF2FD&
      Caption         =   "Print"
      Height          =   315
      Left            =   13575
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   4725
      Width           =   885
   End
   Begin VB.ListBox lstRemark 
      Height          =   2580
      Left            =   7860
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1695
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.ListBox lstDtTm 
      Height          =   2580
      Left            =   5205
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1695
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   435
      Left            =   13275
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "128"
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E4F3F8&
      Caption         =   "(&N) >>"
      Height          =   465
      Index           =   1
      Left            =   8505
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   60
      Width           =   1380
   End
   Begin VB.CheckBox chkGraph 
      BackColor       =   &H00DBE6E6&
      Caption         =   "그래프(&G)"
      ForeColor       =   &H00475765&
      Height          =   270
      Left            =   5640
      TabIndex        =   5
      Tag             =   "40201"
      Top             =   150
      Value           =   1  '확인
      Width           =   1260
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   435
      Left            =   10785
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "132"
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00FFF9F4&
      Caption         =   "조회(&Q)"
      Height          =   450
      Left            =   12015
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "133"
      Top             =   75
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   315
      Left            =   2850
      TabIndex        =   0
      Top             =   120
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
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
      Format          =   64028672
      CurrentDate     =   36567
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
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
      Format          =   64028672
      CurrentDate     =   36567
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   4065
      Left            =   45
      TabIndex        =   4
      Top             =   570
      Width           =   14445
      _Version        =   196608
      _ExtentX        =   25479
      _ExtentY        =   7170
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   17
      MaxRows         =   50
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14870761
      ShadowDark      =   14870761
      SpreadDesigner  =   "frmCumulForReview.frx":0000
      TextTip         =   4
   End
   Begin ChartfxLibCtl.ChartFX cfxResult 
      Height          =   2250
      Left            =   45
      TabIndex        =   11
      Top             =   4710
      Width           =   14445
      _cx             =   1205127
      _cy             =   1183617
      Build           =   7
      TypeMask        =   -1884749823
      Style           =   -67125249
      LeftGap         =   60
      RightGap        =   50
      TopGap          =   40
      BottomGap       =   31
      WallWidth       =   8
      View3DDepth     =   60
      TypeEx          =   32
      StyleEx         =   0
      DblClk          =   0
      RigClk          =   0
      MarkerShape     =   5
      MarkerSize      =   2
      Axis(0).MinorStep=   -1.2
      Axis(0).Max     =   6
      Axis(0).Decimals=   1
      Axis(0).TickMark=   -32767
      Axis(1).Min     =   0
      Axis(1).Max     =   100
      Axis(1).Decimals=   0
      Axis(1).Style   =   10344
      Axis(1).GridColor=   0
      Axis(2).Step    =   1
      Axis(2).MinorStep=   1
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).Style   =   14368
      Axis(2).PixPerUnit=   0
      RGBBk           =   14870761
      RGB2DBk         =   16777215
      RGB3DBk         =   14870761
      nColors         =   1
      Colors          =   "frmCumulForReview.frx":0D81
      TopFontMask     =   268435456
      BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BottomFontMask  =   268435456
      BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PointFontMask   =   268435456
      BeginProperty PointFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      nPts            =   25
      nSer            =   1
      NumPoint        =   25
      NumSer          =   1
      _Data_          =   "frmCumulForReview.frx":0DA9
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "~"
      Height          =   180
      Left            =   2565
      TabIndex        =   12
      Tag             =   "40110"
      Top             =   180
      Width           =   135
   End
End
Attribute VB_Name = "frmCumulForReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'새로운 결과조회화면에서 사용되는 누적결과조회
'Coding By legends 2003/10/02

Private objPatient As New clsPatient      '환자정보를 넘겨받는 object
Private objRvwSQL As New clsLISSqlReview 'Sql문 클래스
Private objRst As New clsLISResultReview

Private mCumCol As Collection

Private CallForm As Form

'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuCopy As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_COPY& = 1

Private Type tpItem
    TestCd As String
    PanelFg As String
    TestDiv As String
    SpcCd As String
    WorkArea As String
    TestNm As String
    SpcNm As String
    RefVal As String
End Type

Private tpItem() As tpItem

Private mItemCount As Long

Private lngPageNo As Long
Private lngPageCnt As Long
Private OldRow As Long
Private OldColor As Long

Private mvarCloseForm As Boolean

Public Sub SetPtinfo(ByRef objPt As clsPatient)
    Set objPatient = objPt
    '누적결과 조회 - 환자 ID : 00000001 , 환자명 : 홍길동
    Me.Caption = "누적 결과조회 - 환자 ID : " & objPatient.PTid & " , 환자명 : " & objPatient.PtNm
End Sub

Public Sub SetFromDate(ByVal vData As String)
    dtpFromDate.Value = vData
End Sub

Public Sub SetToDate(ByVal vData As String)
    dtpToDate.Value = vData
End Sub

Public Sub SetCallForm(ByRef objForm As Form)
    Set CallForm = objForm
End Sub

Public Property Let CloseForm(ByVal vData As Boolean)
    mvarCloseForm = vData
End Property

Public Property Get CloseForm() As Boolean
    CloseForm = mvarCloseForm
End Property

Private Sub chkGraph_Click()
    If chkGraph.Value = 1 Then
        cfxResult.Visible = True
        cmdPrintGrp.Visible = True
        tblResult.Height = 4065
        
        If Screen.ActiveControl.Name = chkGraph.Name Then
            If OldRow > 0 Then
                Call tblResult_Click(2, OldRow)
                tblResult.TopRow = OldRow
            Else
                Call tblResult_Click(2, 1)
                tblResult.TopRow = 1
            End If
        End If
    Else
        cfxResult.Visible = False
        cmdPrintGrp.Visible = False
        tblResult.Height = 6345
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click(Index As Integer)
    
    Select Case Index
    Case 0:
        lngPageNo = lngPageNo - 1
        If lngPageCnt > 1 Then cmdNext(1).Enabled = True
    Case 1:
        lngPageNo = lngPageNo + 1
        If lngPageCnt > 1 Then cmdNext(0).Enabled = True
    End Select
    
    Call DisplayOnePage(lngPageNo)
    
    If chkGraph.Value = 1 Then Call ShowGraph(OldRow)
    
    If lngPageNo = 1 Then
        cmdNext(0).Enabled = False
        If lngPageCnt > 1 Then cmdNext(1).Enabled = True
    End If
    If lngPageNo = lngPageCnt Then
        cmdNext(1).Enabled = False
        If lngPageCnt > 1 Then cmdNext(0).Enabled = True
    End If
End Sub

Private Sub cmdPrint_Click()

    With tblResult
        .PrintMarginTop = 100
        .PrintJobName = "누적결과레포트 출력"
        
        .PrintAbortMsg = "누적결과지를 출력중입니다. "

        .PrintOrientation = PrintOrientationLandscape
        If Printer.PaperSize = vbPRPSA4 Then
            .PrintMarginLeft = 1700
            .PrintMarginRight = 100
            .PrintMarginTop = 800
            .PrintMarginBottom = 800
        Else
            .PrintMarginTop = 300
            .PrintMarginBottom = 500
            .PrintMarginLeft = 250
            .PrintMarginRight = 100
        End If
        .PrintColor = False
        .PrintFirstPageNumber = 1
       
        .PrintHeader = "/n/n/l/fb1 " & "♧ 누적결과 - " & objPatient.PTid & "  " & objPatient.PtNm & "   " & _
                                        objPatient.Sex & "/" & objPatient.Age & " " & objPatient.AGEDIV & " /c/fb1/n/n"
        
        .PrintFooter = "/c/p/fb1"
        
        .PrintGrid = False
        .PrintShadows = False
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintPageEnd = 2
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        '.PrintGrid = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll
         
        .Action = ActionPrint
        .GridSolid = True
    End With
End Sub

Private Sub cmdPrintGrp_Click()
    Call PrintGraph
End Sub

Private Sub cmdQuery_Click()
    MousePointer = vbHourglass
    
    dtpFromDate.Enabled = False
    dtpToDate.Enabled = False
    cmdQuery.Enabled = False
    cmdPrint.Enabled = False

    Call DisplayResult
    
    dtpFromDate.Enabled = True
    dtpToDate.Enabled = True
    cmdQuery.Enabled = True
    cmdPrint.Enabled = True
    
    If lstDtTm.ListCount <= 0 Then
        MsgBox "해당 환자의 누적결과가 없습니다.", vbInformation
        If Screen.ActiveForm.Name <> Me.Name Then
            If mvarCloseForm Then
                Unload Me
                mvarCloseForm = True
            End If
        End If
    Else
        mvarCloseForm = False
    End If
    
    MousePointer = vbDefault
End Sub

Private Sub dtpFromDate_Click()
    If dtpFromDate.Value > GetSystemDate Then
        MsgBox "시작일이 현재날짜보다 큽니다. 다시 설정하십시오.", vbExclamation, "메세지"
        dtpFromDate.SetFocus
    End If
End Sub

Private Sub dtpFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdQuery.Enabled Then cmdQuery.SetFocus
    End If
End Sub

Private Sub Form_Activate()
'    Call medAlwaysOn(Me, 1)
End Sub

Private Sub Form_Load()
    Call InitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call medAlwaysOn(Me, 0)

    Set objRvwSQL = Nothing     'Sql문 클래스
    Set objPatient = Nothing  '환자정보를 넘겨받는 object
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_COPY
            Dim I As Long
            Dim strClip As String
            
            With tblResult
                .Row = OldRow
                .Col = 2: strClip = .Value
                .Col = 4: strClip = strClip & " ; " & .Value & " : "
                For I = 5 To 14
                    .Row = OldRow
                    .Col = I
                    If I >= .SelBlockCol And I <= .SelBlockCol2 Then
                        If Trim(.Value) <> "" Then
                            strClip = strClip & .Value
                            .Row = 0
                            .Col = I: strClip = strClip & "(" & Mid(.Value, 4, 5) & ")" & Space(2)
                        End If
                    End If
                Next
                .Row = OldRow
                .Col = 15: strClip = strClip & Space(3) & "단위(" & .Value & ")"
                .Col = 16: strClip = strClip & Space(3) & "기준치(" & .Value & ")"
            End With
            
            Clipboard.Clear
            Clipboard.SetText strClip
    End Select
End Sub

'Private Sub mnuCopy_Click()
'    Dim I As Long
'    Dim strClip As String
'
'    With tblResult
'        .Row = OldRow
'        .Col = 2: strClip = .Value
'        .Col = 4: strClip = strClip & " ; " & .Value & " : "
'        For I = 5 To 14
'            .Row = OldRow
'            .Col = I
'            If I >= .SelBlockCol And I <= .SelBlockCol2 Then
'                If Trim(.Value) <> "" Then
'                    strClip = strClip & .Value
'                    .Row = 0
'                    .Col = I: strClip = strClip & "(" & Mid(.Value, 4, 5) & ")" & Space(2)
'                End If
'            End If
'        Next
'        .Row = OldRow
'        .Col = 15: strClip = strClip & Space(3) & "단위(" & .Value & ")"
'        .Col = 16: strClip = strClip & Space(3) & "기준치(" & .Value & ")"
'    End With
'
'    Clipboard.Clear
'    Clipboard.SetText strClip
'End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim I As Long
    Dim sDPfg As String
    
    If Row = 0 Then Exit Sub
    If tblResult.DataRowCnt = 0 Then Exit Sub
    If Row = OldRow Then GoTo Skip1
    
    With tblResult
        .ReDraw = False
        If OldRow > 0 Then
            .Col = 2: .Col2 = .MaxCols
            .Row = OldRow: .Row2 = OldRow
            .BlockMode = True
            .FontBold = False
            .BackColor = OldColor
            .CellBorderType = 0
            .Action = ActionSetCellBorder
            .BlockMode = False
            .Col = 2: .BackColor = &HE2E8E9
            .Col = 4: .BackColor = &HEEF4F4
            .RowHeight(OldRow) = 12
            
            .Col = 17
            sDPfg = .Value
            For I = 1 To 10
                If medGetP(sDPfg, I, ":") <> "" Then
                    .Col = I + 4
                    .BackColor = &HC0FFFF
                End If
            Next
            
        End If
        .Row = Row: .Col = 1
        OldColor = .BackColor
        
        .Col = 2: .Col2 = .MaxCols
        .Row = Row:  .Row2 = Row
        .BlockMode = True
        .FontBold = True
        .BackColor = &HC0FFFF
        .CellBorderColor = &H80
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 16
        .Action = ActionSetCellBorder
        .BlockMode = False
        .RowHeight(Row) = 12
        OldRow = Row
        .ReDraw = True
    End With
    
Skip1:
    If chkGraph.Value = 1 Then Call ShowGraph(Row)
End Sub

Private Sub tblResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row = 0 Then Exit Sub
    If tblResult.DataRowCnt = 0 Then Exit Sub
    
    '이건 왜 있는건지 모르겠네..
    Call tblResult_Click(Col, Row)
    
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_COPY, "CLIPBOARD로 복사"
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
    
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuCopy = frmControls.mnuSub
'
'    frmControls.mnuSub1.Visible = False
'
'    mnuCopy.Caption = "Clipboard로 복사"
'
'    Me.PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuCopy = Nothing
'    Unload frmControls
'    Set frmControls = Nothing
End Sub

Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Row = 0 Then Exit Sub
    If tblResult.DataRowCnt = 0 Then Exit Sub
    
    If Col = 2 Or Col = 4 Or Col = 15 Then
        tblResult.Row = Row
        tblResult.Col = Col
        MultiLine = 1
        TipText = "  " & Trim(tblResult.Value)
        TipWidth = Len(TipText) * 150  '3000
        tblResult.TextTipDelay = 200
        Call tblResult.SetTextTipAppearance("Arial", 11, False, False, vbWhite, vbBlue)    '&H996666)
        ShowTip = True
    ElseIf Col >= 5 Then
        tblResult.Row = Row
        tblResult.Col = Col
        If Len(tblResult.Value) > 9 Then
            MultiLine = 1
            TipText = "  " & Trim(tblResult.Value)
            TipWidth = Len(TipText) * 150  '3000
            tblResult.TextTipDelay = 200
            Call tblResult.SetTextTipAppearance("Arial", 11, False, False, vbWhite, vbBlue)    '&H996666)
            ShowTip = True
        Else
            ShowTip = False
        End If
    End If
End Sub

Public Sub DisplayItem(ByRef objTestCd As clsDictionary, ByVal strPtId As String)
'검사항목,검체,참고치를 가져온다.

    Dim objSQL As clsLISSqlStatement
    Dim objPro As jProgressBar.clsProgress
    Dim RS As Recordset
    Dim rsRef As Recordset
    Dim I As Long
    Dim strSQL As String
    Dim dblRefFrom As Double, dblRefTo As Double
    Dim blnDupChk   As Boolean
    
    Set objPro = Nothing
    Set objPro = New jProgressBar.clsProgress
    
    With objPro
        .Container = CallForm
        .Width = CallForm.tblOrd2.Width '9660 '
        .Left = CallForm.tblOrd2.Left '0 ' 'CallForm.fraType.Left +
        .Top = CallForm.tblOrd2.Top '2010 ''CallForm.fraType.Top +
        .Height = 450
        .Message = "화면에 표시할 검사항목 정보를 읽고 있습니다..."
        .Max = objTestCd.RecordCount
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm CallForm
'        .XWidth = 9660 'CallForm.tblOrdSheet.Width
'        .XPos = 0 ' CallForm.tblOrdSheet.Left 'CallForm.fraType.Left +
'        .YPos = 2010 'CallForm.tblOrdSheet.Top 'CallForm.fraType.Top +
'        .YHeight = 450
'        .ForeColor = &H864B24
'        .Msg = "화면에 표시할 검사항목 정보를 읽고 있습니다."
'        .Value = 1
'        .Max = objTestCd.RecordCount
    End With
         
    Erase tpItem
    mItemCount = 0
    ReDim tpItem(mItemCount)
    
    tblResult.ReDraw = False
    
    Set objSQL = New clsLISSqlStatement
    
    objTestCd.MoveFirst
    Do Until objTestCd.EOF
        strSQL = objSQL.GetCumulative(objTestCd.Fields("testcd"), objTestCd.Fields("spccd"))
        
        Set RS = Nothing
        Set RS = New Recordset
        RS.Open strSQL, DBConn
        
        If RS.EOF Then GoTo Skip
        
        blnDupChk = False
        For I = 1 To tblResult.DataRowCnt
            tblResult.Row = I
            tblResult.Col = 1
            If objTestCd.Fields("testcd") = tblResult.Value Then
                blnDupChk = True
                Exit For
            End If
        Next
        
        If Not blnDupChk Then
            mItemCount = mItemCount + 1
            If tblResult.DataRowCnt + 1 > tblResult.MaxRows Then
                tblResult.MaxRows = tblResult.MaxRows + 1
            End If
            tblResult.Row = tblResult.DataRowCnt + 1
            tblResult.Col = 1: tblResult.Value = "" & RS.Fields("TestCd").Value
            tblResult.Col = 2: tblResult.Value = "" & RS.Fields("TestNm").Value
            tblResult.Col = 3: tblResult.Value = "" & RS.Fields("SpcCd").Value
            tblResult.Col = 4: tblResult.Value = "" & RS.Fields("SpcNm").Value
            'tblResult.Col = 16: tblResult.Value = .RefVal
        Else
            GoTo Skip
        End If
        
'        mItemCount = mItemCount + 1
        ReDim Preserve tpItem(mItemCount)
        
        objPro.Value = mItemCount
        
        With tpItem(mItemCount)
            .TestCd = Trim("" & RS.Fields("TestCd").Value)
            .PanelFg = Trim("" & RS.Fields("PanelFg").Value)
            .TestDiv = Trim("" & RS.Fields("TestDiv").Value)
            .SpcCd = Trim("" & RS.Fields("SpcCd").Value)
            .WorkArea = Trim("" & RS.Fields("WorkArea").Value)
            .TestNm = Trim("" & RS.Fields("TestNm").Value)
            .SpcNm = Trim("" & RS.Fields("SpcNm").Value)
            strSQL = objRvwSQL.SqlGetReference(.TestCd, .SpcCd, Format(GetSystemDate, CS_DateDbFormat), "B", _
                                            DateDiff("y", Format(objPatient.Dob, CS_DateMask), GetSystemDate))
            
            Set rsRef = Nothing
            Set rsRef = New Recordset
            rsRef.Open strSQL, DBConn
            If rsRef.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
               strSQL = objRvwSQL.SqlGetReference(.TestCd, .SpcCd, Format(GetSystemDate, CS_DateDbFormat), objPatient.Sex, _
                                            DateDiff("y", Format(objPatient.Dob, CS_DateMask), Now))
                Set rsRef = Nothing
                Set rsRef = New Recordset
                rsRef.Open strSQL, DBConn
            End If
            If rsRef.EOF Then
               .RefVal = ""
            Else
               dblRefFrom = Val("" & rsRef.Fields("RefValFrom").Value)
               dblRefTo = Val("" & rsRef.Fields("RefValTo").Value)
               .RefVal = Trim("" & rsRef.Fields("RefCd").Value)
               If dblRefFrom <> 0 Or dblRefTo <> 0 Then .RefVal = dblRefFrom & " - " & dblRefTo
            End If
            Set rsRef = Nothing
            tblResult.Col = 16: tblResult.Value = .RefVal
            
            
'            tblResult.MaxRows = mItemCount
'            tblResult.Row = mItemCount
'            tblResult.Col = 1: tblResult.Value = .TestCd
'            tblResult.Col = 2: tblResult.Value = .TestNm
'            tblResult.Col = 3: tblResult.Value = .SpcCd
'            tblResult.Col = 4: tblResult.Value = .SpcNm
'            tblResult.Col = 16: tblResult.Value = .RefVal
            
            
            
            
        End With
        If tpItem(mItemCount).PanelFg = PN_Detail Then
            Call DisplayDetail(tpItem(mItemCount).TestCd, tpItem(mItemCount).SpcCd, tpItem(mItemCount).SpcNm)
        End If
                
        Set RS = Nothing
Skip:
        objTestCd.MoveNext
    Loop
   
    Call ChangeColor
    
    tblResult.ReDraw = True
    
    Set RS = Nothing
    Set rsRef = Nothing
    Set objSQL = Nothing
    Set objPro = Nothing
    
    Call cmdQuery_Click
    
'    CallForm.fraType.Refresh
End Sub

Private Sub ChangeColor()
'스프레드 컬러변경
    With tblResult
        .Row = -1
        
        .Col = 2: .Col2 = 2
        .BlockMode = True
        .ForeColor = &H864B24
        .BackColor = &HE2E8E9
        .BlockMode = False
        
        .Col = 4: .Col2 = 4
        .BlockMode = True
        .ForeColor = &H808080
        .BackColor = &HEEF4F4
        .BlockMode = False
        
        .Col = 15: .Col2 = 15
        .BlockMode = True
        .ForeColor = &H80
        .BlockMode = False
        
        .Col = 16: .Col2 = 16
        .BlockMode = True
        .ForeColor = &H136604
        .BlockMode = False
        
        .RowHeight(-1) = 12
    End With
End Sub

Private Sub DisplayDetail(ByVal pTestCd As String, ByVal pSpcCd As String, ByVal pSpcNm As String)
'검사항목,검체,참고치를 가져온다.
    Dim RS      As Recordset
    Dim rsRef   As Recordset
    Dim I       As Long
    Dim strSQL  As String
    Dim dblRefFrom As Double, dblRefTo As Double
    Dim blnDupChk   As Boolean
    
    strSQL = objRvwSQL.SqlGetCumDetail(pTestCd)
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    While (Not RS.EOF)
        blnDupChk = False
        For I = 1 To tblResult.DataRowCnt
            tblResult.Row = I
            tblResult.Col = 1
            If "" & RS.Fields("TestCd").Value = tblResult.Value Then
                blnDupChk = True
                Exit For
            End If
        Next
        
        If Not blnDupChk Then
            mItemCount = mItemCount + 1
            If tblResult.DataRowCnt + 1 > tblResult.MaxRows Then
                tblResult.MaxRows = tblResult.MaxRows + 1
            End If
            tblResult.Row = tblResult.DataRowCnt + 1
            tblResult.Col = 1: tblResult.Value = "" & RS.Fields("TestCd").Value
            tblResult.Col = 2: tblResult.Value = "    " & "" & RS.Fields("TestNm").Value
            tblResult.Col = 3: tblResult.Value = pSpcCd
            tblResult.Col = 4: tblResult.Value = pSpcNm
            
        Else
            GoTo Skip
        End If
        
'        mItemCount = mItemCount + 1
        ReDim Preserve tpItem(mItemCount)
        With tpItem(mItemCount)
            .TestCd = Trim("" & RS.Fields("TestCd").Value)
            .PanelFg = Trim("" & RS.Fields("PanelFg").Value)
            .TestDiv = Trim("" & RS.Fields("TestDiv").Value)
            .SpcCd = pSpcCd
            .SpcNm = pSpcNm
            .WorkArea = Trim("" & RS.Fields("WorkArea").Value)
            .TestNm = "    " & Trim("" & RS.Fields("TestNm").Value)
            strSQL = objRvwSQL.SqlGetReference(.TestCd, .SpcCd, Format(GetSystemDate, CS_DateDbFormat), "B", _
                                            DateDiff("y", Format(objPatient.Dob, CS_DateMask), Now))
            Set rsRef = Nothing
            Set rsRef = New Recordset
            rsRef.Open strSQL, DBConn
            If rsRef.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
               strSQL = objRvwSQL.SqlGetReference(.TestCd, .SpcCd, Format(GetSystemDate, CS_DateDbFormat), objPatient.Sex, _
                                            DateDiff("y", Format(objPatient.Dob, CS_DateMask), Now))
                Set rsRef = Nothing
                Set rsRef = New Recordset
                rsRef.Open strSQL, DBConn
            End If
            If rsRef.EOF Then
               .RefVal = ""
            Else
               dblRefFrom = Val("" & rsRef.Fields("RefValFrom").Value)
               dblRefTo = Val("" & rsRef.Fields("RefValTo").Value)
               .RefVal = Trim("" & rsRef.Fields("RefCd").Value)
               If dblRefFrom <> 0 Or dblRefTo <> 0 Then .RefVal = dblRefFrom & " - " & dblRefTo
            End If
            Set rsRef = Nothing
            tblResult.Col = 16: tblResult.Value = .RefVal
        End With
Skip:
        RS.MoveNext
    Wend

    Set RS = Nothing
    Set rsRef = Nothing
End Sub

Private Sub DisplayResult()
    lngPageNo = 0
    lngPageCnt = 0
    
    Call ReadData
    
    If lstDtTm.ListCount = 0 Then Exit Sub
    
    lngPageCnt = (lstDtTm.ListCount + 9) \ 10
    
    Call cmdNext_Click(1)
End Sub

Private Sub ReadData()
    Dim objPro As jProgressBar.clsProgress
    Dim ObjDic As New clsDictionary
    Dim clsNewData As clsCumResult
    Dim RS As Recordset
    Dim strSQL As String
    Dim strFromDt As String, strToDt As String
    Dim strSpcCd As String, strWorkArea As String, strTestNm As String
    Dim strTestcd As String, strPanelFg As String, strTestDiv As String
    Dim strDtTm As String
    Dim lngseq  As Long
    Dim strList As String
    Dim I As Long, J As Long
    
    Set objPro = Nothing
    Set objPro = New jProgressBar.clsProgress
    
'    Debug.Print Screen.ActiveForm.Name
    
    With objPro
        .Container = IIf(Screen.ActiveForm.Name = Me.Name, Me, CallForm)
        .Width = IIf(Screen.ActiveForm.Name = Me.Name, tblResult.Width, CallForm.tblOrd2.Width) '9660)
        .Left = IIf(Screen.ActiveForm.Name = Me.Name, tblResult.Left, CallForm.tblOrd2.Left) ' 0) 'CallForm.fraType.Left +
        .Top = IIf(Screen.ActiveForm.Name = Me.Name, tblResult.Top, CallForm.tblOrd2.Top) ' 2010) 'CallForm.fraType.Top +
        .Height = IIf(Screen.ActiveForm.Name = Me.Name, 630, 450)
        .Message = "자료를 읽기 위해 준비중입니다..."
        .Max = mItemCount
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm IIf(Screen.ActiveForm.Name = Me.Name, Me, CallForm)
'        .XWidth = IIf(Screen.ActiveForm.Name = Me.Name, tblResult.Width, 9660)
'        .XPos = IIf(Screen.ActiveForm.Name = Me.Name, tblResult.Left, 0) 'CallForm.fraType.Left +
'        .YPos = IIf(Screen.ActiveForm.Name = Me.Name, tblResult.Top, 2010) 'CallForm.fraType.Top +
'        .YHeight = IIf(Screen.ActiveForm.Name = Me.Name, 630, 450)
'        .ForeColor = &H864B24
'        .Msg = "자료를 읽기 위해 준비중입니다..."
'        .Value = 1
'        .Max = mItemCount
    End With
    
    ObjDic.Clear
    ObjDic.FieldInialize "strDtTm, strTestCd, lngSeq", "seq"
    
    strFromDt = Format(dtpFromDate.Value, CS_DateDbFormat)
    strToDt = Format(dtpToDate.Value, CS_DateDbFormat)
    
    lstDtTm.Clear
    lstRemark.Clear
    
    Set mCumCol = New Collection
    For I = 1 To mItemCount
    
        objPro.Value = I
        objPro.Message = tpItem(I).TestNm & " 항목을 읽고 있습니다..."
        DoEvents
        
        strTestcd = tpItem(I).TestCd
        strPanelFg = tpItem(I).PanelFg
        strTestDiv = tpItem(I).TestDiv
        strSpcCd = tpItem(I).SpcCd
        strWorkArea = tpItem(I).WorkArea
        strTestNm = tpItem(I).TestNm
        strSQL = objRvwSQL.SqlCumResult(objPatient.PTid, strFromDt, strToDt, strTestcd, _
                                     strPanelFg, strTestDiv, strSpcCd, strWorkArea)
        If Trim(strSQL) <> "" Then
            Set RS = Nothing
            Set RS = New Recordset
            RS.Open strSQL, DBConn
                    
            While (Not RS.EOF)
                DoEvents
                
                lngseq = 0
                strDtTm = Format("" & RS.Fields("ColDt").Value, CS_DateMask) & "  " & _
                          Format("" & RS.Fields("ColTm").Value, CS_TimeShortMask)
                strList = strDtTm & vbTab & Trim(CStr(lngseq))
                If medListFind(lstDtTm, strList) < 0 Then lstDtTm.AddItem strList
                
                
                Set clsNewData = New clsCumResult
                With clsNewData
                    .ColDt = Format("" & RS.Fields("ColDt").Value, CS_DateShortMask)
                    .ColTm = Format("" & RS.Fields("ColTm").Value, CS_TimeShortMask)
                    .DeptCd = "" & RS.Fields("DeptCd").Value
                    .WardId = "" & RS.Fields("WardId").Value
                    .HosilId = "" & RS.Fields("HosilId").Value
                    .TestCd = "" & RS.Fields("TestCd").Value
                    .TestNm = strTestNm
                    .SpcCd = "" & RS.Fields("SpcCd").Value
                    .SpcNm = "" & RS.Fields("TestCd").Value
                    .DPDiv = "" & RS.Fields("DpDiv").Value
                    .HLDiv = "" & RS.Fields("HlDiv").Value
                    .RstDiv = "" & RS.Fields("RstDiv").Value
                    .WorkArea = "" & RS.Fields("WorkArea").Value
                    .AccDt = "" & RS.Fields("AccDt").Value
                    .AccSeq = "" & RS.Fields("AccSeq").Value
                    .FootNoteFg = "" & RS.Fields("FootNoteFg").Value
                    .RmkCd = "" & RS.Fields("RmkCd").Value
                    If Trim("" & RS.Fields("RstCdNm").Value) <> "" Then
                        .RstCd = "" & RS.Fields("RstCdNm").Value
                    Else
                        .RstCd = "" & RS.Fields("RstCd").Value
                    End If
                    .RstUnit = "" & RS.Fields("RstUnit").Value
                    .TxtFg = "" & RS.Fields("TxtFg").Value
                    .Remark = ""
                    '리마크랑 풋노트는 화면에 표시해주지 않는 관계로 막아버림..
    '                If Trim(.RmkCd) <> "" Then
    '                    .Remark = objRst.ReadRemark(.RmkCd)
    '                    If medListFind(lstRemark, strDtTm) < 0 Then lstRemark.AddItem strDtTm & vbTab & CStr(mCumCol.Count + 1)
    '                End If
    '                .FootNote = ""
    '                If Trim(.FootNoteFg) <> "0" Then
    '                    .FootNote = objRst.ReadFootNote(.WorkArea, .AccDt, .AccSeq)
    '                    If medListFind(lstRemark, strDtTm) < 0 Then lstRemark.AddItem strDtTm & vbTab & CStr(mCumCol.Count + 1)
    '                End If
                        
                    .RstText = "" & RS.Fields("TextResult").Value
                    
                    On Error GoTo Dup_Err
                                 
                    If ObjDic.Exists(strDtTm & COL_DIV & .TestCd & Trim(CStr(lngseq))) = False Then
                        ObjDic.AddNew strDtTm & COL_DIV & .TestCd & Trim(CStr(lngseq)), Trim(CStr(lngseq))
                        
                        mCumCol.Add clsNewData, strDtTm & ":" & .TestCd & ":" & Trim(CStr(lngseq))
                    End If
                End With
                
                RS.MoveNext
            Wend
            
        End If
        Set RS = Nothing
    Next
    
    Set ObjDic = Nothing
    Set objPro = Nothing
    Exit Sub
    
Dup_Err:
    If Err.Number = 457 Then
        lngseq = lngseq + 1
        strList = strDtTm & vbTab & Trim(CStr(lngseq))
        If medListFind(lstDtTm, strList) < 0 Then lstDtTm.AddItem strList
        Resume
    Else
        MsgBox Err.Number & "  " & Err.Description, vbCritical
        Set RS = Nothing
    End If
    
    Set ObjDic = Nothing
    Set objPro = Nothing
End Sub

Private Sub DisplayOnePage(ByVal iCurPage As Integer)
    Dim I As Integer
    Dim J As Integer
    Dim iListIndex As Integer
    Dim sDtTm As String
    Dim sSeq As String
    Dim sDPfg As String
    Dim clsData As clsCumResult
    Dim ErrFg As Boolean
    Dim EvenBkColor As Long, OddBkColor As Long
    
    EvenBkColor = &HF9FBFA
    OddBkColor = &HFFFFFF
    
    With tblResult
        .Row = 0: .Row2 = .MaxRows
        .Col = 5: .Col2 = 14
        .BlockMode = True
        .Text = ""
        .BlockMode = False
        
        For I = 1 To .MaxRows
            .Row = I: .Row2 = I
            .Col = 5: .Col2 = 14
            .BlockMode = True
            If I <> OldRow Then
                .BackColor = IIf((I Mod 2) = 0, EvenBkColor, OddBkColor)
            End If
            .ForeColor = vbBlack
            .BlockMode = False
            .Col = 17: .Value = ""
        Next
        
        For I = (iCurPage - 1) * 10 To iCurPage * 10 - 1
            .Row = 0
            If I >= lstDtTm.ListCount Then Exit For
            
            '가장 최근날짜부터 Display하기 위해 Index계산...
            iListIndex = lstDtTm.ListCount - I - 1
            
            .Col = I - ((iCurPage - 1) * 10) + 5
            .Text = Format(medGetP(lstDtTm.List(iListIndex), 1, vbTab), CS_DateShortFormat & "  " & CS_TimeShortFormat)
            sDtTm = medGetP(lstDtTm.List(iListIndex), 1, vbTab)
            sSeq = medGetP(lstDtTm.List(iListIndex), 2, vbTab)
            For J = 1 To mItemCount
                
                sDPfg = ""
                ErrFg = False
                
                On Error GoTo Err_Trap
    
                .Row = J
                .Col = I - ((iCurPage - 1) * 10) + 5
                Set clsData = mCumCol.Item(sDtTm & ":" & tpItem(J).TestCd & ":" & sSeq)
                If ErrFg Then GoTo Skip
                .Value = clsData.RstCd
                If clsData.HLDiv = "H" Then
                    .ForeColor = &H7477EF   'vbRed
                ElseIf clsData.HLDiv = "L" Then
                    .ForeColor = &HDF6A3E  'vbBlue
                End If
                If clsData.DPDiv <> "" Then
                    sDPfg = clsData.DPDiv
                    .Value = .Value & " " & clsData.DPDiv
                    .ForeColor = vbRed
                    .BackColor = &HC0FFFF     '&HFFF7FF
                End If
                .Col = 15
                If .Value = "" Then .Value = clsData.RstUnit
Skip:
                .Col = 17
                .Value = .Value & sDPfg & ":"
                
            Next
            
            If I = (iCurPage - 1) * 10 Then
            '일단 막어놓구.. 낭중에 다시 봐야지..
'                If ObjLISComCode.DeptCd.Exists(clsData.DeptCd) Then
'                    ObjLISComCode.DeptCd.KeyChange (clsData.DeptCd)
'                    lblDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'                    'lblDeptNm.Caption = objPatient.GetDeptNm(clsData.DeptCd)
'                End If
'                lblWardId.Caption = clsData.WardId & " - " & clsData.HosilId
            End If
                
        Next
        '.ReDraw = True
    End With
    Set clsData = Nothing
    Exit Sub
    
Err_Trap:
    ErrFg = True
    Resume Next
End Sub

Private Sub InitForm()
    Set mCumCol = New Collection
    Erase tpItem
    mItemCount = 0
    ReDim tpItem(mItemCount)
    
    chkGraph.Value = 0
    
    Call medClearTable(tblResult)
'    Call ClearTable
    Call ClearGraph
    
    cmdNext(0).Enabled = False
    cmdNext(1).Enabled = False
    
    OldRow = -1
    
    lngPageNo = 0
    lngPageCnt = 0
    
    lstDtTm.Clear
    lstRemark.Clear
End Sub

Private Sub ClearTable()
    With tblResult
        .MaxRows = 0
        .Col = 5: .Col2 = 14
        .Row = 0: .Row2 = 0
        .BlockMode = True
        .Text = ""
        .BlockMode = False
    End With
End Sub

Private Sub ClearGraph()
    With cfxResult
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
    End With
End Sub


Private Sub ShowGraph(ByVal iGrpRow As Integer)
'여기도 타네
    Dim I As Integer, J As Integer
    Dim FirstFg As Boolean
    Dim iSeries As Integer, iPoints As Integer
    Dim iMaxValue As Double, iMinValue As Double
    Dim iFromRef As Double, iToRef As Double
    Dim sPnt As Integer, ePnt As Integer
    Dim sXVal As Integer, eXVal As Integer
    Dim tmpStr As String
    Dim clsData As clsCumResult
    Dim ErrFg As Boolean
    Dim sDtTm  As String, sSeq As String
    
    
    FirstFg = True
    
    iSeries = 1
    iPoints = 0
    
    Call SetDateRange(sPnt, ePnt)
    Call ClearGraph
    
    With tblResult
        .Row = iGrpRow: .Col = 2
        cfxResult.Title(CHART_TOPTIT) = .Value
        cfxResult.ClearData CD_VALUES
        
        cfxResult.RealTimeStyle = CRT_LOOPPOS Or CRT_NOWAITARROW
        cfxResult.OpenDataEx COD_VALUES, iSeries, lstDtTm.ListCount
        
        cfxResult.TopGap = 20
        cfxResult.BottomGap = 25
        cfxResult.FixedGap = 33
        cfxResult.Grid = CHART_NOGRID
        cfxResult.Scrollable = True
        
        .Row = iGrpRow: .Col = 16
        iFromRef = Val(medGetP(.Value, 1, "-"))
        iToRef = Val(medGetP(.Value, 2, "-"))
        iMinValue = iFromRef '- (iFromRef / 50) '2
        iMaxValue = iToRef '+ (iFromRef / 50) '2
         
        cfxResult.Scrollable = True
        cfxResult.PointLabels = True
        cfxResult.RGBFont(CHART_POINTFT) = vbBlue
        cfxResult.Axis(AXIS_X).STEP = 1
        
        cfxResult.ThisSerie = 0
        For I = lstDtTm.ListCount - 1 To 0 Step -1
            
            sDtTm = medGetP(lstDtTm.List(I), 1, vbTab)
            sSeq = medGetP(lstDtTm.List(I), 2, vbTab)
            
            ErrFg = False
                
            On Error GoTo Err_Trap
    
            Set clsData = mCumCol.Item(sDtTm & ":" & tpItem(iGrpRow).TestCd & ":" & sSeq)
            If ErrFg Then GoTo Skip
            If Not IsNumeric(clsData.RstCd) Then GoTo Skip
            
            cfxResult.KeyLeg(iPoints) = Format(sDtTm, "MM-DD")
            cfxResult.Value(iPoints) = Val(clsData.RstCd)
'            cfxResult
            iPoints = iPoints + 1
            
            If I = sPnt Then sXVal = iPoints
            If I = ePnt Then eXVal = iPoints
            
            If iMinValue > Val(clsData.RstCd) Then iMinValue = Val(clsData.RstCd)
            If iMaxValue < Val(clsData.RstCd) Then iMaxValue = Val(clsData.RstCd)
                    
Skip:
        Next
        
        If iPoints = 0 Then
            Call ClearGraph
            Exit Sub
        End If
        
        cfxResult.CloseData COD_VALUES
        
        cfxResult.OpenDataEx COD_STRIPES, 2, 0
        '참고치 구간 표시...
        cfxResult.Stripe(0).Axis = AXIS_Y
        cfxResult.Stripe(0).COLOR = &HC0FFFF
        cfxResult.Stripe(0).From = iFromRef
        cfxResult.Stripe(0).To = iToRef
        'Spread에 보여지고 있는 구간 표시...
        cfxResult.Stripe(1).Axis = AXIS_X
        cfxResult.Stripe(1).COLOR = &HDBF2FD          '&HD6EAFA       ' &HD6EAFA        '&HFFF9F4     '&HF4FEED   '&HD6D7FA     '&HFFF4FF  '&HF7FFFF  '&HEEF4F4  '&HEEEEEE
        cfxResult.Stripe(1).From = sXVal
        cfxResult.Stripe(1).To = eXVal
        cfxResult.CloseData COD_STRIPES
        
        cfxResult.OpenDataEx COD_CONSTANTS, 2, 0
        
        cfxResult.ConstantLine(0).Value = iFromRef
        cfxResult.ConstantLine(0).LineColor = &H808080
        cfxResult.ConstantLine(0).Axis = AXIS_Y
        cfxResult.ConstantLine(0).Label = CStr(iFromRef)
        cfxResult.ConstantLine(0).LineWidth = 1
        cfxResult.ConstantLine(0).LineStyle = CHART_DOT
        
        cfxResult.ConstantLine(1).Value = iToRef
        cfxResult.ConstantLine(1).LineColor = &H808080  '&H80&
        cfxResult.ConstantLine(1).Axis = AXIS_Y
        cfxResult.ConstantLine(1).Label = CStr(iToRef)
        cfxResult.ConstantLine(1).LineWidth = 1
        cfxResult.ConstantLine(1).LineStyle = CHART_DOT
        
        cfxResult.CloseData COD_CONSTANTS
        
        cfxResult.OpenDataEx COD_VALUES, iSeries, iPoints
            
        cfxResult.Axis(AXIS_Y).Min = iMinValue - ((iMaxValue - iFromRef) / 10) '1
        cfxResult.Axis(AXIS_Y).Max = iMaxValue + ((iMaxValue - iFromRef) / 10) '1
        
        cfxResult.Axis(AXIS_Y).STEP = (iMaxValue - iMinValue) / 3
        
        cfxResult.CloseData COD_VALUES
    
    End With
    Exit Sub

Err_Trap:
    ErrFg = True
    Resume Next
End Sub

Private Sub SetDateRange(sPnt As Integer, ePnt As Integer)
'여기도 타네
    Dim I As Integer
    Dim sDt As String, eDt As String
    
    With tblResult
        For I = 1 To 10
            .Row = OldRow
            .Col = I + 4
            If IsNumeric(.Value) Then
                sPnt = lstDtTm.ListCount - ((lngPageNo - 1) * 10) - I
                Exit For
            End If
        Next
        For I = 10 To 1 Step -1
            .Row = OldRow
            .Col = I + 4
            If IsNumeric(.Value) Then
                ePnt = lstDtTm.ListCount - ((lngPageNo - 1) * 10) - I
                Exit For
            End If
        Next
    End With
    
End Sub

Private Sub PrintGraph()
    With cfxResult
        .Printer.TopMargin = 2
        .Printer.LeftMargin = 0
        .Printer.RightMargin = 1
        .Printer.BottomMargin = 2
        .Printer.Compress = True
        .Printer.Orientation = ORIENTATION_LANDSCAPE
        .Printer.ForceColors = True
        .PrintIt 0, 0
    End With
    
End Sub
