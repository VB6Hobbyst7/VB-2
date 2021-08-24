VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS110 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E8EEEE&
   Caption         =   "Trans Request List"
   ClientHeight    =   8205
   ClientLeft      =   9465
   ClientTop       =   1995
   ClientWidth     =   7080
   ClipControls    =   0   'False
   Icon            =   "frmBBS110.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optDiv 
      BackColor       =   &H00E8EEEE&
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   1275
      TabIndex        =   3
      Top             =   300
      Width           =   900
   End
   Begin VB.OptionButton optDiv 
      BackColor       =   &H00E8EEEE&
      Caption         =   "OLD"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1275
      TabIndex        =   7
      Top             =   75
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.CheckBox chkCancel 
      BackColor       =   &H00E8EEEE&
      Caption         =   "취소목록"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3630
      TabIndex        =   6
      Top             =   195
      Width           =   1050
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C8CEDF&
      Height          =   375
      Left            =   4710
      Picture         =   "frmBBS110.frx":000C
      Style           =   1  '그래픽
      TabIndex        =   5
      ToolTipText     =   "혈액Tag 재출력"
      Top             =   135
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CheckBox chkStat 
      BackColor       =   &H00E8EEEE&
      Caption         =   "응급만"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   2595
      TabIndex        =   2
      Top             =   195
      Width           =   840
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1995
      Top             =   3795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS110.frx":053E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS110.frx":0862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS110.frx":0B7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   3810
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C8CEDF&
      Caption         =   "Refresh"
      Height          =   390
      Left            =   5790
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   7605
      Left            =   30
      TabIndex        =   0
      Top             =   540
      Width           =   7020
      _Version        =   196608
      _ExtentX        =   12382
      _ExtentY        =   13414
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
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
      GrayAreaBackColor=   15265518
      GridColor       =   16703181
      GridShowVert    =   0   'False
      MaxCols         =   14
      MaxRows         =   10
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS110.frx":0EA2
      TextTip         =   2
   End
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   2700
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   3825
      Width           =   855
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   600
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   510
      Left            =   30
      Shape           =   4  '둥근 사각형
      Top             =   0
      Width           =   510
   End
   Begin VB.Image imgSound 
      Height          =   480
      Index           =   1
      Left            =   45
      Picture         =   "frmBBS110.frx":154B
      Top             =   15
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSound 
      Height          =   480
      Index           =   0
      Left            =   45
      Picture         =   "frmBBS110.frx":1855
      Top             =   15
      Width           =   480
   End
End
Attribute VB_Name = "frmBBS110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Public Event LastFormUnload()
'Public Event ThisFormUnload()
'Public Event ListSelected(ByVal SelPtId As String, ByVal SelFrDt As String, ByVal SelToDt As String)
'Public Event MouseMove()
'
'Private blnInitFg As Boolean
'Private blnStopFg As Boolean
'
'Private Type NOTIFYICONDATA
'    cbSize As Long
'    hwnd As Long
'    uId As Long
'    uFlags As Long
'    ucallbackMessage As Long
'    hIcon As Long
'    szTip As String * 64
'End Type
'
'Private Const NIM_ADD = &H0
'Private Const NIM_MODIFY = &H1
'Private Const NIM_DELETE = &H2
'Private Const NIF_MESSAGE = &H1
'Private Const NIF_ICON = &H2
'Private Const NIF_TIP = &H4
'
'Private Const WM_LBUTTONDBLCLK = &H203
'Private Const WM_LBUTTONDOWN = &H201
'Private Const WM_LBUTTONUP = &H202
'Private Const WM_MBUTTONDBLCLK = &H209
'Private Const WM_MBUTTONDOWN = &H207
'Private Const WM_MBUTTONUP = &H208
'Private Const WM_RBUTTONDBLCLK = &H206
'Private Const WM_RBUTTONDOWN = &H204
'Private Const WM_RBUTTONUP = &H205
'
'Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Private TrayI As NOTIFYICONDATA
'Private blnSound As Boolean
'Private blnFormLoad As Boolean
'
'
'
'Private Sub cmdPrint_Click()
'
''출력하자.....크리스탈
'    Dim strPtid As String, strPtNm As String, strABO As String, strOrdNm As String, strReqCnt As String, strReqDt As String
'    Dim strStat As String, strSpcNo As String, strLoc As String
'    Dim StrWARD As String, StrACC As String, strReason As String
'    Dim strTmp  As String
'
'    Dim strRfile   As String
'    Dim strRptPath As String
'    Dim intFNum    As Integer
'    Dim ii         As Integer
'
'    If tblPtList.MaxRows = 0 Then Exit Sub
'    Me.MousePointer = 11
'    strTmp = ""
'
'    With tblPtList
'        For ii = 1 To .MaxRows
'            .Row = ii
'            .Col = 1:   strPtid = .value
'            .Col = 2:   strPtNm = .value
'            .Col = 3:   strABO = .value
'            .Col = 4:   strOrdNm = .value
'            .Col = 5:   strReqCnt = .value
'            .Col = 6:   strSpcNo = .value
'            .Col = 7:   StrACC = .value
'            .Col = 8:   strLoc = .value
'            .Col = 9:   StrWARD = .value
'            .Col = 10:  StrWARD = StrWARD & "-" & .value
'            .Col = 11:  strReason = .value
'            .Col = 12:  strStat = .value
'            .Col = 13:  strReqDt = .value
'
'            strTmp = strTmp & strPtid & vbTab & strPtNm & vbTab & strABO & vbTab & strOrdNm & vbTab & strReqCnt & vbTab & _
'                     strStat & vbTab & strLoc & vbTab & strSpcNo & vbTab & StrACC & vbTab & StrWARD & vbTab & strReqDt & vbTab & _
'                     strReason & vbCr
'
'        Next ii
'    End With
'
'    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
'
'    strRfile = InstallDir & "BBS\Rpt" & "\CrystalReport.txt"
'    strRptPath = InstallDir & "BBS\Rpt" & "\frmBBS110.rpt"
'
'    intFNum = FreeFile
'    Open strRfile For Output As #intFNum
'    Print #intFNum, strTmp
'    Close #intFNum
'    With CReport
'        .ReportFileName = strRptPath
'        .ParameterFields(0) = "buildingnm;" & ObjSysInfo.BuildingNm & ";TRUE"
'        .RetrieveDataFiles
'        .WindowState = crptMaximized
'        .Action = 1
'        .Reset
'    End With
'
'    Me.MousePointer = 0
'
'End Sub
'
'Private Sub cmdRefresh_Click()
'
'    Me.MousePointer = 11
'
'    blnStopFg = True
''    mmSound.Command = "Stop"
''    mmSound.Command = "Close"
'    DoEvents
'
'    Timer1.Enabled = True
''    Query
'    '길병원은 적용후일정 시점이 지나면 아래 쿼리를 사용할예정임
'    If optDiv(0).value Then
'        Query
'    Else
'        Query_New
'    End If
'
'
''    Call Query
''    If blnFormLoad = True Then SetCancel
'
'    Me.MousePointer = 0
'
'
'End Sub
'
'Private Sub Form_Load()
'
'
''    cmdAll.Caption = "All"
''    cmdAll.tag = "1"
'    blnInitFg = False
'    blnStopFg = False
'    blnSound = True
'    imgSound(0).Visible = True
'    imgSound(1).Visible = False
''    optDeptDiv(0).value = True
'
''    mmSound.Notify = False
''    mmSound.Wait = True
''    mmSound.Shareable = False
''    mmSound.DeviceType = "WaveAudio"
''    mmSound.FileName = gBloodRequestMusic
''    mmSound.Enabled = True
''    mmSound.Command = "Open"
'
'    Timer1.Enabled = True
'
'    optDiv(1).value = True
'    optDiv(0).Visible = False: optDiv(1).Visible = False
''    Query
'    '길병원은 적용후일정 시점이 지나면 아래 쿼리를 사용할예정임
'    If optDiv(0).value Then
'        Query
'    Else
'        Query_New
'    End If
''    Query
'
'    If blnFormLoad = False Then blnFormLoad = True
'
'    Call medAlwaysOn(frmBBS110, 1)
'
'    TrayI.cbSize = Len(TrayI)
'    TrayI.hwnd = pichook.hwnd 'Link the trayicon to this picturebox
'    TrayI.uId = 1&
'    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'    TrayI.ucallbackMessage = WM_LBUTTONDOWN
'    TrayI.hIcon = ImgList.ListImages(1).Picture
'    TrayI.szTip = "수혈요청 리스트" & Chr$(0)
'    'Create the icon
'    Shell_NotifyIcon NIM_ADD, TrayI
'
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'
''    RaiseEvent ThisFormUnload
''    If IsLastForm Then RaiseEvent LastFormUnload
'
'    If medMain.mnuBldRequest.Checked And Not gblnEndSystem Then
'        Cancel = True
'        Me.Hide
'        Exit Sub
'    End If
'
'    mmSound.Command = "Stop"
'    mmSound.Command = "Close"
'    TrayI.cbSize = Len(TrayI)
'    TrayI.hwnd = pichook.hwnd
'    TrayI.uId = 1&
'    'Delete the icon
'    Shell_NotifyIcon NIM_DELETE, TrayI
'    Set frmBBS110 = Nothing
'
'End Sub
'
'
'Private Sub imgSound_Click(Index As Integer)
'
'    imgSound(Index).Visible = False
'    imgSound((Index + 1) Mod 2).Visible = True
'    blnSound = Choose(Index + 1, False, True)
'    If Not blnSound Then mmSound.Command = "Stop"
'
'End Sub
'
'Private Sub mmSound_Done(NotifyCode As Integer)
'
'    If Not blnStopFg Then
'        mmSound.Command = "Stop"
'        mmSound.Command = "Close"
'        mmSound.FileName = gBloodRequestMusic
'        mmSound.Enabled = True
'        mmSound.Command = "Open"
'        If blnSound Then mmSound.Command = "Play"
'        mmSound.StopVisible = True
'        mmSound.StopEnabled = True
'    End If
'
'End Sub
'
'Private Sub mmSound_PauseClick(Cancel As Integer)
'    mmSound.Command = "Stop"
'    'mmSound.StopVisible = False
'    mmSound.StopEnabled = False
'
'    blnStopFg = Not blnStopFg
'    Timer1.Enabled = Not Timer1.Enabled
'    If Timer1.Enabled Then
'        Me.Caption = "수혈요청 리스트..(RUN)"
'    Else
'        Me.Caption = "수혈요청 리스트..(STOP)"
'    End If
'End Sub
'
'Private Sub mmSound_StopClick(Cancel As Integer)
'    blnStopFg = True
''    mmSound.StopVisible = False
'    mmSound.StopEnabled = False
'End Sub
'
'Private Sub tblPtList_DblClick(ByVal Col As Long, ByVal Row As Long)
'    If Row = 0 Then Exit Sub
'    If Col = 1 Then Exit Sub
'    frmBBS201.Show
'    tblPtList.Row = Row
'    tblPtList.Col = 8: frmBBS201.txtSpcNO.Text = tblPtList.value
'    tblPtList.Col = 3: frmBBS201.lblPtNm.Caption = tblPtList.value
'    tblPtList.Col = 2: frmBBS201.lblPtId.Caption = tblPtList.value
'    frmBBS201.ClickQueryButton
'    DoEvents
'    Me.WindowState = 1
'    If frmBBS201.txtBldNo.Enabled Then frmBBS201.txtBldNo.SetFocus
'   ' frmBBS102.ClickQueryButton
'End Sub
'Private Function GetTestInformation(ByVal sPtid As String) As String
'    Dim objSql As New clsCrossMatching
'    Dim RS     As Recordset
'    Dim strTmp As String
'    Dim SSQL   As String
'    Dim ii     As Integer
'
'    SSQL = objSql.TestResultXM(sPtid)
'    If SSQL <> "" Then
'    Set RS = New Recordset
'    RS.Open SSQL, DBConn
'        If Not RS.EOF Then
'             Do Until RS.EOF
'                 strTmp = strTmp & RS.Fields("workarea").value & "" & "-" & _
'                          RS.Fields("accdt").value & "" & "-" & _
'                          RS.Fields("accseq").value & "" & _
'                          "    " & RS.Fields("abbrnm10").value & "" & " : " & _
'                          RS.Fields("rstcd").value & "" & vbNewLine & "       "
'                RS.MoveNext
'            Loop
'        End If
'        Set RS = Nothing
'    End If
'
'    If strTmp <> "" Then
'        strTmp = "  ★ 관련검사 ★ " & vbNewLine & "       " & strTmp
'        GetTestInformation = strTmp
'    End If
'
'    Set objSql = Nothing
'End Function
'
'Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
'    Dim strtip  As String
'    Dim strPtid As String
'    Dim strTmp  As String
'    Dim sICSStr As String
'
'    If Row = 0 Then Exit Sub
'    With tblPtList
'        Call .SetTextTipAppearance("굴림체", 9, False, False, &HFFFFC0, vbBlack)
'        .Row = Row
'        .Col = 2: strPtid = Trim(.value)
'
'        sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
'
'        .Col = 3: strtip = vbCrLf & " [ " & .value & sICSStr & " ]  "
'
'        .Col = 4: strtip = strtip & .value & vbCrLf & vbCrLf
'        .Col = 5: strtip = strtip & " 제재 : " & .value & vbCrLf
'        .Col = 7: strtip = strtip & " 검체번호 : " & .value & vbCrLf
'        .Col = 8: strtip = strtip & " 접수번호 : " & .value & vbCrLf
'        .Col = 9: strtip = strtip & " 검체위치 : " & .value & vbCrLf
'        .Col = 14: strtip = strtip & " 요청일자 : " & .value & vbCrLf
'
'    End With
'    strTmp = GetTestInformation(strPtid)
'    If strTmp <> "" Then strtip = strtip & vbNewLine & strTmp
'
'    TipWidth = 5000
'    MultiLine = 1
'    TipText = strtip
'    ShowTip = True
'
'End Sub
'
'Private Sub Timer1_Timer()
'
'    Static TimeCount As Long
'    Static ImgCount As Integer
'
'    'Static mPic As Integer
'    Me.Icon = ImgList.ListImages(TimeCount Mod 3 + 1).Picture
'    TrayI.hIcon = ImgList.ListImages(TimeCount Mod 3 + 1).Picture
'    'Me.Icon = imgIcon(mPic).Picture
'    'TrayI.hIcon = imgIcon(mPic).Picture
'    'mPic = mPic + 1
'    'If mPic = 3 Then mPic = 0
'    Shell_NotifyIcon NIM_MODIFY, TrayI
'
'    TimeCount = TimeCount + 1
'
'    If Timer1.Enabled Then
'        Me.Caption = "수혈요청 리스트..(RUN)..." & TimeCount
'    Else
'        Me.Caption = "수혈요청 리스트..(STOP)"
'    End If
'
'    TrayI.szTip = Me.Caption & Chr$(0)
'
'    If TimeCount = 300 Then
'        blnStopFg = True
'        mmSound.Command = "Stop"
'        mmSound.Command = "Close"
'        DoEvents
'        '길병원은 적용후일정 시점이 지나면 아래 쿼리를 사용할예정임
'        Query_New
''        Call Query
'
''        If blnFormLoad = True Then SetCancel
'
'        TimeCount = 0
'    End If
'
'End Sub
'Private Sub SetCancel()
'    Dim objSql    As clsQueryOrder
'    Dim pWorkArea As String
'    Dim pAccDt    As String
'    Dim pAccSeq   As String
'    Dim pReqDt    As String
'    Dim pReqTm    As String
'    Dim SSQL      As String
'    Dim Cancelfg  As String
'    Dim ii        As Integer
'
'On Error GoTo DbUpdate_ERROR
'    DBConn.BeginTrans
'    Set objSql = New clsQueryOrder
'    For ii = 1 To tblPtList.DataRowCnt
'        tblPtList.Row = ii
'        tblPtList.Col = 1
'        If tblPtList.value = 1 Then
'            If chkCancel.value = 1 Then
'                Cancelfg = ""
'            Else
'                Cancelfg = "1"
'            End If
'
'
'            tblPtList.Col = 1:  pWorkArea = "B"
'            tblPtList.Col = 8:  pAccDt = medGetP(tblPtList.value, 1, "-")
'                                pAccSeq = medGetP(tblPtList.value, 2, "-")
'            tblPtList.Col = 14: pReqDt = Replace(medGetP(tblPtList.value, 1, " "), "/", "")
'                                pReqTm = Replace(medGetP(tblPtList.value, 2, " "), ":", "")
'            SSQL = objSql.RequsetCancel(pWorkArea, pAccDt, pAccSeq, pReqDt, pReqTm, Cancelfg)
'            DBConn.Execute SSQL
'        End If
'    Next
'
'    DBConn.CommitTrans
'    Set objSql = Nothing
'    Exit Sub
'
'DbUpdate_ERROR:
'    DBConn.RollbackTrans
'    Set objSql = Nothing
'End Sub
'Private Function SetRequestDelete()
'    Dim objSql As clsQueryOrder
'    Dim SSQL   As String
'    Dim sReqDt As String
'
'    On Error GoTo Delete_Error
'    Set objSql = New clsQueryOrder
'    DBConn.BeginTrans
'    sReqDt = Format(DateAdd("d", -7, GetSystemDate), "yyyymmdd")
'
'    SSQL = objSql.RequestDelete(sReqDt)
'
'    DBConn.Execute SSQL
'
'    DBConn.CommitTrans
'    Set objSql = Nothing
'    Exit Function
'
'Delete_Error:
'    DBConn.RollbackTrans
'    Set objSql = Nothing
'    MsgBox Err.Description, vbExclamation
'End Function
'Private Sub Query()
'    Dim i           As Long
'    Dim j           As Long
'
'    Dim DrRS        As Recordset
'    Dim RsTime      As Recordset
'    Dim QueryOrder  As clsQueryOrder
'    Dim ObjABO      As clsABO
'
'    Dim accno       As String
'    Dim reason      As String
'    Dim status      As String
'    Dim spcno       As String
'    Dim storeleg    As String
'    Dim storerow    As String
'    Dim storecol    As String
'    Dim center      As String
'
'    Dim strLeg      As String
'    Dim strRow      As String
'    Dim strCol      As String
'    Dim inout       As String
'    Dim MaxRowCnt   As Long
'    Dim TestDiv     As String
'    Dim blnComplete As Boolean
'
'    Dim objPrgBar   As clsProgress
'
'    Dim otherCenter As Boolean
'
'    '윗줄과 같은내용이면 글자를 감추기 위한변수들
'    Dim bkPtId      As String
'    Dim bkReason    As String
'    Dim bkReqDt     As String
'    Dim bkOrdDt     As String
'    Dim bkRoomid    As String
'    Dim bkWard      As String
'    Dim bkDept      As String
'
'    Dim strDc       As String
'    Dim PreCnt      As Long
'
'    PreCnt = tblPtList.DataRowCnt
'
'   '
'
'    Set QueryOrder = New clsQueryOrder
'
'
'
'    '2001-11-15 수정 : 요청상태의 데이타만 조회...
'    QueryOrder.stscd = "'" & BBSOrdStatus.stsREQUEST & "'"    ' "'3'"
'
'    inout = ""
'    If chkCancel.value = 1 Then QueryOrder.Cancelfg = "1"
'
'    If blnFormLoad = True Then SetCancel
'    If blnFormLoad = False Then Call SetRequestDelete
'
'    Set DrRS = QueryOrder.QueryRequest(Format(DateAdd("d", -7, Now), PRESENTDATE_FORMAT), Format(Now, PRESENTDATE_FORMAT), chkStat.value, "", inout, "", "", "")
'
'    If DrRS Is Nothing Then
'        Set DrRS = Nothing
'        Set QueryOrder = Nothing
'        Exit Sub
'    End If
'
'
''    If DrRS.RecordCount > PreCnt Then
''    'If tblPtList.MaxRows > 0 Then
''        'Me.Show
''        Me.WindowState = 0
''        blnStopFg = False
''        mmSound.Command = "Stop"
''        mmSound.Command = "Close"
''        mmSound.FileName = gBloodRequestMusic
''        mmSound.Enabled = True
''        mmSound.Command = "Open"
''        mmSound.StopVisible = True
''        mmSound.StopEnabled = True
''        mmSound.Command = "Prev"
''        If blnSound Then mmSound.Command = "Play"
''    End If
'
'
'    Set ObjABO = New clsABO
'
'    Set objPrgBar = New clsProgress
''    Set objPrgBar.StatusBar = medMain.stsBar
'    objPrgBar.Container = MainFrm.stsBar
'
'    objPrgBar.Min = 1
'    objPrgBar.Max = DrRS.RecordCount
'
'    tblPtList.MaxRows = 0
'    With tblPtList
'        bkPtId = ""
'        .ReDraw = False
'        For i = 1 To DrRS.RecordCount
'
'            objPrgBar.value = i
'
'            '-------------------------------------------
'            '처방이 irradiation 처방이 아닌 처방일경우만
'            '-------------------------------------------
'            Call QueryOrder.GetSpcNoAndStore(DrRS.Fields("ptid").value & "", spcno, storeleg, storerow, storecol, center)
'            If medGetP(center, 1, vbTab) <> ObjSysInfo.BuildingCd Then GoTo Skip
'
'            blnComplete = CompleteOrderChk(DrRS.Fields("accdt").value & "", DrRS.Fields("accseq").value & "", DrRS.Fields("unitqty").value & "")
'            If blnComplete Then GoTo Skip
'
'            MaxRowCnt = MaxRowCnt + 1
'            .MaxRows = MaxRowCnt
'            .Row = MaxRowCnt
'
'            accno = Trim(DrRS.Fields("accdt").value & "") & "-" & Val(Trim(DrRS.Fields("accseq").value & ""))
'            If accno = "-0" Then accno = "" 'accno = "미접수"
'
'            '수혈사유 구하기...
''            reason = QueryOrder.GetransReason(DrRS.Fields("ptid").value, DrRS.Fields("orddt").value, DrRS.Fields("ordno").value)
''            If reason = "" Then reason = "(없음)"
'
'            .Col = 8:   .value = accno
'            .Col = 1:
'                If chkCancel.value = 1 Then
'                    .value = IIf(DrRS.Fields("cancelfg").value & "" = "1", 0, 1)
'                Else
'                    .value = IIf(DrRS.Fields("cancelfg").value & "" = "1", 1, 0)
'                End If
'
'            .Col = 2:   .value = DrRS.Fields("ptid").value & ""
'
'            .Col = 3:   .value = DrRS.Fields("ptnm").value & ""
'            .Col = 5:   .value = DrRS.Fields("testnm").value & ""
'            .Col = 6:   .value = DrRS.Fields("reqcnt").value & ""
'
'            .Col = 10:  .value = DrRS.Fields("wardid").value & ""
'                        If Trim(DrRS.Fields("hosilid").value & "") <> "" Then
'                            .value = .value & "-" & DrRS.Fields("hosilid").value & ""
'                        End If
'            .Col = 11:  .value = DrRS.Fields("deptcd").value & ""
'
'
'            .Col = 12:  .value = Trim(Trim0(reason))
'            .Col = 13:  .value = IIf(DrRS.Fields("statfg").value = "1", "Y", "")
'                        .ForeColor = vbRed
'                        .FontBold = True
'            .Col = 14:  .value = Format("" & DrRS.Fields("ReqDt").value, CS_DateLongMask) & " " & _
'                                 Format("" & DrRS.Fields("ReqTm").value, CS_TimeLongMask)
'
'            '혈액형을 구한다.
''            ObjABO.Ptid = DrRS.Fields("ptid").value & ""
''            If ObjABO.GetABO = False Then
''                .Col = 4:    .value = ""
''            Else
''                .Col = 4:    .value = ObjABO.ABO & ObjABO.Rh
''            End If
'
'            '--------------------------
'            '검체번호와 보관장소 구하기
'            '--------------------------
'            If storerow = "0" Then storerow = ""
'            If storecol = "0" Then storecol = ""
'
'            .Col = 9:   .value = storeleg & ";" & storerow & ";" & storecol
'
'            .Col = 7:   .value = spcno
'
'            If spcno = "" Then
'                .Col = 7:   .value = "" '.value = "미채혈"
'            Else
'                If storeleg = "" Then
'                    .Col = 9:    .value = ""
'                Else
'                    .Col = 9:    .value = storeleg & "(" & storerow & "," & storecol & ")"
'                End If
'            End If
'            If DrRS.Fields("stat").value & "" = "1" Then
'                .Row = .Row: .Row2 = .Row
'                .Col = 1: .COL2 = .MaxCols
'                .BlockMode = True: .ForeColor = vbBlue: .FontBold = True
'                .BlockMode = False
'            End If
'
'Skip:
'            DrRS.MoveNext
'        Next i
'        Set objPrgBar = Nothing
'        If .DataRowCnt > 0 Then GetBatchABO
'        .ReDraw = True
'    End With
'
'    Set DrRS = Nothing
'    Set ObjABO = Nothing
'    Set objPrgBar = Nothing
'    Set QueryOrder = Nothing
'
'    If tblPtList.DataRowCnt <> PreCnt Then
'    'If tblPtList.MaxRows > 0 Then
'        'Me.Show
'        Me.WindowState = 0
'
'        blnStopFg = False
'
'        mmSound.Command = "Stop"
'        mmSound.Command = "Close"
'        mmSound.FileName = gBloodRequestMusic
'        mmSound.Enabled = True
'        mmSound.Command = "Open"
'        mmSound.StopVisible = True
'        mmSound.StopEnabled = True
'        mmSound.Command = "Prev"
'        If blnSound Then mmSound.Command = "Play"
'    End If
'
'    'lblPtCnt.Caption = "현재 " & tblPtList.MaxRows & " 명"
'
'End Sub
'
'Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim Msg As Long
'    Msg = X / Screen.TwipsPerPixelX
'    If Msg = WM_LBUTTONDBLCLK Then  'If the user dubbel-clicked on the icon
'        Me.Show
'        Me.WindowState = 0
'    'ElseIf Msg = WM_RBUTTONUP Then  'Right click
'    '    Me.PopupMenu mnuPopup
'    End If
'End Sub
'
'
'Private Sub Query_New()
'    Dim i           As Long
'
'    Dim DrRS        As Recordset
'    Dim QueryOrder  As clsQueryOrder
'
'    Dim accno       As String
'    Dim spcno       As String
'    Dim storeleg    As String
'    Dim storerow    As String
'    Dim storecol    As String
'    Dim MaxRowCnt   As Long
'    Dim blnComplete As Boolean
'
'    Dim objPrgBar   As clsProgress
'
'
'    '윗줄과 같은내용이면 글자를 감추기 위한변수들
'    Dim bkPtId      As String
'    Dim PreCnt      As Long
'
'    PreCnt = tblPtList.DataRowCnt
'    Set QueryOrder = New clsQueryOrder
'    '2001-11-15 수정 : 요청상태의 데이타만 조회...
'
'    QueryOrder.stscd = "'" & BBSOrdStatus.stsREQUEST & "'"    ' "'3'"
'
'    If blnFormLoad = True Then SetCancel
'    If blnFormLoad = False Then Call SetRequestDelete
'
'    Set DrRS = QueryOrder.TransRequest_Contents(Format(DateAdd("d", -3, Now), PRESENTDATE_FORMAT), _
'                                                Format(Now, PRESENTDATE_FORMAT), IIf(chkCancel.value = 1, "1", ""), ObjSysInfo.BuildingCd)
'
'    If DrRS Is Nothing Then
'        Set DrRS = Nothing
'        Set QueryOrder = Nothing
'        Exit Sub
'    End If
'
'
'    Set objPrgBar = New clsProgress
''    Set objPrgBar.StatusBar = medMain.stsBar
'    objPrgBar.Container = MainFrm.stsBar
'
'    objPrgBar.Min = 1
'    objPrgBar.Max = DrRS.RecordCount
'
'    tblPtList.MaxRows = 0
'    With tblPtList
'        bkPtId = ""
'        .ReDraw = False
'        For i = 1 To DrRS.RecordCount
'
'            objPrgBar.value = i
'
'            If Val(DrRS.Fields("assigncnt").value & "") - Val(DrRS.Fields("assigncancelcnt").value & "") = Val(DrRS.Fields("unitqty").value & "") Then GoTo Skip
'            MaxRowCnt = MaxRowCnt + 1
'            .MaxRows = MaxRowCnt
'            .Row = MaxRowCnt
'
'            accno = Trim(DrRS.Fields("accdt").value & "") & "-" & Val(Trim(DrRS.Fields("accseq").value & ""))
'            If accno = "-0" Then accno = "" 'accno = "미접수"
'
'
'            .Col = 8:   .value = accno
'            .Col = 1:
'                If chkCancel.value = 1 Then
'                    .value = IIf(DrRS.Fields("cancelfg").value & "" = "1", 0, 1)
'                Else
'                    .value = IIf(DrRS.Fields("cancelfg").value & "" = "1", 1, 0)
'                End If
'
'            .Col = 2:   .value = DrRS.Fields("ptid").value & ""
'
'            .Col = 3:   .value = GetPtNm(DrRS.Fields("ptid").value & "")
'            .Col = 5:   .value = DrRS.Fields("testnm").value & ""
'            .Col = 6:   .value = DrRS.Fields("reqcnt").value & ""
'
'            .Col = 10:  .value = DrRS.Fields("wardid").value & ""
'                        If Trim(DrRS.Fields("hosilid").value & "") <> "" Then
'                            .value = .value & "-" & DrRS.Fields("hosilid").value & ""
'                        End If
'            .Col = 11:  .value = DrRS.Fields("deptcd").value & ""
'
'            .Col = 13:  .value = IIf(DrRS.Fields("statfg").value = "1", "Y", "")
'                        .ForeColor = vbRed
'                        .FontBold = True
'            .Col = 14:  .value = Format("" & DrRS.Fields("ReqDt").value, CS_DateLongMask) & " " & _
'                                 Format("" & DrRS.Fields("ReqTm").value, CS_TimeLongMask)
'
'
'            '--------------------------
'            '검체번호와 보관장소 구하기
'            '--------------------------
'            storeleg = DrRS.Fields("storeleg").value & ""
'            storerow = DrRS.Fields("storerno").value & ""
'            storecol = DrRS.Fields("storecno").value & ""
'
'            If storerow = "0" Then storerow = ""
'            If storecol = "0" Then storecol = ""
'
'            .Col = 9:   .value = storeleg & ";" & storerow & ";" & storecol
'
'            .Col = 7:   .value = DrRS.Fields("spcyy").value & "" & "-" & DrRS.Fields("spcno").value & ""
'
'            If .value = "" Then
'                .Col = 7:   .value = "" '.value = "미채혈"
'
'            Else
'                If storeleg = "" Then
'                    .Col = 9:    .value = ""
'                Else
'                    .Col = 9:    .value = storeleg & "(" & storerow & "," & storecol & ")"
'                End If
'            End If
'
'            If DrRS.Fields("stat").value & "" = "1" Then
'                .Row = .Row: .Row2 = .Row
'                .Col = 1: .COL2 = .MaxCols
'                .BlockMode = True: .ForeColor = vbBlue: .FontBold = True
'                .BlockMode = False
'            End If
'
'Skip:
'            DrRS.MoveNext
'        Next i
'        Set objPrgBar = Nothing
'        If .DataRowCnt > 0 Then GetBatchABO
'    End With
'
'    Set DrRS = Nothing
'    Set QueryOrder = Nothing
'
'    If tblPtList.DataRowCnt <> PreCnt Then
'        Me.WindowState = 0
'        blnStopFg = False
'        mmSound.Command = "Stop"
'        mmSound.Command = "Close"
'        mmSound.FileName = gBloodRequestMusic
'        mmSound.Enabled = True
'        mmSound.Command = "Open"
'        mmSound.StopVisible = True
'        mmSound.StopEnabled = True
'        mmSound.Command = "Prev"
'        If blnSound Then mmSound.Command = "Play"
'    End If
'End Sub
'
'Private Sub GetBatchABO()
'    Dim ObjABO      As New clsABO
''    Dim objPrgBar   As New clsprogress
'    Dim ii          As Integer
'    Dim tmpptid     As String
'    Dim sPtid       As String
'    Dim sORDDT      As String
'    Dim objPrgBar   As clsProgress
'
'    Set objPrgBar = Nothing
'    Set objPrgBar = New clsProgress
''    Set objPrgBar.StatusBar = medMain.stsBar
'    objPrgBar.Container = MainFrm.stsBar
'
'    With tblPtList
'        objPrgBar.Max = .DataRowCnt
'        .ReDraw = False
'        For ii = 1 To .DataRowCnt
'            .Row = ii
'            .Col = 2
'            If tmpptid <> Trim(.value) Then
'                ObjABO.PtId = .value
'
'                If ObjABO.GetABO = False Then
'                    .Col = 4:     .value = ""
'                Else
'                    .Col = 4:     .value = ObjABO.ABO & ObjABO.Rh
'                End If
'            Else
'                .Col = 4: .value = ObjABO.ABO & ObjABO.Rh
'            End If
'            .Col = 2: tmpptid = Trim(.value)
'            objPrgBar.value = ii
'            objPrgBar.Message = tmpptid & " 의 혈액형을 검색중입니다."
'        Next
'        .ReDraw = True
'    End With
'
'    Set ObjABO = Nothing
'    Set objPrgBar = Nothing
'End Sub
'
'Private Function CompleteOrderChk(ByVal accdt As String, ByVal accseq As String, ByVal unitqty As Long) As Boolean
'    Dim objXM As New clsCrossMatching
'    Dim A_Cnt As Long   'Assign수량
'    Dim C_Cnt As Long   'Assign Cancel 수량
'    Dim O_Cnt As Long   '출고수량
'    Dim R_Cnt As Long   '반환수량
'    Dim X_Cnt As Long   '폐기수량
'    Dim T_Cnt As Long   '총Assign 수량
'
'
'    'CompleteOrderChk=True이면 완결처방
'    'CompleteOrderChk=미완결처방
'    CompleteOrderChk = False
'    If accdt <> "" Then
'
'        With objXM
'            .Assign_Cnt accdt, Val(accseq)
'            A_Cnt = .AssignCnt
'            C_Cnt = .CancelCnt
'            O_Cnt = .OutCnt
'            R_Cnt = .RetCnt
'            X_Cnt = .ExpCnt
'        End With
'
'        T_Cnt = A_Cnt - C_Cnt - R_Cnt - X_Cnt
'
'        If unitqty = T_Cnt Then
'            CompleteOrderChk = True
'        End If
'    End If
'    Set objXM = Nothing
'
'End Function
'
