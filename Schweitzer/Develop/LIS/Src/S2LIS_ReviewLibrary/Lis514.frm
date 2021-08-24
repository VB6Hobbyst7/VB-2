VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm4NewResultView 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "최근결과리스트"
   ClientHeight    =   9195
   ClientLeft      =   1005
   ClientTop       =   5865
   ClientWidth     =   17460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   17460
   ShowInTaskbar   =   0   'False
   Tag             =   "Abnormal Result"
   WindowState     =   2  '최대화
   Begin VB.CheckBox chkDate 
      BackColor       =   &H00DBE6E6&
      Caption         =   "일자별"
      Height          =   315
      Left            =   12330
      TabIndex        =   14
      Top             =   150
      Width           =   1905
   End
   Begin VB.ComboBox cboWorkarea 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1520
      Style           =   2  '드롭다운 목록
      TabIndex        =   10
      Top             =   120
      Width           =   2250
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   4838
      Top             =   8565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   345
      Left            =   4860
      TabIndex        =   6
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   84213761
      CurrentDate     =   36483
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "To &Excel"
      Height          =   510
      Left            =   13365
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "127"
      Top             =   8490
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Print"
      Height          =   510
      Left            =   14670
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "132"
      Top             =   8490
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   16005
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8490
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FCEFE9&
      Caption         =   "&Start Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15480
      MaskColor       =   &H00D4D4D4&
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "158"
      Top             =   80
      Width           =   1860
   End
   Begin FPSpread.vaSpread ssAbnormal 
      Height          =   7800
      Left            =   120
      TabIndex        =   1
      Tag             =   "45410"
      Top             =   600
      Width           =   17205
      _Version        =   196608
      _ExtentX        =   30348
      _ExtentY        =   13758
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16252413
      MaxCols         =   12
      MaxRows         =   5
      OperationMode   =   1
      ProcessTab      =   -1  'True
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis514.frx":0000
      UserResize      =   1
      VisibleCols     =   7
   End
   Begin MSComCtl2.DTPicker dtpEndDt 
      Height          =   345
      Left            =   8880
      TabIndex        =   7
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   84213761
      CurrentDate     =   36483
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   345
      Index           =   0
      Left            =   3960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BackColor       =   14411494
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "From"
      Appearance      =   0
      LeftGab         =   0
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   915
      Left            =   5640
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
      _ExtentY        =   1614
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Lis514.frx":0A62
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   1295
      _ExtentX        =   2275
      _ExtentY        =   609
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "Workarea"
      Appearance      =   0
      LeftGab         =   0
   End
   Begin MSComCtl2.DTPicker dtpStartTm 
      Height          =   345
      Left            =   6720
      TabIndex        =   12
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   84213763
      CurrentDate     =   37770
   End
   Begin MSComCtl2.DTPicker dtpEndTm 
      Height          =   345
      Left            =   10750
      TabIndex        =   13
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   84213763
      CurrentDate     =   37770
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DBE6E6&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8520
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frm4NewResultView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Click()

Private WithEvents fL401 As frm401ResultView
Attribute fL401.VB_VarHelpID = -1


Public Event LastFormUnload()

Private Sub cmdExcel_Click()

    Dim strTmp  As String
    
    If ssAbnormal.DataRowCnt = 0 Then Exit Sub
    
    With ssAbnormal
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblExcel.MaxRows = .MaxRows + 1
        tblExcel.MaxCols = .MaxCols
        tblExcel.Row = 1: tblExcel.Row2 = tblExcel.MaxRows
        tblExcel.Col = 1: tblExcel.Col2 = tblExcel.MaxCols
        tblExcel.BlockMode = True
        tblExcel.Clip = strTmp
        tblExcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "AbnormalList"
    DlgSave.ShowSave

    tblExcel.SaveTabFile (DlgSave.FileName)


End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub


Private Sub cmdStart_Click()
    Dim ChoiceCondFlag As Boolean
    Dim I%
    
    ChoiceCondFlag = False
    
    If dtpStartDt.Value > dtpEndDt.Value Then
        MsgBox "Duration input Error"
        Exit Sub
    End If
    
'    For I = 0 To 3
'        If chkCondition1(I).Value = 1 Then
'            ChoiceCondFlag = True
'        End If
'    Next I
    
'    If ChoiceCondFlag = True Then
        Call StartQuery
'    Else
'        MsgBox " 검색조건(Low,High,Panic,Delta)를 설정하세요.", vbInformation, "검색조건"
'        Exit Sub
'    End If
End Sub

Private Sub StartQuery()
    Dim objsSQL     As clsLISSqlStatistic
    Dim objProBar   As jProgressBar.clsProgress
    Dim rsGetinfo   As Recordset
    Dim sSqlGetinfo As String
    Dim sStartDt    As String
    Dim SendDt      As String
    Dim sStartTm    As String
    Dim SendTm      As String
    Dim strWorkArea As String
    Dim I           As Double
    Dim strRstVal   As String
    Dim strKiSun    As String
    Dim strTemp     As String
    
    sStartDt = Format(dtpStartDt.Value, CS_DateDbFormat)
    SendDt = Format(dtpEndDt.Value, CS_DateDbFormat)
    sStartTm = Format(dtpStartTm.Value, CS_TimeDbFormat)
    SendTm = Format(dtpEndTm.Value, CS_TimeDbFormat)
    
    '## 5.0.2: 이상대(2004-12-29)
    '   - 검색조건에 Workarea 추가
    Set objsSQL = New clsLISSqlStatistic
    If chkDate.Value = 0 Then
        If cboWorkArea.ListIndex = 0 Then
            sSqlGetinfo = objsSQL.GetNewResultLst(sStartDt, SendDt, sStartTm, SendTm)
        Else
            strWorkArea = Trim$(medGetP(cboWorkArea.Text, 2, COL_DIV))
            sSqlGetinfo = objsSQL.GetNewResultLst(sStartDt, SendDt, sStartTm, SendTm, strWorkArea)
        End If
    Else
        If cboWorkArea.ListIndex = 0 Then
            sSqlGetinfo = objsSQL.GetNewResultLst_New(sStartDt, SendDt, sStartTm, SendTm)
        Else
            strWorkArea = Trim$(medGetP(cboWorkArea.Text, 2, COL_DIV))
            sSqlGetinfo = objsSQL.GetNewResultLst_New(sStartDt, SendDt, sStartTm, SendTm, strWorkArea)
        End If
    End If
    
    Set objProBar = New jProgressBar.clsProgress
    With objProBar
        .Container = Me
        .Width = ssAbnormal.Width
        .Left = ssAbnormal.Left
        .Top = ssAbnormal.Top - 280
        .Height = 280
        .Message = "자료를 읽기 위해 준비중입니다..."
    End With
    
    Set rsGetinfo = New Recordset
    rsGetinfo.Open sSqlGetinfo, DBConn
    
    objProBar.DisplayMessage = False
    
    If rsGetinfo.RecordCount > 0 Then
        objProBar.Max = rsGetinfo.RecordCount
    Else
        MsgBox "데이타가 없습니다..", vbExclamation
    End If
    
    Call ClearSSAbnormal
    With ssAbnormal
        
        .MaxRows = 1
        For I = 1 To rsGetinfo.RecordCount
            objProBar.Value = I
            DoEvents

'        .MaxRows = rsGetinfo.RecordCount
        If strKiSun = "" & rsGetinfo.Fields("workarea").Value & "-" & rsGetinfo.Fields("accdt").Value & "-" & rsGetinfo.Fields("accseq").Value Then
            .Col = 2: .Value = .Value & "," & rsGetinfo.Fields("testnm").Value
        Else
            strKiSun = "" & rsGetinfo.Fields("workarea").Value & "-" & rsGetinfo.Fields("accdt").Value & "-" & rsGetinfo.Fields("accseq").Value
            If .DataRowCnt + 1 > .MaxRows Then
                .MaxRows = .MaxRows + 1
            End If
            .Row = .DataRowCnt + 1
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
            .Col = 1: .Value = strKiSun
            .Col = 2: .Value = "" & rsGetinfo.Fields("testnm").Value
            .Col = 3: .Value = GetDeptNm("" & rsGetinfo.Fields("deptcd").Value) & "-" & rsGetinfo.Fields("wardid").Value & "-" & rsGetinfo.Fields("roomid").Value ' 과-병동/호실
            .Col = 4: .Value = "" & rsGetinfo.Fields("ptid").Value
            .Col = 5: .Value = "" & rsGetinfo.Fields("ptnm").Value
            .Col = 6: .Value = Trim("" & rsGetinfo.Fields("spcyy").Value) & Format$(Trim("" & rsGetinfo.Fields("spcno").Value), "000000000")   ' 검체번호
            .Col = 7: .Value = "" & rsGetinfo.Fields("empnm").Value
            .Col = 9: .Value = Format$(Mid$(Trim("" & rsGetinfo.Fields("coldt").Value), 3), "0#-##-##") & " " & Format$(Mid$(Trim("" & rsGetinfo.Fields("coltm").Value), 1, 4), "0#:0#")          ' 보고일시
            '.Col = 12: .Value = "" & rsGetinfo.Fields("rsttxt").Value    ' 코멘트
            .Col = 12: .Value = "" & rsGetinfo.Fields("spcnm").Value    ' 검체명으로 변경
            .Col = 8: .Value = Format$(Mid$(Trim("" & rsGetinfo.Fields("vdt").Value), 3), "0#-##-##") & " " & Format$(Mid$(Trim("" & rsGetinfo.Fields("vtm").Value), 1, 4), "0#:0#")   ' 채혈일시
            .Col = 10: .Value = GetEmpNm("" & rsGetinfo.Fields("colid").Value)    ' 채혈자
            .Col = 11: .Value = Format$(Mid$(Trim("" & rsGetinfo.Fields("rcvdt").Value), 3), "0#-##-##") & " " & Format$(Mid$(Trim("" & rsGetinfo.Fields("rcvtm").Value), 1, 4), "0#:0#")    ' 접수일시
        End If

          rsGetinfo.MoveNext
        Next I
    End With
    
    Set rsGetinfo = Nothing
    Set objsSQL = Nothing
    Set objProBar = Nothing
End Sub
Private Sub ssAbnormal_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strMask As String
    Dim strPtId     As String
    Static iSortOrder As Integer
    
    With ssAbnormal
        If Row = 0 Then
            .Row = 0: .Col = Col
            .Row = -1: .Col = -1
            .SortBy = SortByRow
            .SortKey(1) = Col
            If iSortOrder = SortKeyOrderAscending Then
                .SortKeyOrder(1) = SortKeyOrderDescending
                iSortOrder = SortKeyOrderDescending
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                iSortOrder = SortKeyOrderAscending
            End If
            .Action = ActionSort
            Exit Sub
        End If
        If Row > .DataRowCnt Then Exit Sub
        
        '소요시간
        If Col = 4 Or Col = 5 Then
            .Row = Row
            .Col = 4: strPtId = .Value
            MainFrm.lblSubMenu.Caption = "처방결과조회"  ' Me.Caption

            If fL401 Is Nothing Then Set fL401 = New frm401ResultView
            Call SetParent(fL401.hwnd, Me.hwnd)
            If gUsingInWardMenu = False Then
                Me.WindowState = 2
                fL401.WindowState = 2
            End If
'            fL401.DeptCd = mvarDeptCd
'            P_ReviewStartDate = DateAdd("yyyy", -0.5, GetSystemDate)
            fL401.accPTid (strPtId)
            fL401.ZOrder 0
            DoEvents
            Exit Sub
        End If
        Call MouseDefault
    End With
End Sub

'Private Sub DspSpd(PTid As String, Sex As String, AgeDay As Long, WardId As String, _
'                   RoomId As String, BuildNm As String, TestCd As String, HLDiv As String, _
'                   DPDiv As String, LastRst As String, RstTxt As String, ptnt_nm As String, RstVal As String, _
'                   TestNm As String, ByVal pVfyNm As String, ByVal pVfyDt As String, _
'                   ByVal pVfyTm As String, RowNm As Integer)
 '   Dim sAge As String
 '   Dim Age As Integer
 '   Dim Location As String
 '   Dim tmpPtId As String
 '
 '   Age = (AgeDay / 365) + 1
 '   sAge = Sex & "/" & CStr(Age)
 '   Location = WardId & "-" & RoomId'
'
'    With ssAbnormal
'        .MaxRows = RowNm
'        .Row = RowNm - 1
'        .Col = 1
'        If .Value <> PTid Then
'            .Row = RowNm
'            .Col = 1: .Text = Trim(PTid)
'            .Col = 2: .Text = Trim(ptnt_nm)
'            .Col = 3: .Text = sAge
'            .Col = 4: .Text = Location
'            .Col = 5: .Text = Trim(BuildNm)
'        Else
'            .Row = RowNm
'            .Col = 1: .Text = Trim(PTid): .ForeColor = .BackColor
'            .Col = 2: .Text = Trim(ptnt_nm): .ForeColor = .BackColor
'            .Col = 3: .Text = sAge: .ForeColor = .BackColor
'            .Col = 4: .Text = Location: .ForeColor = .BackColor
'            .Col = 5: .Text = Trim(BuildNm): .ForeColor = .BackColor
'        End If
'        .Row = RowNm
'        .Col = 6: .Text = Trim(TestNm)
'        .Col = 7: .Text = Trim(RstVal)
'
 '       If Trim(HLDiv) = "L" Then
 '           .Col = 8: .Text = "L": .ForeColor = DCM_LightBlue
 '       ElseIf Trim(HLDiv) = "H" Then
 '           .Col = 9: .Text = "H": .ForeColor = DCM_LightRed
 '       ElseIf Trim(HLDiv) = "N" Then    ' blank 일 경우 아무것도 안한다.
 '           .Col = 8: .Text = "N"
 '       Else
 '
 '       End If
 '
 '       If Trim(DPDiv) = "P" Then
 '           .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
 '       ElseIf Trim(DPDiv) = "D" Then
 '           .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
 '       ElseIf Trim(DPDiv) = "" Then    ' blank 일 경우 아무것도 안한다.
 '
 '       Else
 '           .Col = 10: .Text = "P": .ForeColor = DCM_Red: .FontBold = True
 '           .Col = 11: .Text = "D": .ForeColor = DCM_Red: .FontBold = True
 '       End If
 '
 '       .Col = 12: .Text = DelEnterKey(LastRst)
 '       .Col = 13: .Text = pVfyNm
 '       .Col = 14: .Text = Format$(Mid$(pVfyDt, 3), "0#-##-##") & " " & Format$(Mid$(pVfyTm, 1, 4), "0#:0#")
 '       .Col = 15: .Text = DelEnterKey(RstTxt)
 '   End With
'End Sub

Private Function DelEnterKey(RstTxt As String) As String
    Dim StartPos As Long
    Dim EnterKeyPos As Long
    
    StartPos = 1
    
    Do
        EnterKeyPos = InStr(StartPos, RstTxt, Chr(13), 0)
        If EnterKeyPos = 0 Then Exit Do
        Mid(RstTxt, EnterKeyPos, 2) = "  "
        StartPos = EnterKeyPos + 2
    Loop
    
    DelEnterKey = RstTxt
End Function

Private Sub ClearSSAbnormal()
    With ssAbnormal
        .Col = -1
        .Row = -1
        .Action = ActionClearText
        .MaxRows = 0
    End With
End Sub

Private Sub dtpEndDt_Validate(Cancel As Boolean)
    ClearSSAbnormal
End Sub


Private Sub dtpStartDt_Validate(Cancel As Boolean)
    ClearSSAbnormal
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    
    Call medAlwaysOn(frm4NewResultView, 1)
    
    Screen.MousePointer = vbHourglass
    
    dtpStartDt.Value = GetSystemDate
    dtpEndDt.Value = dtpStartDt.Value
    dtpStartTm.Value = dtpStartDt.Value - 0.04167
    dtpEndTm.Value = dtpStartDt.Value
    
'    chkCondition1(2).Value = 1
'    chkCondition1(3).Value = 1
    Call ClearSSAbnormal
    Call GetWorkArea
    
    Screen.MousePointer = vbDefault
End Sub

'Private Sub AbnormalHead()
'    Dim strTmp  As String
'    Dim ii      As Integer
'
'    strTmp = "Abnormal Result"
'    Printer.DrawStyle = 0: Printer.DrawWidth = 6
'    lngCurYPos = 10
'
'    Printer.FontSize = 20: Printer.FontBold = True
'    Call Print_Setting("Abnormal Result", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
'    Printer.FontSize = 9: Printer.FontBold = False
'
'    strTmp = "조회기간 : " & Format(dtpStartDt.Value, "YYYY년 MM월 DD일") & " ~ " & Format(dtpEndDt.Value, "YYYY년 MM월 DD일")
'    Call Print_Setting(strTmp, PrtLeft, LineSpace, Printer.Width - PrtLeft, "L", "C", True)
'
'    strTmp = "조회조건 : "
'    For ii = 0 To 3
'        If chkCondition1(ii).Value = 1 Then
'            Select Case ii
'                Case 0: strTmp = strTmp & "     " & "(√)High"
'                Case 1: strTmp = strTmp & "     " & "(√)Low "
'                Case 2: strTmp = strTmp & "     " & "(√)Panic"
'                Case 3: strTmp = strTmp & "     " & "(√)Delta"
'            End Select
'        Else
'            Select Case ii
'                Case 0: strTmp = strTmp & "     " & "(  )High"
'                Case 1: strTmp = strTmp & "     " & "(  )Low "
'                Case 2: strTmp = strTmp & "     " & "(  )Panic"
'                Case 3: strTmp = strTmp & "     " & "(  )Delta"
'            End Select
'        End If
'    Next
'
'    Call Print_Setting(strTmp, PrtLeft, LineSpace, Printer.Width - PrtLeft, "L", "C", True)
'
'    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
'    Call PrintString("환자ID", "환자명", "성별/나이", "병실", "검사명", "결과", "Low", "High", "Panic", "Delta", "최근결과", "")
'
'    Printer.DrawStyle = 0: Printer.DrawWidth = 6
'    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
'End Sub
'Private Sub PrintString(ByVal sPTID As String, ByVal sPtnm As String, ByVal sSexAge As String, ByVal sLocation As String, _
'                        ByVal sTestNm As String, ByVal sResult As String, ByVal sLow As String, ByVal sHigh As String, _
'                        ByVal sPanic As String, ByVal sDelta As String, ByVal sLastResult As String, ByVal sMesg As String)
'    Dim aryTmp() As String
'    Dim ii As Integer
'
'
'    If lngCurYPos > Printer.ScaleHeight - 6 Then
'        Printer.NewPage
'        Call AbnormalHead
'    End If
'
'    Call Print_Setting(sPTID, PrtLeft, LineSpace, 20, "L", "C", False)
'    Call Print_Setting(sPtnm, 25, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sSexAge, 40, LineSpace, 20, "L", "C", False)
'    Call Print_Setting(sLocation, 60, LineSpace, 20, "L", "C", False)
'    Call Print_Setting(sTestNm, 80, LineSpace, 30, "L", "C", False)
'    Call Print_Setting(sResult, 110, LineSpace, 20, "L", "C", False)
'    Call Print_Setting(sLow, 130, LineSpace, 10, "L", "C", False)
'    Call Print_Setting(sHigh, 140, LineSpace, 10, "L", "C", False)
'    Call Print_Setting(sPanic, 150, LineSpace, 10, "L", "C", False)
'    Call Print_Setting(sDelta, 160, LineSpace, 10, "L", "C", False)
'    Call Print_Setting(sLastResult, 170, LineSpace, 20, "L", "C")
'
'    If sMesg <> "" Then
'        Printer.FontBold = True
'        For ii = 1 To 5
'            If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
'                sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
'                If Mid(sMesg, Len(sMesg) - 1, 1) = vbCr Or Mid(sMesg, Len(sMesg) - 1, 1) = vbLf Then
'                    sMesg = Mid(sMesg, 1, Len(sMesg) - 1)
'                End If
'            End If
'        Next
'        aryTmp() = Split(Trim(sMesg), vbCrLf)
'        For ii = LBound(aryTmp) To UBound(aryTmp)
'            If lngCurYPos > Printer.ScaleHeight - 6 Then
'                Printer.NewPage
'                Call AbnormalHead
'            End If
'            Call Print_Setting(aryTmp(ii), PrtLeft, LineSpace, 20, "L", "C")
'        Next
'        Printer.FontBold = False
'    End If
'
'    Printer.DrawStyle = 1: Printer.DrawWidth = 2
'    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
'
'End Sub

Private Sub cmdPrint_Click()
    Dim sPTID       As String
    Dim sPtnm       As String
    Dim sSexAge     As String
    Dim sLocation   As String
    Dim sTestNm     As String
    Dim sResult     As String
    Dim sLow        As String
    Dim sHigh       As String
    Dim sPanic      As String
    Dim sDelta      As String
    Dim sLastResult As String
    Dim sMesg       As String
    Dim ii As Integer
                        If ssAbnormal.DataRowCnt < 1 Then Exit Sub
    
'    Call P_PrtSet
'    Call AbnormalHead
'
'    With ssAbnormal
'        For ii = 1 To .DataRowCnt
'            .Row = ii
'            .Col = 1:   sPTID = .Value
'            .Col = 2:   sPtnm = .Value
'            .Col = 3:   sSexAge = .Value
'            .Col = 4:   sLocation = .Value
'            .Col = 6:   sTestNm = .Value
'            .Col = 7:   sResult = .Value
'            .Col = 8:   sLow = .Value
'            .Col = 9:   sHigh = .Value
'            .Col = 10:  sPanic = .Value
'            .Col = 11:  sDelta = .Value
'            .Col = 12:  sLastResult = .Value
'            .Col = 15:  sMesg = .Value
'            Call PrintString(sPTID, sPtnm, sSexAge, sLocation, _
'                             sTestNm, sResult, sLow, sHigh, _
'                             sPanic, sDelta, sLastResult, sMesg)
'        Next
'    End With
'
'    Printer.EndDoc
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Workarea 조회 - 5.0.2: 이상대(2004-12-29)
'-----------------------------------------------------------------------------'
Private Sub GetWorkArea()
    Dim objSQL  As clsLISSqlQc
    Dim RS      As Recordset

    cboWorkArea.Clear
    cboWorkArea.AddItem "전 체"
    
On Error GoTo Errors
    Set objSQL = New clsLISSqlQc
    Set RS = New Recordset
    RS.Open objSQL.GetWorkArea, DBConn
    If Not (RS.BOF Or RS.EOF) Then
        Do Until RS.EOF
            cboWorkArea.AddItem Format(RS.Fields("field1").Value & "", "!" & String(50, "@")) & COL_DIV & _
                               RS.Fields("cdval1").Value & ""
            RS.MoveNext
        Loop
    End If
    RS.Close
    Set RS = Nothing
    Set objSQL = Nothing
    
    If cboWorkArea.ListCount > 0 Then cboWorkArea.ListIndex = 0
    cboWorkArea.ListIndex = 1
    Exit Sub

Errors:
    Set RS = Nothing
    Set objSQL = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Sub

