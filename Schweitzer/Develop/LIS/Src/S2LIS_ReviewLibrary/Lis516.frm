VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm516Ab 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "경고치 시행율"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14955
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   11640
      Style           =   1  '그래픽
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   570
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H00DBE6E6&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   75
      ScaleHeight     =   930
      ScaleWidth      =   14355
      TabIndex        =   3
      Top             =   345
      Width           =   14385
      Begin VB.Frame fraDt 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회 기간"
         Height          =   705
         Left            =   90
         TabIndex        =   12
         Top             =   120
         Width           =   1815
         Begin MSComCtl2.DTPicker dtpFromDt 
            Height          =   315
            Left            =   180
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   24313856
            CurrentDate     =   36342.5951388889
         End
         Begin MSComCtl2.DTPicker dtpToDt 
            Height          =   315
            Left            =   2460
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   24313856
            CurrentDate     =   36342.5951388889
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "부터"
            Height          =   180
            Left            =   1830
            TabIndex        =   16
            Tag             =   "15104"
            Top             =   300
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "까지"
            Height          =   180
            Left            =   4110
            TabIndex        =   15
            Tag             =   "15104"
            Top             =   330
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame fraWa 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Work Area"
         Height          =   705
         Left            =   1920
         TabIndex        =   10
         Top             =   120
         Width           =   2415
         Begin VB.ComboBox cboWA 
            Height          =   300
            Left            =   120
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사항목"
         Height          =   705
         Left            =   4350
         TabIndex        =   6
         Top             =   120
         Width           =   4995
         Begin VB.CommandButton cmdHelpList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2025
            MaskColor       =   &H00F4F0F2&
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   8
            Tag             =   "DeptCd"
            Top             =   270
            Width           =   285
         End
         Begin VB.TextBox txtTestCd 
            Height          =   315
            Left            =   135
            TabIndex        =   7
            Top             =   285
            Width           =   1875
         End
         Begin MedControls1.LisLabel lblTestNm 
            Height          =   330
            Left            =   2370
            TabIndex        =   9
            Top             =   285
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            BackColor       =   15988984
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
         End
      End
      Begin VB.CommandButton cmdQuary 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조 회(&Q)"
         Height          =   510
         Left            =   10230
         Style           =   1  '그래픽
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1320
      End
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Excel(&E)"
         Height          =   510
         Left            =   12900
         Style           =   1  '그래픽
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   210
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   285
      Left            =   75
      TabIndex        =   17
      Top             =   60
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   503
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "검사항목 별 양성율"
      LeftGab         =   100
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   7110
      Left            =   60
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1350
      Width           =   14415
      _Version        =   196608
      _ExtentX        =   25426
      _ExtentY        =   12541
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   6
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
      GrayAreaBackColor=   15463405
      MaxCols         =   6
      MaxRows         =   1
      ScrollBars      =   2
      ShadowColor     =   14411494
      SpreadDesigner  =   "Lis516.frx":0000
      UserResize      =   0
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   6090
      Top             =   8550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   750
      _Version        =   196608
      _ExtentX        =   1323
      _ExtentY        =   1191
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
      SpreadDesigner  =   "Lis516.frx":04A1
   End
End
Attribute VB_Name = "frm516Ab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' 폼의 속성중 다음은 유지해야 합니다.
'
' BorderStyle : 0 - 없음
' MdiChild    : False
' WindowState : 0 - 표준
' Top         : 0
' Left        : 0
'
Public Event FormClose()
Public Event LastFormUnload()

Private Const FAddCol = 1


'리스트 팝업
Private WithEvents objListPop   As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1
Private WithEvents fL401 As S2LIS_ReviewLib.clsLisReviewForm
Attribute fL401.VB_VarHelpID = -1

Private objSQL  As New clsLISSqlStatistic
Private objIcdList  As clsDictionary
Private objRstCd    As clsDictionary

Private aryResultText() As String

Private blnCHkLoad As Boolean

Dim CaseStudy_TestCd As String


Private Sub chkIndex_Click()
    
    txtTblClear
End Sub

Private Sub chkShow_Click()
    txtTblClear
End Sub

Private Function PrintOut() As Boolean
'    Dim strTmp      As String
'    Dim strFileNm   As String
'    Dim strRptNm    As String
'    Dim strMyFile   As String
'    Dim strTemp     As String
'    Dim strOption   As String
'    Dim lngFNum     As Long
'    Dim lngCnt      As Long
'    Dim i           As Long
'    Dim j           As Long
'
'
'    strMyFile = Dir(APSAppPath & "\..\rpt\CrystalReport.txt")
'
'    If strMyFile = "" Then
'        PrintOut = True
'        MsgBox "CrystalReport.txt 파일이 없습니다.", vbCritical, "정보확인"
'        Exit Function
'    End If
'    strMyFile = ""
'
'    strFileNm = APSAppPath & "\..\rpt\CrystalReport.txt"
'
'    strMyFile = Dir(APSAppPath & "\..\rpt\rptAPS021.rpt")
'
'    If strMyFile = "" Then
'        PrintOut = True
'        MsgBox "rptAPS021.rpt 파일이 없습니다.", vbCritical, "정보확인"
'        Exit Function
'    End If
'
'    strRptNm = APSAppPath & "\..\rpt\rptAPS021.rpt"
'
'    With tblIndex
'        For i = 1 To .DataRowCnt '.MaxRows
'            .Row = i
'            For j = 1 To 8
'                .Col = j
'                strTmp = strTmp & .Value & vbTab
'                lngCnt = lngCnt + 1
'            Next
'
'            If (lngCnt Mod 8) = 0 Then
'                strTmp = strTmp & vbCr
'            End If
'        Next
'    End With
'
'    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
'
'    Debug.Print strTmp
'
'    lngFNum = FreeFile
'
'On Error GoTo ErrPrint
'
'    Open strFileNm For Output As #lngFNum
'    Print #lngFNum, strTmp
'    Close #lngFNum
'    With crtReport
'        .ReportFileName = strRptNm
'        .ParameterFields(0) = "hostnm;" & AC5_HOSPITAL_DEPT_NAME & ";true"
''        .ParameterFields(0) = "HostNm;" & objSysInfo.Hospital & ";true"
'        .RetrieveDataFiles
'        .WindowState = 2 ' crptMaximized
'        .Destination = crptToWindow
'        .Action = 1
'        .Reset
'    End With
'    PrintOut = True
'    Exit Function
'
'ErrPrint:
'    PrintOut = False
End Function

Private Sub cboWA_Click()
    Call TxtClear
    Call txtTblClear
    If cboWA.ListIndex <> -1 Then
        If cboWA.Text <> CaseStudy_TestCd Then
            CaseStudy_TestCd = cboWA.Text
            txtTestCd.Text = ""
            lblTestNm.Caption = ""
        End If
    End If
End Sub


Private Sub cmdExcel_Click()

    Dim strTmp  As String
    
    If tblResult.DataRowCnt = 0 Then Exit Sub
    
    With tblResult
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
    DlgSave.FileName = "CVR 시행율"
    DlgSave.ShowSave

    tblExcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdHelpList_Click()
    Dim objTestDiv As New clsDictionary
    Dim objRs As Recordset
    
    If cboWA.ListIndex = -1 Then Exit Sub
    
    Set objListPop = New clsPopUpList
    
    Call TxtClear
    Call txtTblClear
    
    With objTestDiv
        .Clear
        .FieldInialize "검사항목코드", "검사명,구분"
        Set objRs = New Recordset
        objRs.Open objSQL.GetWAvsTest(medGetP(cboWA.Text, 1, " ")), DBConn
        While Not objRs.EOF
            .AddNew objRs.Fields("testcd").Value & "", objRs.Fields("abbrnm10").Value & COL_DIV & objRs.Fields("testdiv").Value
            objRs.MoveNext
        Wend
    End With
    Set objRs = Nothing
    
    With objListPop
        .Connection = DBConn
        .FormCaption = "검사항목 조회"
        .ColumnHeaderText = "검사항목코드;검사명;구분"
        .ColumnHeaderWidth = "1440;1260.284;750.0473"
        .FormWidth = 3900
        .LoadPopUp objSQL.GetWAvsTest(medGetP(cboWA.Text, 1, " "))
        txtTestCd.Text = medGetP(.SelectedString, 1, ";")
        lblTestNm.Caption = medGetP(.SelectedString, 2, ";")
        Call GetRstCdList
    End With
    Set objListPop = Nothing
End Sub

Private Sub tblResult_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'    If Col = 15 Then
'        If Trim(aryResultText(Row)) <> "" Then
'            txtRst.TextRTF = aryResultText(Row)
'            txtRst.Visible = True
'            txtRst.ZOrder 0
'            DoEvents
'        End If
'    End If
End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    Static iSortOrder As Integer
    Dim I As Double
    
    '-- 추가 Colum별 Sort By M.G.Choi 2002.10.09
    With tblResult
        If Row = 0 Then
            .SortBy = SortByRow
            .SortKey(1) = Col
            If iSortOrder = SortKeyOrderAscending Then
                .SortKeyOrder(1) = SortKeyOrderDescending
                iSortOrder = SortKeyOrderDescending
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                iSortOrder = SortKeyOrderAscending
            End If
            .Col = 1
            .Col2 = .MaxCols
            .Row = 0
            .Row2 = .MaxRows
            .Action = ActionSort
        End If
'    End With
    
    If Col > 1 And Col < 5 Then
' 2008.12.17. 양성현 작업중입니다.
' 2009.01.09 양성현 환자ID 파라메터 추가
        Dim pFrmName As String
        Dim strPtId  As String
        .Col = 3
        .Row = Row
        strPtId = .Value
        If Len(strPtId) < 2 Then GoTo End2Stop

        pFrmName = "frm401ResultView"
    
        If ObjMyUser(pFrmName) Is Nothing Then GoTo End2Stop
        If Not ObjMyUser(pFrmName).CanRead Then GoTo End2Stop

'        medMain.lblSubMenu.Caption = "처방결과조회"

'        frmLisReviewInStatisticLib.ButtonKey = "LIS155B" 'Button.Key
'        frmLisReviewInStatisticLib.PTid = strPtId
'        frmLisReviewInStatisticLib.show
'        frmLisReview.show
'        frmLisReviewInStatisticLib.ShowThisForm
'        frmLisReviewInStatisticLib.ZOrder 0
End2Stop:
    Exit Sub


    End If
    If Col = 15 Then
' 2009.04.13 양성현 ary결과를 연계하기위해 i를 선언하고 버튼의 숫자를 Row로 설정함.
'    With tblResult
        .Row = Row: .Col = Col: I = Val(.TypeButtonText)
'    End With

    End If
    
    End With

End Sub

'마우스가 가면 포커스를 테이블로 옮기자 Tooltip 보여주기위해..
Private Sub tblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tblResult.SetFocus
End Sub

Private Sub cmdClear_Click()
    Call TxtClear
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    ' 이곳에서 이벤트를 발생시켜야 합니다.
    If IsLastForm Then RaiseEvent LastFormUnload
    RaiseEvent FormClose
End Sub

Private Sub cmdQuary_Click()
    Dim objProgress  As jProgressBar.clsProgress
    Dim RS           As New Recordset
    Dim RS1          As New Recordset
    Dim objPatient   As New clsPatient      '환자 클래스
    Dim SSQL         As String
    Dim strRstCdSql  As String
    Dim strDeptCd    As String
    Dim I            As Long
    Dim lngMaxHeight As Long
    Dim iCnt         As Integer
    Dim strDate      As String
    Dim strTmp       As Double
    Dim strWardTm    As String
    Dim strEmTm      As String
    Dim strOutTm     As String
    Dim strTotTm     As String
    Dim strEm1Tm     As String
    Dim strWorkArea  As String
    Dim strTestNm    As String
    Dim varTestNm    As Variant
    Dim iRowCnt      As Integer
    Dim varTmp
    Dim intCnt       As Integer
    Dim J, k, l, m, n, o, p As Integer
    Dim strAbNormal  As String
    Dim tmpAbNormal  As String
    Dim varTestCd    As Variant
    Dim str양성율    As String
    
    On Error Resume Next
    
    If cboWA.ListIndex < 0 Then
        MsgBox "WA(검사부서)를 입력하여 주세요", vbCritical, "조회조건"
        cboWA.ListIndex = 0
        Exit Sub
    End If
    
     '스프래드
    Call txtTblClear
    
    strWorkArea = Mid(cboWA.Text, 1, 2)
        
    Select Case strWorkArea
        Case "03": tmpAbNormal = "POSITIVE,REACTIVE"
        Case "05": tmpAbNormal = "POSITIVE,TRACE,1+,2+,3+,4+,RHO(D)+,DU(+),VARIANT"
        Case "07": tmpAbNormal = "POSITIVE,TRACE,1+,2+,3+,4+"
        Case Else: tmpAbNormal = "POSITIVE"
    End Select
    
    varTestCd = Split(tmpAbNormal, ",")
    
    strAbNormal = ""
    
    For intCnt = 0 To UBound(varTestCd) - 1
        strAbNormal = strAbNormal & "'" & varTestCd(intCnt) & "'" & ","
    Next
    
    strAbNormal = Mid(strAbNormal, 1, Len(strAbNormal) - 1)
    
'    strRstCdSql = RstCdSql
       
    '프로그래스바 생성..
    Set objProgress = New jProgressBar.clsProgress

    With objProgress
        .Container = Me
        .Width = tblResult.Width
        .Left = tblResult.Left
        .Top = tblResult.Top
        .Height = 530
        .Message = "결과내역을 검색하고 있습니다..."
    End With

    strDate = Format(dtpFromDt.Value, "yyyymm")
    
    SSQL = ""
    SSQL = SSQL & vbLf & "SELECT B.TESTNM, B.TESTCD, COUNT(a.testcd) AS abcount FROM S2LAB302 a, S2LAB001 b "
    SSQL = SSQL & vbLf & " WHERE substr(A.ACCDT,1,6) = '" & strDate & "' "
    
    If strWorkArea = "05" Then
        SSQL = SSQL & vbLf & "   AND A.WORKAREA = '" & strWorkArea & "'"
        SSQL = SSQL & vbLf & "   AND A.TESTCD IN ('B2047','B2061') "
        SSQL = SSQL & vbLf & "   AND A.RSTCD IN ('Positive','POSITIVE','TRACE','Trace','TR','1+','2+','3+','4+','P','T') "
    ElseIf strWorkArea = "03" Then
        SSQL = SSQL & vbLf & "   AND A.WORKAREA = '" & strWorkArea & "'"
        SSQL = SSQL & vbLf & "   AND A.RSTCD IN ('POSITIVE','REACTIVE','P') "
    Else
        SSQL = SSQL & vbLf & "   AND A.WORKAREA = '" & strWorkArea & "'"
        SSQL = SSQL & vbLf & "   AND A.HLDIV = 'N' "
    End If

    SSQL = SSQL & vbLf & "   AND A.TESTCD = B.TESTCD"
    If txtTestCd <> "" Then
        SSQL = SSQL & vbLf & "   AND A.TESTCD = '" & Trim(txtTestCd) & "' "
    End If
    SSQL = SSQL & vbLf & " GROUP BY B.TESTNM, B.TESTCD"
    
    RS.Open SSQL, DBConn
    
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        With tblResult
            For intCnt = 1 To RS.RecordCount
                If .MaxRows < intCnt Then .MaxRows = .MaxRows + 1
                .SetText 1, intCnt, Trim(Mid(cboWA.Text, 3))
                .SetText 2, intCnt, Trim(RS.Fields("TESTNM") & "")
                .SetText 3, intCnt, Trim(RS.Fields("TESTCD") & "")
                .SetText 5, intCnt, Trim(RS.Fields("abcount") & "")
                
                SSQL = ""
                SSQL = SSQL & vbLf & "SELECT A.TESTCD, COUNT(a.testcd) AS normatcount FROM S2LAB302 a"
                SSQL = SSQL & vbLf & " WHERE substr(A.ACCDT,1,6) = '" & strDate & "' "
                SSQL = SSQL & vbLf & "   AND A.WORKAREA = '" & strWorkArea & "'"
                SSQL = SSQL & vbLf & "   AND A.TESTCD =  '" & Trim(RS.Fields("TESTCD") & "") & "' "
                SSQL = SSQL & vbLf & " GROUP BY A.TESTCD"
                
                RS1.Open SSQL, DBConn
                
                .SetText 4, intCnt, Trim(RS1.Fields("normatcount") & "")
                str양성율 = Round(Val(Trim(RS.Fields("abcount") & "")) / Val(Trim(RS1.Fields("normatcount") & "")) * 100, 2)
                .SetText 6, intCnt, str양성율
                
                RS1.Close
                RS.MoveNext
            Next
        End With
    End If
    
    RS.Close
    
    With tblResult
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .BlockMode = False
    End With
    
    Set RS = Nothing
    Set RS1 = Nothing
    Set objPatient = Nothing
End Sub

'Private Function IcdSql() As String
'
'    If Trim(txtICd(0).Text) <> "" Then
'        IcdSql = "'" & Trim(txtICd(0).Text) & "'"
'    Else
'        IcdSql = ""
'    End If
'
'    If Trim(txtICd(1).Text) <> "" Then
'        If IcdSql <> "" Then
'            IcdSql = IcdSql & "," & "'" & Trim(txtICd(1).Text) & "'"
'        Else
'            IcdSql = "'" & Trim(txtICd(1).Text) & "'"
'        End If
'    End If
'
'    If Trim(txtICd(2).Text) <> "" Then
'        If IcdSql <> "" Then
'            IcdSql = IcdSql & "," & "'" & Trim(txtICd(2).Text) & "'"
'        Else
'            IcdSql = "'" & Trim(txtICd(2).Text) & "'"
'        End If
'    End If
'
'End Function

Private Function RstCdSql() As String
    
'    If Trim(txtRstCd(0).Text) <> "" Then
'        RstCdSql = "'" & Trim(txtRstCd(0).Text) & "'"
'    Else
'        RstCdSql = ""
'    End If
'
'    If Trim(txtRstCd(1).Text) <> "" Then
'        If RstCdSql <> "" Then
'            RstCdSql = RstCdSql & "," & "'" & Trim(txtRstCd(1).Text) & "'"
'        Else
'            RstCdSql = "'" & Trim(txtRstCd(1).Text) & "'"
'        End If
'    Else
'        If RstCdSql = "" Then RstCdSql = ""
'    End If
'
'    If Trim(txtRstCd(2).Text) <> "" Then
'        If RstCdSql <> "" Then
'            RstCdSql = RstCdSql & "," & "'" & Trim(txtRstCd(2).Text) & "'"
'        Else
'            RstCdSql = "'" & Trim(txtRstCd(2).Text) & "'"
'        End If
'    Else
'        If RstCdSql = "" Then RstCdSql = ""
'    End If

End Function

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
    If blnCHkLoad = False Then
        DoEvents
        blnCHkLoad = True
        Call GetWorkAreaCombo
    End If
End Sub

Private Sub Form_Load()
    blnCHkLoad = False
    TxtClear
    chkIndex_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSQL = Nothing
    Set objListPop = Nothing
''    Set objTMCd = Nothing
End Sub

Private Sub GetWorkAreaCombo()
    
    Dim sSqlGetWA As String
    Dim rsGetWA As Recordset
    Dim I%
    
    Set rsGetWA = New Recordset
    rsGetWA.Open objSQL.GetWACd, DBConn
    
    cboWA.Clear
    For I = 1 To rsGetWA.RecordCount
        cboWA.AddItem "" & rsGetWA.Fields("WACd").Value & "   " & _
                            "" & rsGetWA.Fields("WANm").Value
        rsGetWA.MoveNext
    Next I

    Set rsGetWA = Nothing

End Sub

Private Sub cmdListPop_Click(Index As Integer)
'    Dim objData As clsBasisData
    
    '리스트 팝업을 불러오자...
    Set objListPop = New clsPopUpList
'    Set objData = New clsBasisData
    
    With objListPop
        .Connection = DBConn
'        .BackColor = Me.BackColor
        Select Case Index
            '검체코드 불러오기
            Case 0:
'                .Caption = "검체코드 조회"
'                .HeadName = "검체코드, 검체명"
'                .Width = .Width + 700
'                Call .ListPop(objSql.GetSpcList, 2950, 4700)
'                txtSpcCd.Text = medGetP(.SelectedString, 1, ";")
'                lblTNm.Caption = medGetP(.SelectedString, 2, ";")
                
            '상병코드 불러오기
            Case 1:
'                If objIcdList Is Nothing Then
'                    Call GetIcdList
'                End If
'                .Caption = "상병코드 조회"
'                .HeadName = "상병코드, 상병명"
'                .Width = .Width + 700
'                Call .ListPop(, 3350, 4700, objIcdList)
'                If Trim(txtICd(0).Text) = "" Then
'                    txtICd(0).Text = medGetP(.SelectedString, 1, ";")
'                ElseIf Trim(txtICd(1).Text) = "" Then
'                    If Trim(txtICd(0).Text) = Trim(medGetP(.SelectedString, 1, ";")) Then
'                        txtICd(1).Text = ""
'                    Else
'                        txtICd(1).Text = medGetP(.SelectedString, 1, ";")
'                    End If
'                Else
'                    If Trim(txtICd(0).Text) = Trim(medGetP(.SelectedString, 1, ";")) Or _
'                       Trim(txtICd(1).Text) = Trim(medGetP(.SelectedString, 1, ";")) Then
'                        txtICd(2).Text = ""
'                    Else
'                        txtICd(2).Text = medGetP(.SelectedString, 1, ";")
'                    End If
'                End If
            '결과코드 불러오기
            Case 2:
                Dim objRstSQL As New clsLISSqlETest
                .FormCaption = "결과코드 조회"
                .ColumnHeaderText = "결과코드;결과명"
'                .Width = .Width + 700
                Call .LoadPopUp(objRstSQL.SqlGetSpeRstCode(txtTestCd.Text))  ', 3750, 4700, objRstCd)
'                If Trim(txtRstCd(0).Text) = "" Then
'                    txtRstCd(0).Text = medGetP(.SelectedString, 1, ";")
'                ElseIf Trim(txtRstCd(1).Text) = "" Then
'                    If Trim(txtRstCd(0).Text) = Trim(medGetP(.SelectedString, 1, ";")) Then
'                        txtRstCd(1).Text = ""
'                    Else
'                        txtRstCd(1).Text = medGetP(.SelectedString, 1, ";")
'                    End If
'                Else
'                    If Trim(txtRstCd(0).Text) = Trim(medGetP(.SelectedString, 1, ";")) Or _
'                       Trim(txtRstCd(1).Text) = Trim(medGetP(.SelectedString, 1, ";")) Then
'                        txtRstCd(2).Text = ""
'                    Else
'                        txtRstCd(2).Text = medGetP(.SelectedString, 1, ";")
'                    End If
'                End If
                Set objRstSQL = Nothing
            '진료과 불러오기
            Case 3:
                .FormCaption = "진료과 조회"
                .ColumnHeaderText = "진료과코드;진료과명"
'                .Width = .Width + 300
'                .ColSize(0) = 1000
                Call .LoadPopUp(GetSQLDeptList) ', 3950, 9300) ', ObjLISComCode.DeptCd)
'                txtDeptCd.Text = medGetP(.SelectedString, 1, ";")
'                lblDeptNm.Caption = medGetP(.SelectedString, 2, ";")
'
            Case 4:
'                .Caption = "검체코드 조회"
'                .HeadName = "검체코드, 검체명"
'                .Width = .Width + 700
'                Call .ListPop(objSql.GetSpcListByTest(txtTestCd.Text), 2950, 4700)
'                txtSpcCd.Text = medGetP(.SelectedString, 1, ";")
'                lblTNm.Caption = medGetP(.SelectedString, 2, ";")
        End Select
    End With
'    Set objData = Nothing
    Set objListPop = Nothing
    
End Sub

Private Sub TxtClear()
    
   
    '조회기간
    dtpFromDt.Value = GetSystemDate
    dtpToDt.Value = GetSystemDate
       
    '스프래드
    Call txtTblClear
End Sub

Private Sub txtTblClear()
    medClearTable tblResult
    tblResult.MaxRows = 0
    tblResult.RowHeight(-1) = 15

'    cmdPrint.Enabled = False
    cmdExcel.Enabled = True
End Sub

'Private Sub txtAccDt_LostFocus()
'    If Trim(txtAccDt.Text) <> "" And Len(txtAccDt.Text) >= 2 Then
'        dtpFromDt.Year = "20" & Mid(txtAccDt.Text, 1, 2)
'    End If
'End Sub
'
'Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtDeptCd_LostFocus()
''    Dim objDept As clsBasisData
'    Dim strDept As String
'
'    If Trim(txtDeptCd.Text) = "" Then
'        lblDeptNm.Caption = ""
'        Exit Sub
'    End If
'
''    Set objDept = New clsBasisData
'    strDept = GetDeptNm(txtDeptCd.Text)
''    Set objDept = Nothing
'
'    If strDept <> "" Then
'        lblDeptNm.Caption = strDept
'    Else
'        medBeep (1)
'        txtDeptCd.Text = ""
'        lblDeptNm.Caption = ""
'        txtDeptCd.SetFocus
'        Exit Sub
'    End If
''
''    With ObjAPSComCode.DeptCd
''
''        If .Exists(Trim(txtDeptCd.Text)) = True Then
''            .KeyChange Trim(txtDeptCd.Text)
''            lblDeptNm.Caption = .Fields("deptnm")
''        Else
''            medbeep (1)
''            txtDeptCd.Text = ""
''            lblDeptNm.Caption = ""
''            txtDeptCd.SetFocus
''            Exit Sub
''        End If
''    End With
'End Sub

Private Sub txtFromSeq_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

'Private Sub txtPtId_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtPtId_LostFocus()
'    Dim objPatient As New clsPatient      '환자 클래스
'
'    If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
'
'    With objPatient
'        If Trim(txtPtId.Text) <> "" Then
'            If .GETPatient(txtPtId.Text) Then
'                lblPtInfo.Caption = .PtNm & "   " & .SEXNM & " / " & .Age & " " & .AGEDIV
'            Else
'                lblPtInfo.Caption = ""
'                MsgBox "등록되지 않은 환자ID 입니다.", vbExclamation, "메세지"
'                Exit Sub
'            End If
'        Else
'            lblPtInfo.Caption = ""
'        End If
'    End With
'    Set objPatient = Nothing
'End Sub

'Private Sub txtRst_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 27 Then
'        txtRst.Visible = False
'    End If
'End Sub
'
'Private Sub txtRstCd_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub txtRstCd_LostFocus(Index As Integer)
'
'    If Trim(txtRstCd(Index).Text) = "" Then Exit Sub
'
'    With objRstCd
'        If .Exists(Trim(txtRstCd(Index).Text)) = True Then
'            Exit Sub
'        Else
'            medBeep (1)
'            txtRstCd(Index).Text = ""
'            Exit Sub
'        End If
'    End With
'
'End Sub

Private Sub PrintSpread()
    Dim objValue    As New clsDictionary
    Dim I           As Long
    Dim J           As Long
    Dim strLabNo    As String
    Dim strPtNm     As String
    Dim strPtId     As String
    Dim strSpcnm    As String
    Dim strDeptCd   As String
    Dim strDx       As String
    Dim strData     As String
    
    objValue.Clear
    objValue.FieldInialize "labno", "ptnm,ptid,spcnm,deptcd,dx"
    
    With tblResult
        For I = 1 To .MaxRows
            .Row = I
            For J = 1 To .MaxCols
                .Col = J
                Select Case J
                    Case 1: strLabNo = .Value
                    Case 2: strPtNm = .Value
                    Case 3: strPtId = .Value
                    Case 5: strSpcnm = .Value
                    Case 9: strDeptCd = .Value
                    Case 11: strDx = .Value
                End Select
            Next J
            strData = Join(Array(strPtNm, strPtId, strSpcnm, strDeptCd, strDx), COL_DIV)
            objValue.AddNew strLabNo, strData
        Next I
    End With
    
    Set objValue = Nothing
    
End Sub

Private Sub GetIcdList()

    Dim objRs As Recordset
'    Dim objIcdSql   As New clsBasisData  'clsHosComSQLStmt
    Dim objStatus As New jProgressBar.clsProgress
    
    With objStatus
        .Container = Me
        .Width = lblTitle.Width
        .Left = lblTitle.Left
        .Top = lblTitle.Top
        .Height = 280
        .Message = "상병코드 마스터를 로드하고 있습니다..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = lblTitle.Width
'        .XPos = lblTitle.Left
'        .YPos = lblTitle.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "상병코드 마스터를 로드하고 있습니다..."
'        .Value = 0
    End With

    Set objIcdList = New clsDictionary
    objIcdList.Clear
    objIcdList.FieldInialize "icd", "icdenm"
    
    Set objRs = New Recordset
    objRs.Open GetSQLIcdList, DBConn
    
    objStatus.Max = objRs.RecordCount
    
    objIcdList.Sort = False
    While Not objRs.EOF
        objStatus.Value = objStatus.Value + 1
        objStatus.Message = "상병코드 마스터를 로드하고 있습니다...(" & CInt(objStatus.Value / objStatus.Max * 100) & "%)"
        objIcdList.AddNew objRs.Fields("icd").Value & "", objRs.Fields("ienm").Value & ""
        objRs.MoveNext
    Wend
    
    Set objRs = Nothing
'    Set objIcdSql = Nothing
    Set objStatus = Nothing
    
End Sub

Private Sub GetRstCdList()

    Dim objRs As Recordset
    Dim objRstSQL As New clsLISSqlETest

    Set objRstCd = New clsDictionary
    objRstCd.Clear
    objRstCd.FieldInialize "rstcd", "rstnm"
    
    Set objRs = New Recordset
    objRs.Open objRstSQL.SqlGetSpeRstCode(txtTestCd.Text), DBConn
    
    objRstCd.Sort = False
    While Not objRs.EOF
        objRstCd.AddNew objRs.Fields("rstcd").Value & "", objRs.Fields("rstnm").Value & ""
        objRs.MoveNext
    Wend
    objRstCd.Sort = True
    
    Set objRs = Nothing
    Set objRstSQL = Nothing
    
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        Call txtTestCd_LostFocus
    End If
End Sub

Private Sub txtTestCd_LostFocus()

    Dim strSQL As String
    Dim objRs As Recordset
    
    Call TxtClear
    Call txtTblClear
    
    If Trim(txtTestCd.Text) = "" Then Exit Sub
    
    strSQL = objSQL.GetAccTest(txtTestCd.Text)
    Set objRs = New Recordset
    objRs.Open strSQL, DBConn
    
    If objRs.EOF Then
        MsgBox "처방코드를 다시 입력하십시오.", vbInformation, "처방코드 입력"
        Set objRs = Nothing
        txtTestCd.SelStart = 0
        txtTestCd.SelLength = Len(txtTestCd.Text)
        txtTestCd.SetFocus
        Exit Sub
    Else
        lblTestNm.Caption = "" & objRs.Fields("abbrnm10").Value
    End If
    
    Set objRs = Nothing
    
    Call GetRstCdList
End Sub

Private Sub txtToSeq_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdPrint_Click()
    Dim objSpread    As vaSpread
    Dim strTitle     As String
    Dim strPrintDate As String
    Dim strTestNm    As String
    Dim strPDate     As String
    Dim tmpTitle     As String
    Dim strDate      As String
    Dim strGb        As String
    
    strGb = ""
    strPDate = Format(Now, "yyyy-mm-dd hh:mm:ss")

    With tblResult
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .FontBold = False
        .FontSize = 9
        .BlockMode = False
               
        .PrintJobName = "CVR 보고체계 시행율"

        .PrintAbortMsg = "CVR 보고체계 시행율를 출력중입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1
        
        tmpTitle = "CVR 보고체계 시행율"
'        strTitle = "/fn""굴림체""/fz""18""/fb1/fi0/fu1/fk0/fs1" _
'              & "/f1/c" & tmpTitle & "/n/n/n"
        strTitle = "/fn""굴림체"" /fz""18"" /fb1/fi0/fu0/fk0/fs1" _
                  & "/f1/c" & tmpTitle & "/n/n/n"
        strPrintDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "출력일자 : " & strPDate & "/n/n"
        strTestNm = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "WorkArea : " & cboWA.Text & "   검사항목 : " & lblTestNm.Caption & "/n"
        strDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "조회기간 : " & Format(dtpFromDt.Value, "yyyy") & "/n"
        .PrintHeader = strTitle & strTestNm & strDate 'strPrintDate
        .PrintMarginLeft = 10
'        .PrintMarginRight = 10
        .PrintOrientation = PrintOrientationPortrait 'PrintOrientationLandscape
'        .PrintOrientation = PrintOrientationLandscape 'PrintOrientationLandscape
        
        
'        P_HOSPITALNAME = "한마음혈액원"
        .PrintFooter = " /l " & String(130, Chr(6)) & "/n/l " & P_HOSPITALNAME & "/c/p/fb1"
     
        .PrintMarginBottom = 100
        .PrintShadows = True
        .PrintMarginTop = 300
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint

    End With
End Sub

'Private Sub CaseStudyHead()
'    Dim strTmp  As String
'
'    lngCurYPos = 10
'    Printer.DrawStyle = 0: Printer.DrawWidth = 6
'    Printer.FontSize = 20: Printer.FontBold = True
'    Call Print_Setting("Case Study", 0, LineSpace * 3, Printer.ScaleWidth - 0, "C", "C", True)
'    Printer.FontSize = 9: Printer.FontBold = False
'
'    strTmp = Format(dtpFromDt.Value, "YYYY년 MM월 DD일") & " ~ " & Format(dtpToDt.Value, "YYYY년 MM월 DD일")
'
'    Call Print_Setting("조회기간 : " & strTmp, 0, LineSpace, Printer.ScaleWidth, "L", "C")
'    Call Print_Setting("업무영역 : " & cboWA.Text, 120, LineSpace, Printer.ScaleWidth, "L", "C", False)
'    Call Print_Setting("검사항목 : " & txtTestCd.Text & "[" & lblTestNm.Caption & "]", 0, LineSpace, Printer.ScaleWidth, "L", "C")
'    strTmp = "[ 전체 ]": If txtPtId.Text <> "" Then strTmp = "[ " & txtPtId.Text & " ] " & lblPtInfo.Caption
'    Call Print_Setting("환자조건 : " & strTmp, 0, LineSpace, Printer.ScaleWidth, "L", "C", False)
'    strTmp = "[ 전체 ]": If txtAccDt.Text <> "" Then strTmp = "[ " & txtAccDt.Text & " ] " & txtFromSeq.Text & " ~ " & txtToSeq.Text
'    Call Print_Setting("접수번호 : " & strTmp, 120, LineSpace, Printer.ScaleWidth, "L", "C")
'    strTmp = "[ 전체 ]": If txtRstCd(0).Text <> "" Then strTmp = "[ " & txtRstCd(0).Text & " ] " & txtRstCd(1).Text & " ~ " & txtRstCd(2).Text
'    Call Print_Setting("결과코드 : " & strTmp, 0, LineSpace, Printer.ScaleWidth, "L", "C", False)
'    strTmp = "[ 전체 ]": If txtDeptCd.Text <> "" Then strTmp = "[ " & txtDeptCd.Text & " ] " & lblDeptNm.Caption
'    Call Print_Setting("의 뢰 과 : " & strTmp, 120, LineSpace, Printer.ScaleWidth, "L", "C")
'    strTmp = Format(GetSystemDate, "YYYY년 MM월 DD일")
'    Call Print_Setting("출 력 일 : " & strTmp, 0, LineSpace, Printer.ScaleWidth, "L", "C")
'
'    Printer.Line (0, lngCurYPos)-(Printer.Width - 0, lngCurYPos)
'
'    '-- 원본
''    Call CaseStudyBody("접수번호", "환자ID", "환자명", "성/나이", "검체명", "접수일자", "진료과", _
'                       "병동", "결과1", "결과2", "결과3", "text결과")
'
'    Call CaseStudyBody("접수번호", "환자ID", "환자명", "성/나이", "검체명", "접수일자", "보고일자", _
'                       "진료과", "병동", "결과1", "", "", "")
'
'    Printer.DrawStyle = 0: Printer.DrawWidth = 6
'    Printer.Line (0, lngCurYPos)-(Printer.Width - 0, lngCurYPos)
'End Sub
'
'Private Sub CaseStudyBody(ByVal sAccno As String, ByVal sPtid As String, ByVal sPtnm As String, _
'                          ByVal sSexAge As String, ByVal sSpcNm As String, ByVal sAccDt As String, _
'                          ByVal sVfydt As String, ByVal sDept As String, ByVal sWard As String, ByVal sRst1 As String, _
'                          ByVal sRst2 As String, ByVal sRst3 As String, ByVal sTxtFg As String)
'
'    If lngCurYPos > Printer.ScaleHeight - 6 Then
'        Printer.NewPage
'        Call CaseStudyHead
'    End If
'
'    Call Print_Setting(sAccno, 0, LineSpace, 30, "L", "C", False)
'    Call Print_Setting(sPtid, 25, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sPtnm, 40, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sSexAge, 55, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sSpcNm, 70, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sAccDt, 85, LineSpace, 20, "L", "C", False)
'    Call Print_Setting(sVfydt, 120, LineSpace, 20, "L", "C", False)
'    Call Print_Setting(sDept, 155, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sWard, 170, LineSpace, 15, "L", "C", False)
'    Call Print_Setting(sRst1, 185, LineSpace, 15, "L", "C")
'
'    '** 원본 -------------------------------------------------------
''    Call Print_Setting(sDept, 105, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sWard, 120, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sRst1, 135, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sRst2, 150, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sRst3, 165, LineSpace, 15, "L", "C", False)
''    Call Print_Setting(sTxtFg, 180, LineSpace, 35, "L", "C")
'    '---------------------------------------------------------------
'    Printer.DrawStyle = 2: Printer.DrawWidth = 2
'    Printer.Line (0, lngCurYPos)-(Printer.Width - 0, lngCurYPos)
'End Sub
'
'Private Sub PrintCaseStudy()
'    Dim sAccno  As String
'    Dim sPtid   As String
'    Dim sPtnm   As String
'    Dim sSexAge As String
'    Dim sSpcNm  As String
'    Dim sAccDt  As String
'    Dim sVfydt  As String
'    Dim sDept   As String
'    Dim sWard   As String
'    Dim sRst1   As String
'    Dim sRst2   As String
'    Dim sRst3   As String
'    Dim sTxtFg  As String
'
'    Dim ii          As Integer
'
'    If tblResult.DataRowCnt < 1 Then Exit Sub
'
'    Call P_PrtSet
'    Call CaseStudyHead
'
'    With tblResult
'        For ii = 1 To .DataRowCnt
'            .Row = ii
'            .Col = 1:   sAccno = .Value
'            .Col = 2:   sPtid = .Value
'            .Col = 3:   sPtnm = .Value
'            .Col = 4:   sSexAge = .Value
'            .Col = 5:   sSpcNm = .Value
'            .Col = 7:   sAccDt = .Value
'            .Col = 8:   sVfydt = .Value
'            .Col = 9:   sDept = .Value
'            .Col = 10:   sWard = .Value
'            .Col = 11:   sRst1 = .Value
''            .Col = 12:  sRst2 = .Value
''            .Col = 13:  sRst3 = .Value
''            .Col = 14:  sTxtFg = "Y"
'                        If .CellType = CellTypeStaticText Then sTxtFg = ""
'            Call CaseStudyBody(sAccno, sPtid, sPtnm, sSexAge, sSpcNm, sAccDt, sVfydt, sDept, sWard, sRst1, sRst2, sRst3, sTxtFg)
'        Next
'    End With
'
'    Printer.EndDoc
'End Sub



