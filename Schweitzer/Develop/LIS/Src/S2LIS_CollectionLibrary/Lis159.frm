VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm159BatchBarReprint 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15270
   ControlBox      =   0   'False
   Icon            =   "Lis159.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCorpCd 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   315
      Left            =   6090
      TabIndex        =   20
      Top             =   270
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdCorpList 
      BackColor       =   &H009CA7B8&
      Caption         =   "▼"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6975
      MousePointer    =   14  '화살표와 물음표
      Style           =   1  '그래픽
      TabIndex        =   19
      Top             =   270
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtToSeq 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   11730
      TabIndex        =   16
      Top             =   690
      Width           =   1350
   End
   Begin VB.TextBox txtFromSeq 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   10020
      TabIndex        =   15
      Top             =   690
      Width           =   1350
   End
   Begin VB.CheckBox chkOut 
      BackColor       =   &H00DFE3E8&
      Caption         =   "외래"
      Height          =   255
      Left            =   3330
      TabIndex        =   11
      Top             =   375
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.CommandButton cmdGetOrders 
      BackColor       =   &H00E0E0E0&
      Caption         =   "조회(&Q)"
      Height          =   660
      Left            =   6405
      Style           =   1  '그래픽
      TabIndex        =   10
      Tag             =   "15101"
      Top             =   750
      Width           =   1320
   End
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "재출력(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   8535
      Width           =   1320
   End
   Begin VB.ComboBox cboColTm 
      Height          =   300
      Left            =   4200
      Style           =   2  '드롭다운 목록
      TabIndex        =   6
      Top             =   1125
      Width           =   2130
   End
   Begin VB.CommandButton cmdWardList 
      BackColor       =   &H009CA7B8&
      Caption         =   "▼"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2490
      MousePointer    =   14  '화살표와 물음표
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   585
      Width           =   300
   End
   Begin VB.TextBox txtWardId 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   315
      Left            =   1155
      TabIndex        =   1
      Top             =   585
      Width           =   1305
   End
   Begin MedControls1.LisLabel lblWardNm 
      Height          =   315
      Left            =   2820
      TabIndex        =   0
      Top             =   585
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      BackColor       =   15988216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin MSComCtl2.DTPicker dtpToTime 
      Height          =   330
      Left            =   1155
      TabIndex        =   3
      Top             =   1110
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
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
      Format          =   23724032
      UpDown          =   -1  'True
      CurrentDate     =   36342.5951388889
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   345
      Left            =   75
      TabIndex        =   12
      Top             =   45
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   609
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "재발행대상조회"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   345
      Left            =   7845
      TabIndex        =   13
      Top             =   45
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   609
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "출력범위설정"
      LeftGab         =   100
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   6900
      Left            =   75
      TabIndex        =   14
      Tag             =   "10114"
      Top             =   1545
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   12171
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   15924219
      GridColor       =   14737632
      MaxCols         =   27
      MaxRows         =   10
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis159.frx":08CA
      StartingColNumber=   2
      UserResize      =   1
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   10
   End
   Begin MedControls1.LisLabel lblCorpNm 
      Height          =   315
      Left            =   7305
      TabIndex        =   21
      Top             =   270
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      BackColor       =   15988216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "거래처코드"
      Height          =   225
      Left            =   5115
      TabIndex        =   22
      Tag             =   "15105"
      Top             =   330
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "출력범위 :     From                         To "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   8340
      TabIndex        =   18
      Top             =   735
      Width           =   3705
   End
   Begin VB.Label lblDt 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "출력일시"
      Height          =   180
      Left            =   300
      TabIndex        =   5
      Tag             =   "15104"
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label lblWardLabel 
      BackStyle       =   0  '투명
      Caption         =   "부서코드"
      Height          =   225
      Left            =   300
      TabIndex        =   4
      Tag             =   "15105"
      Top             =   645
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      FillColor       =   &H00DFE3E8&
      FillStyle       =   0  '단색
      Height          =   1080
      Left            =   90
      Top             =   420
      Width           =   7740
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00ABABAB&
      FillColor       =   &H00DFE3E8&
      FillStyle       =   0  '단색
      Height          =   405
      Left            =   7980
      Top             =   615
      Width           =   6405
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "☞ 처방내역을 검색중입니다..."
      ForeColor       =   &H00553755&
      Height          =   270
      Left            =   8070
      TabIndex        =   17
      Top             =   1170
      Width           =   6435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CCFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      Height          =   300
      Left            =   7950
      Shape           =   4  '둥근 사각형
      Top             =   1110
      Width           =   6450
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      FillColor       =   &H00DFE3E8&
      FillStyle       =   0  '단색
      Height          =   1095
      Left            =   7860
      Top             =   405
      Width           =   6600
   End
End
Attribute VB_Name = "frm159BatchBarReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'** 주의 :  건물구분을 OCS프로그램에서 넘겨준 부서코드로 부서마스터를 검색해서
'           bld_gb를 가져온다.

Option Explicit

'---- Collect
Private objSQL                  As clsLISSqlCollection
Private MySql                   As clsLISSqlStatement
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private CleanFg                 As Boolean
Private IsFirst                 As Boolean
Private intPtCount              As Integer
Private intErrCount             As Integer
Private mvarDeptCd              As String

Public Event LastFormUnload()
'
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property
Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property


Private Sub cboColTm_Click()
    txtFromSeq.Text = ""
    txtToSeq.Text = ""
    TableClear (0)
End Sub

Private Sub cboColTm_GotFocus()
    If Trim(txtWardID.Text) = "" Then
        Exit Sub
    End If
    
'    If Trim(txtCorpCd.Text) = "" Then
'        Exit Sub
'    End If
    
    Dim Rs      As Recordset
    Dim SqlStmt As String

    Set Rs = New Recordset
    Rs.Open MySql.SqlGetColTimes(Format(dtpToTime.Value, CS_DateDbFormat), txtWardID.Text), DBConn
    
    cboColTm.Clear
    While (Not Rs.EOF)
        cboColTm.AddItem Format("" & Rs.Fields("WorkTm").Value, CS_TimeLongMask)
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If cboColTm.ListCount <= 0 Then
        MsgBox "해당일엔 작업하신 채취내역이 없습니다.", vbInformation, "채취"
        cmdGetOrders.Enabled = False
    Else
        cboColTm.ListIndex = 0
        cmdGetOrders.Enabled = True
    End If
End Sub

Private Sub cboColTm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboColTm.ListCount > 0 Then cmdGetOrders.SetFocus
    End If
End Sub

Private Sub chkOut_Click()
    
    lblWardLabel.Caption = IIf(chkOut.Value = 1, "진료과", "병동 ID")
    
End Sub

Private Sub cmdClear_Click()
   Call ClearRtn(1)
   txtWardID.SetFocus
End Sub

Private Sub cmdCorpList_Click()
    Dim strSQL  As String
    
    strSQL = " select custcode, custname " & _
             "   from oras1.sg1custt "
    
    '% 거래처정보 리스트를 팝업한다.
    Set objMyList = New clsPopUpList
    
    txtCorpCd.Text = "": lblCorpNm.Caption = ""
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "거래처 조회"
        .ColumnHeaderText = "거래처코드;거래명"
        .Tag = "CorpID"
        Me.ScaleMode = 1
        Call .LoadPopUp(strSQL)
         
        If .SelectedString <> "" Then
            txtCorpCd.Text = medGetP(.SelectedString, 1, ";")
            lblCorpNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    
    Set objMyList = Nothing
    
End Sub

Private Sub cmdGetOrders_Click()

    If Trim(txtWardID.Text) = "" Then
        MsgBox "병동ID를 입력하세요.", vbInformation, "병동입력"
        txtWardID.SetFocus
        Exit Sub
    End If
    tblOrdSheet.MaxRows = 0
    DoEvents
    lblMessage.Caption = "채취내역을 조회중입니다..."
    DoEvents
    MouseRunning
    Call DisplayOrder
    MouseDefault
    lblMessage.Caption = ""

End Sub

Private Sub cmdReprint_Click()

    Dim objBAR As clsBarcode
    Dim BarBuffer(0 To 15)  As String
    Dim TestNames           As String
    Dim tmpLabNo            As Variant
    Dim strOrdDiv           As String
    Dim AccFg               As Boolean
    Dim FzFg                As Boolean
    
    Dim i                   As Long
    
    Set objBAR = New clsBarcode
'    Set objBAR.MyDB = dbconn
    Set objBAR.TableInfo = New clsTables
    Set objBAR.FieldInfo = New clsFields
    
    lblMessage.Caption = "Barcode Label을 출력중입니다..."
    Me.MousePointer = 11
    DoEvents

    With tblOrdSheet
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If (Val(.Value) >= Val(txtFromSeq.Text)) And _
               (Val(.Value) <= Val(txtToSeq.Text)) Then
                Erase BarBuffer
                .Col = 26:   BarBuffer(0) = .Value              '처방구분
                .Col = 14:   BarBuffer(2) = .Value              'WorkArea
                .Col = 17:   BarBuffer(3) = Mid(.Value, 3)      'AccDt
                .Col = 15:
                             Select Case BarBuffer(0)           'OrdDiv
                                 Case BBS_ORDDIV:
                                     .Col = 8
                                     BarBuffer(4) = Format(.Value, String(11, "@"))
                                 Case Else:
                                     .Col = 15
                                     BarBuffer(4) = IIf(.Value = "0", "", Format(.Value, String(4, "@")))    'SpcNo
                             End Select
                .Col = 21:
                        If P_ApplyBuildingInfo Then
                            If BarBuffer(0) = APS_ORDDIV Then
                                BarBuffer(1) = APSName
                            Else
                                BarBuffer(1) = Mid(.Value, 1, 2)        '건물명
                            End If
                        Else
                            Select Case BarBuffer(0)
                                Case LIS_ORDDIV: BarBuffer(1) = LABName
                                Case BBS_ORDDIV: BarBuffer(1) = BBSName
                                Case APS_ORDDIV: BarBuffer(1) = APSName
                            End Select
                        End If
                            
                .Col = 20:   BarBuffer(5) = .Value              'SpcNo
                .Col = 2:    BarBuffer(6) = .Value              '환자ID
                .Col = 3:    BarBuffer(7) = .Value              '환자ID
                .Col = 13:   BarBuffer(8) = .Value              '검체명
                .Col = 16:   BarBuffer(9) = .Value              '보관코드
                .Col = 18:   BarBuffer(10) = .Value             'StatFg
                .Col = 22:   BarBuffer(11) = .Value             '병동ID
                             If .Value = "" Then _
                                 .Col = 23: BarBuffer(11) = .Value   '진료과코드
                .Col = 9:    BarBuffer(12) = Mid(.Value, 5, 2) & "/" & Mid(.Value, 7, 2)
                .Col = 25:   BarBuffer(13) = .Value             '채혈시간
                .Col = 5:    TestNames = Mid(.Value, 1, Len(.Value) - 1)
                 BarBuffer(14) = TestNames                      '검사명
                 BarBuffer(15) = 1                              '라벨출력장수
                .Col = 24:
                             AccFg = IIf(Val(.Value) >= 2, True, False)
                .Col = 27:
                             FzFg = IIf(.Value <> "", True, False)   'Frozen Fg

                objBAR.WardTmp = ""
                Call objBAR.Label_PrintOut( _
                        BarBuffer(1), BarBuffer(2), BarBuffer(3), BarBuffer(4), BarBuffer(5), BarBuffer(6), _
                        BarBuffer(7), BarBuffer(8), BarBuffer(9), BarBuffer(10), BarBuffer(11), BarBuffer(12), _
                        BarBuffer(13), BarBuffer(14), BarBuffer(15), AccFg, FzFg)
                
'                Call medSleep(2000)
            End If
        Next
    End With
    Set objBAR = Nothing
    Call ClearRtn(0)
    Me.MousePointer = 0
    
    lblMessage.Caption = ""

End Sub

Private Sub dtpToTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpToTime_LostFocus()
'    Dim Rs      As Recordset
'    Dim SqlStmt As String
'
'    Set Rs = New Recordset
'    Rs.Open MySql.SqlGetColTimes(Format(dtpToTime.Value, CS_DateDbFormat), txtWardID.Text), DBConn
'
'    cboColTm.Clear
'    While (Not Rs.EOF)
'        cboColTm.AddItem Format("" & Rs.Fields("WorkTm").Value, CS_TimeLongMask)
'        Rs.MoveNext
'    Wend
'    Set Rs = Nothing
'
'    If cboColTm.ListCount <= 0 Then
'        MsgBox "해당일엔 작업하신 채취내역이 없습니다.", vbInformation, "채취"
'        cmdGetOrders.Enabled = False
'    Else
'        cboColTm.ListIndex = 0
'        cmdGetOrders.Enabled = True
'    End If
End Sub

Private Sub dtpToTime_Validate(Cancel As Boolean)
'    Dim Rs      As Recordset
'    Dim SqlStmt As String
'
'    Set Rs = New Recordset
'    Rs.Open MySql.SqlGetColTimes(Format(dtpToTime.Value, CS_DateDbFormat), txtWardID.Text), DBConn
'
'    cboColTm.Clear
'    While (Not Rs.EOF)
'        cboColTm.AddItem Format("" & Rs.Fields("WorkTm").Value, CS_TimeLongMask)
'        Rs.MoveNext
'    Wend
'    Set Rs = Nothing
'
'    If cboColTm.ListCount <= 0 Then
'        MsgBox "해당일엔 작업하신 채취내역이 없습니다.", vbInformation, "채취"
'        cmdGetOrders.Enabled = False
'    Else
'        cboColTm.ListIndex = 0
'        cmdGetOrders.Enabled = True
'    End If
End Sub

Private Sub Form_Activate()

    If Not IsFirst Then Exit Sub
    IsFirst = False
    
    Call ClearRtn(0)
    dtpToTime.Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
    intErrCount = 0
    txtWardID.Text = gWardId
    lblWardNm.Caption = gWardNm
    txtWardID.SetFocus
    
End Sub

Private Sub Form_Load()
    
    IsFirst = True
    Set objSQL = New clsLISSqlCollection
    Set MySql = New clsLISSqlStatement

End Sub

'% 종료
Private Sub cmdExit_Click()
    Unload Me
    Set objMyList = Nothing
    Set MySql = Nothing
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

' 기준시간이 변경되면 Clear
Private Sub dtpToTime_Change()
    If Not CleanFg Then
        Call TableClear(1)
        cboColTm.Clear
    End If
End Sub

'% 병동코드 리스트를 팝업한다.
Private Sub cmdWardList_Click()
'    Dim objDept As clsBasisData
    
    Set objMyList = New clsPopUpList
'    Set objDept = New clsBasisData
    
    txtWardID.Text = "": lblWardNm.Caption = ""
    With objMyList
        If chkOut.Value = 0 Then
            .Connection = DBConn
            .FormCaption = "진료과 조회"
            .ColumnHeaderText = "부서코드;부서명"
            .Tag = "WardID"
             Me.ScaleMode = 1
             Call .LoadPopUp(objSQL.SqlGetBatchDept) ', 2700, cmdWardList.Left)
        Else
            .Connection = DBConn
            .FormCaption = "병동 조회"
            .ColumnHeaderText = "병동코드;병동명"
            .Tag = "WardID"
             Me.ScaleMode = 1
             Call .LoadPopUp(GetSQLWardList) ', 2750, 2650)
             
        End If
        If .SelectedString <> "" Then
            txtWardID.Text = medGetP(.SelectedString, 1, ";")
            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With

'    Set objDept = Nothing
    Set objMyList = Nothing

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set MySql = Nothing
    Set objSQL = Nothing
End Sub

Private Sub txtCorpCd_Change()
    If Not CleanFg Then
        Call TableClear(0)
        cboColTm.Clear
    End If
End Sub

Private Sub txtCorpCd_GotFocus()
    With txtCorpCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCorpCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objMyList Is Nothing Then
            dtpToTime.SetFocus
        End If
    End If
End Sub

Private Sub txtCorpCd_KeyPress(KeyAscii As Integer)
    On Error GoTo Err_Trap

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = vbKeyReturn Then
        If txtCorpCd.Text = "" Then
            txtCorpCd.SetFocus
            Exit Sub
        Else
            
            Dim strCorp As String
            
            strCorp = GetCorpNm(txtCorpCd.Text)
            
            If strCorp = "" Then
                MsgBox "거래처코드를 확인하세요.", vbExclamation
                txtCorpCd.Text = ""
                Call cmdCorpList_Click
                Exit Sub
            Else
                lblCorpNm.Caption = strCorp
                SendKeys "{TAB}"
            End If
        End If
    End If
    
    Exit Sub

Err_Trap:
    Resume Next

End Sub

Private Function GetCorpNm(ByVal pCorpCd As String) As String
    Dim strSQL  As String
    Dim Rs      As New ADODB.Recordset
    
    strSQL = " select custname from oras1.sg1custt " & _
             "  where custcode = " & DBS(pCorpCd)
             
    Rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        GetCorpNm = Rs.Fields("custname").Value & ""
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Function

Private Sub txtFromSeq_Change()
    If Val(txtFromSeq.Text) > 0 Then tblOrdSheet.TopRow = Val(txtFromSeq.Text)
End Sub

'% 대상 병동이 변경되면 Clear
Private Sub txtWardId_Change()
    If Not CleanFg Then
        Call TableClear(0)
        cboColTm.Clear
    End If
End Sub

Private Sub ClearRtn(ByVal intOpt As Integer)
    'Unlocking...
    txtWardID.Enabled = True
    txtWardID.BackColor = &H80000005
    cmdWardList.Enabled = True
    dtpToTime.Enabled = True
    txtFromSeq.Text = ""
    txtToSeq.Text = ""

    txtWardID.Text = ""
    lblWardNm.Caption = ""
    lblMessage.Caption = ""
    dtpToTime.Value = Format(Now, "YYYY/MM/DD HH:MM:SS")
    Call TableClear(0)
    cmdGetOrders.Enabled = False
    CleanFg = True
End Sub
'% Table들을 Clear한다
Private Sub TableClear(ByVal intOpt As Integer)
    tblOrdSheet.MaxRows = 0
    CleanFg = True
End Sub

'% 병동 ID
Private Sub txtWardId_GotFocus()
    With txtWardID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWardId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objMyList Is Nothing Then
            If lblCorpNm.Caption <> "" Then
                dtpToTime.SetFocus
            Else
                Call cmdCorpList_Click
            End If
        End If
    End If
End Sub


Private Sub txtWardId_KeyPress(KeyAscii As Integer)

On Error GoTo Err_Trap

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = vbKeyReturn Then
        If txtWardID.Text = "" Then
            lblWardNm.Caption = ""
            txtWardID.SetFocus
            Exit Sub
        Else
            Dim strWard As String
            
            strWard = GetDeptNm(txtWardID.Text)
            
            If strWard = "" Then
                MsgBox "부서코드를 확인하세요.", vbInformation
                txtWardID.Text = ""
                lblWardNm.Caption = ""
                txtWardID.SetFocus
                Call txtWardId_KeyDown(vbKeyDown, 0)
            Else
                lblWardNm.Caption = strWard
                txtCorpCd.SetFocus
'                SendKeys "{TAB}"
            End If
            
            '** 원본 ====================================================
'            Dim Rs As Recordset
'
'            Set Rs = New Recordset
'
'            strWard = GetSQLWard(txtWardID.Text)
'
'            Rs.Open strWard, DBConn
'
'            If Rs.EOF = False Then
'                ObjSysInfo.BuildingCd = Rs.Fields("bldgb").Value & ""
'                ObjSysInfo.BuildingNm = Rs.Fields("bldnm").Value & ""
'                ObjSysInfo.BuildingNo = Rs.Fields("bldno").Value & ""
'                txtWardID.Tag = txtWardID.Text
'            Else
'                MsgBox "부서코드를 확인하세요.", vbInformation
'                txtWardID.Text = ""
'                lblWardNm.Caption = ""
'                txtWardID.SetFocus
'                Call txtWardId_KeyDown(vbKeyDown, 0)
'            End If
'            Set Rs = Nothing
            '============================================================

        End If
    End If
    Exit Sub

Err_Trap:
    Resume Next

End Sub

'% 검색한 처방을 테이블에 디스플레이 한다.
Private Sub DisplayOrder()
    Dim Rs          As Recordset
    Dim SqlStmt     As String
    Dim SvAccNo     As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim strBuildCd  As String
    
    Dim i           As Integer
    
    ' 처방내역 검색
    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(cboColTm.Text, CS_TimeDbFormat)

    SqlStmt = MySql.SqlBatchBarReprint_New(tmpDate, tmpTime, txtWardID.Text)

    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        MsgBox "해당 날짜/시간에 작업한 병동채취내역이 없습니다.", vbExclamation, "메세지"
        If cboColTm.ListCount <= 0 Then
            Call ClearRtn(0)
            txtWardID.SetFocus
        Else
            cboColTm.SetFocus
        End If
        GoTo NoData
    End If

    SvAccNo = ""

    With tblOrdSheet

        .ReDraw = False
        .MaxRows = 0
        .RowHeight(-1) = 12
        'Locking Cells
        .Row = -1
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
        
        .Row = 0
        
        For i = 1 To Rs.RecordCount
        
            If SvAccNo <> Trim(Rs.Fields("LabNo").Value) Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                SvAccNo = Trim(Rs.Fields("LabNo").Value)
            
                .Col = 1:  .Value = Trim("" & Rs.Fields("Seq").Value)
                .Col = 2:  .Value = "" & Rs.Fields("PtId").Value
                .Col = 3:  .Value = Trim("" & Rs.Fields("PtNm").Value)
                .Col = 4:  .Value = Format(Trim("" & Rs.Fields("OrdDt").Value), CS_DateLongMask) '처방일
                .Col = 6:  .Value = Trim("" & Rs.Fields("SpcNm").Value)      '검체명
                .Col = 7:  .Value = Choose(Val("" & Rs.Fields("StatFg").Value) + 1, "", "Y")  '응급여부
                .Col = 8:  .Value = Trim("" & Rs.Fields("LabNo").Value)      'LabNo
                .Col = 9:  .Value = Trim("" & Rs.Fields("OrdDt").Value)      '처방일
                .Col = 10: .Value = Trim("" & Rs.Fields("OrdNo").Value)      '처방번호
                .Col = 11: .Value = Trim("" & Rs.Fields("OrdSeq").Value)     '처방Seq
                .Col = 12: .Value = Trim("" & Rs.Fields("OrdCd").Value)      '검사코드
                .Col = 13: .Value = Trim("" & Rs.Fields("SpcNm").Value)      '검체명
                .Col = 14: .Value = Trim("" & Rs.Fields("WorkArea").Value)   'WorkArea
                .Col = 15: .Value = Trim("" & Rs.Fields("AccSeq").Value)     'AccSeq
                .Col = 16: .Value = Trim("" & Rs.Fields("StoreCd").Value)    '보관코드
                .Col = 17: .Value = Trim("" & Rs.Fields("AccDt").Value)      'AccDt  채혈일
                .Col = 18: .Value = Trim("" & Rs.Fields("StatFg").Value)     '응급여부
                .Col = 19: .Value = Trim("" & Rs.Fields("AbbrNm5").Value)    '약어명
                .Col = 20: .Value = Trim("" & Rs.Fields("SpcYy").Value) & Format(Val(Rs.Fields("SpcNo").Value), CS_BarFormat)     '검체번호
                .Col = 21:
                            strBuildCd = Trim("" & Rs.Fields("buildcd").Value)
'                            Dim objBld As clsBasisData
                            Dim strBld As String
'                            Set objBld = Nothing
'                            Set objBld = New clsBasisData
                            strBld = GetBuildNm(strBuildCd)
'                            Set objBld = Nothing
                            
'                            If ObjLISComCode.Building.Exists(strBuildCd) Then
'                                Call ObjLISComCode.Building.KeyChange(strBuildCd)
                                .Value = strBld 'ObjLISComCode.Building.Fields("BuildNm")   '건물명
'                            Else
'                                .Value = ""                                     '건물명
'                            End If
                .Col = 22:  .Value = Trim("" & Rs.Fields("wardId").Value)
                            If Trim("" & Rs.Fields("HosilId").Value) <> "" Then
                                .Value = .Value & "/" & Trim("" & Rs.Fields("HosilId").Value)    '병동코드
                            End If
                .Col = 23: .Value = Trim("" & Rs.Fields("DeptCd").Value)     '진료과코드
                .Col = 24: .Value = Trim("" & Rs.Fields("StsCd").Value)      'status
                .Col = 25: .Value = Mid(Trim("" & Rs.Fields("ReqTm").Value), 1, 2) & ":" & _
                                    Mid(Trim("" & Rs.Fields("ReqTm").Value), 3, 2)   '희망채혈일시
                .Col = 26: .Value = Trim("" & Rs.Fields("OrdDiv").Value)     '처방구분
                .Col = 27: .Value = Trim("" & Rs.Fields("FzFg").Value)       '냉동절편구분
            End If
            
            .Col = 1    'Seq
                If i = 1 Then txtFromSeq.Text = .Value
                If i = Rs.RecordCount Then txtToSeq.Text = .Value
            .Col = 5: .Value = .Value & Trim("" & Rs.Fields("AbbrNm5").Value) & "," '처방명
             Rs.MoveNext
        Next
        .ReDraw = True

    End With
    cmdReprint.Enabled = True
    CleanFg = False

NoData:
    Set Rs = Nothing
End Sub

Public Sub Call_WardId_KeyPress()

   Call txtWardId_KeyPress(vbKeyReturn)

End Sub


Public Sub Call_dtpToTime_Validate()

    Call dtpToTime_Validate(False)

End Sub

Public Sub Call_cmdGetOrders_click()

    Call cmdGetOrders_Click

End Sub
