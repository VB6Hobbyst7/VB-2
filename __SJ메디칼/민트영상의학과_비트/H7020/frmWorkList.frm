VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "워크리스트 조회"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17205
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   17205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   900
      TabIndex        =   10
      Top             =   810
      Width           =   225
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   8
      Text            =   "1"
      Top             =   210
      Width           =   1125
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   7
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownClose 
      Caption         =   "Down >> Close"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "조회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   180
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownLoad 
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   150
      Width           =   1395
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   315
      Left            =   1290
      TabIndex        =   2
      Top             =   180
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   63569921
      CurrentDate     =   40457
   End
   Begin MSComCtl2.DTPicker dtpStopDt 
      Height          =   315
      Left            =   3030
      TabIndex        =   3
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   63569921
      CurrentDate     =   40457
   End
   Begin FPSpread.vaSpread vasWorkList 
      Height          =   5865
      Left            =   300
      TabIndex        =   11
      Top             =   720
      Width           =   16695
      _Version        =   393216
      _ExtentX        =   29448
      _ExtentY        =   10345
      _StockProps     =   64
      ButtonDrawMode  =   4
      ColHeaderDisplay=   0
      ColsFrozen      =   16
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   17
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmWorkList.frx":000C
   End
   Begin VB.Label Label2 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6360
      TabIndex        =   9
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "조회일자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   5
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label7 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2820
      TabIndex        =   4
      Top             =   270
      Width           =   105
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub chkAll_Click()
'    Dim iRow As Long
'
'    If chkAll.Value = 1 Then
'        For iRow = 1 To vasWorkList.DataRowCnt
'            vasWorkList.Row = iRow
'            vasWorkList.Col = 1
'
'            vasWorkList.Value = 1
'        Next iRow
'    ElseIf chkAll.Value = 0 Then
'        For iRow = 1 To vasWorkList.DataRowCnt
'            vasWorkList.Row = iRow
'            vasWorkList.Col = 1
'
'            vasWorkList.Value = 0
'        Next iRow
'    End If
'
'End Sub
'
'Private Sub cmdClose_Click()
'
'    Unload Me
'
'End Sub
'
'Private Sub cmdDownClose_Click()
'
'    Call cmdDownLoad_Click
'
'    Call cmdClose_Click
'
'End Sub
'
'Private Sub cmdDownLoad_Click()
'    Dim intVasRow As Integer
'    Dim intRow As Integer
'    Dim j  As Integer
'
'    j = 0
'    With vasWorkList
'        For intRow = 1 To .MaxRows
'            .Row = intRow
'            .Col = colCheckBox
'            If .Value = 1 Then
'                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
'                DoEvents
'                intVasRow = frmInterface.vasID.MaxRows
'
'                If GetText(vasWorkList, intRow, colBARCODE) = "" Then
'                    frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows - 1
'                    Exit Sub
'                End If
'
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colSpecNo), intVasRow, colSpecNo)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colCheckBox), intVasRow, colCheckBox)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colHOSPDATE), intVasRow, colHOSPDATE)
''                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colGubun), intVasRow, colGubun)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colBARCODE), intVasRow, colBARCODE)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colCHARTNO), intVasRow, colCHARTNO)
'                'Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colRack), intVasRow, colRack)
'                'Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPos), intVasRow, colPos)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPID), intVasRow, colPID)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPNAME), intVasRow, colPNAME)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPSEX), intVasRow, colPSEX)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPAGE), intVasRow, colPAGE)
'
'               ' frmInterface.txtNum = frmInterface.txtNum + 1
'
'                .Col = 1
'                .Value = "0"
'            End If
'        Next
'        frmInterface.vasID.RowHeight(-1) = 12
'    End With
'
'
'
''    Dim i As Integer
''
''    If KeyAscii = vbKeyReturn Then
''        For i = 1 To vasWorkList.MaxRows
''            vasWorkList.Row = i
''            vasWorkList.Col = 1
''            If vasWorkList.Value = "1" Then
''                If Trim(txtPos.Text) = "" Then
''                    txtPos.Text = "1"
''                End If
''                Call SetText(frmInterface.vasworklist, Format(txtPos.Text, "0000"), i, 0)
''                txtPos.Text = Format(txtPos.Text + 1, "0000")
''            End If
''        Next
''    End If
'End Sub
'
'Private Sub cmdSearch_Click()
'
'    Call GetWorkList_MCC(Format(dtpStartDt, "yyyymmdd"), Format(dtpStopDt, "yyyymmdd"))
'
'End Sub
'
''Private Sub GetWorkList_BIT(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
''    Dim RS          As ADODB.Recordset
''    Dim i           As Integer
''    Dim iCnt        As Long
''    Dim intRow      As Long
''    Dim intCol      As Integer
''    Dim strDate     As String
''    Dim strChart    As String
''    Dim blnSame     As Boolean
''
''    If pBarNo = "" Then
''        vasID.MaxRows = 0
''        intRow = 0
''    End If
''
''    blnSame = False
''    vasID.ReDraw = False
''
''    SQL = ""
''    SQL = SQL & "SELECT L.LABSERIAL, L.LABATTEND as 내원번호, L.LABCHTNUM as 챠트번호, L.LABODRDTE as 접수일자, M.MANADMFOR as 입외," & vbCrLf
''    SQL = SQL & "       M.MANRESNUM as 주민번호, M.MANPATNAM as 이름, L.LABINSNUM as 처방순서,L.LABSMPNAM as 검체명, L.LABBARCOD as 바코드번호, L.LABODRCOD as ITEM, L.LABODRSTP as SEQ " & vbCrLf
''    SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D, ME_MAN M" & vbCrLf
''    SQL = SQL & " WHERE L.LABODRDTE between  '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf
''    SQL = SQL & "   AND L.LABKEYNUM = D.DATKEYNUM " & vbCrLf                    '-- 테이블연결키값
''    SQL = SQL & "   AND L.LABATTEND = D.DATATTEND " & vbCrLf                    '-- 내원번호
''    SQL = SQL & "   AND L.LABATTEND = M.MANATTEND " & vbCrLf                    '-- 내원번호
''    SQL = SQL & "   AND L.LABCHTNUM = D.DATCHTNUM " & vbCrLf                    '-- 챠트번호
''    SQL = SQL & "   AND L.LABCHTNUM = M.MANCHTNUM " & vbCrLf                    '-- 챠트번호
''    SQL = SQL & "   AND L.LABODRDTE = D.DATODRDTE " & vbCrLf                    '-- 처방일자
''    SQL = SQL & "   AND L.LABODRCOD IN (" & gAllExam & ")" & vbCrLf
'''    SQL = SQL & "   AND L.LABSUBYON = 'Y' " & vbCrLf                           '-- 서브코드여부 (결과입력용 서브코드이면 Y)
''    SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL) " & vbCrLf    '-- 취소여부
''
''    '-- 저장미포함
''    If chkSaveAll.Value = "0" Then
''        SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)" & vbCrLf
''        SQL = SQL & "   AND L.LABENDDEP < '3' " & vbCrLf                        '-- 처리상태 (2:접수, 3:결과입력)
''        'SQL = SQL & "   AND D.DATENDDEP < '3' " & vbCrLf                        '-- 지원부서처리여부    CHAR(2)   1:대기, 2:결과입력대기, 3:완료, 9.예약
''    ElseIf chkSaveAll.Value = "1" Then
''        SQL = SQL & "   AND L.LABENDDEP <= '3' " & vbCrLf                       '-- 처리상태 (2:접수, 3:결과입력)
''        'SQL = SQL & "   AND D.DATENDDEP <= '3' " & vbCrLf                       '-- 지원부서처리여부    CHAR(2)   1:대기, 2:결과입력대기, 3:완료, 9.예약
''    End If
''    SQL = SQL & " ORDER BY L.LABODRDTE, L.LABBARCOD, L.LABINSNUM, L.LABODRCOD "
''
''    Call SetSQLData("워크조회", SQL)
''
''    '-- Record Count 가져옴
''    cn_Ser.CursorLocation = adUseClient
''    Set RS = cn_Ser.Execute(SQL, , 1)
''    If Not RS.EOF = True And Not RS.BOF = True Then
''        frmProgress.Show
''        frmProgress.ZOrder 0
''        frmProgress.Xprog.Min = 1
''        frmProgress.Xprog.Max = RS.RecordCount + 1
''
''        Do Until RS.EOF
''            iCnt = iCnt + 1
''            With vasID
''                .ReDraw = False
''                For i = 1 To .DataRowCnt
''                    strDate = GetText(vasID, i, colHOSPDATE)
''                    strChart = GetText(vasID, i, colBARCODE)
''                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strChart Then
''                        blnSame = True
''                    End If
''                    For intCol = colState + 1 To vasID.MaxCols
''                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
''                            vasID.Row = .MaxRows
''                            vasID.Col = intCol
''                            vasID.BackColor = vbYellow
''                            Exit For
''                        End If
''                    Next
''                Next
''                If blnSame = False Then
''                    .MaxRows = .MaxRows + 1
''                    SetText vasID, "1", .MaxRows, colCheckBox
''                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
''                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
''                    SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colCHARTNO
''                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
''                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
''                    SetText vasID, Trim(RS.Fields("검체명")) & "", .MaxRows, colSPCNM
''                    SetText vasID, Trim(RS.Fields("SEQ")) & "", .MaxRows, colPAGE
''
''                    Select Case Trim(Trim(RS.Fields("입외")) & "")
''                        Case "A":   SetText vasID, "외래", .MaxRows, colINOUT
''                        Case "F":   SetText vasID, "입원", .MaxRows, colINOUT
''                        Case Else:  SetText vasID, "", .MaxRows, colINOUT
''                    End Select
''
''                    If optBW(1).Value = True Then
''                        SetText vasID, txtRack.Text, .MaxRows, colDISKNO
''                        SetText vasID, txtPos.Text, .MaxRows, colPOSNO
''
''                        txtPos.Text = Format(Val(txtPos.Text) + 1, "00")
''                        If txtPos.Text = "11" Then
''                            txtPos.Text = "01"
''                            txtRack.Text = Format(Val(txtRack.Text) + 1, "0000")
''                        End If
''                    End If
''
''                    For intCol = colState + 1 To vasID.MaxCols
''                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
''                            vasID.Row = .MaxRows
''                            vasID.Col = intCol
''                            vasID.BackColor = vbYellow
''                            Exit For
''                        End If
''                    Next
''                End If
''                blnSame = False
''            End With
''            '-- 프로그레스바 진행
''            frmProgress.Xprog.Value = iCnt
''            DoEvents
''
''            RS.MoveNext
''        Loop
''        chkWAll.Value = "1"
''    Else
''        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
''        chkWAll.Value = "0"
''    End If
''
''    RS.Close
''    '-- 프로그레스바 닫기
''    Unload frmProgress
''
''    vasID.RowHeight(-1) = 12
''    vasID.ReDraw = True
''
''End Sub
'
'Private Sub GetWorkList_MCC(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
'    Dim RS          As ADODB.Recordset
'    Dim i           As Integer
'    Dim iCnt        As Long
'    Dim intRow      As Long
'    Dim intCol      As Integer
'    Dim strDate     As String
'    Dim strChart    As String
'    Dim strBarcode    As String
'    Dim blnSame     As Boolean
'
'    If pBarNo = "" Then
'        vasWorkList.MaxRows = 0
'        intRow = 0
'    End If
'
'    blnSame = False
'    vasWorkList.ReDraw = False
'
'
''          SQL = "SELECT DISTINCT ORD_YMD as 접수일자, BCODE_NO as 바코드번호, RECEPT_NO as 챠트번호, PTNT_NO as 내원번호,PTNT_NM as 이름, AGE as 나이,SEX as 성별,ORD_CD as ITEM"
'          SQL = "SELECT DISTINCT ORD_YMD, BCODE_NO, RECEPT_NO, PTNT_NO,PTNT_NM,AGE,SEX,ORD_CD" & vbCr
'    SQL = SQL & "  FROM MCCSI.H7LIS_BCODE_ORD " & vbCr
'    SQL = SQL & " WHERE ORD_YMD between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
'    SQL = SQL & "   AND ORD_CD IN (" & gAllExam & ") " & vbCr
'    SQL = SQL & "   AND RESULT_TYPE = '20'" & vbLf & vbCr
'    SQL = SQL & "  ORDER BY ORD_YMD,RECEPT_NO,BCODE_NO "
'
'    Call SetSQLData("워크조회", SQL)
'
'    '-- Record Count 가져옴
'    cn_Ser.CursorLocation = adUseClient
'    Set RS = cn_Ser.Execute(SQL, , 1)
'    If Not RS.EOF = True And Not RS.BOF = True Then
'
'        Do Until RS.EOF
'            iCnt = iCnt + 1
'            With vasWorkList
'                .ReDraw = False
'                For i = 1 To .DataRowCnt
'                    strDate = GetText(vasWorkList, i, colHOSPDATE)
'                    strBarcode = GetText(vasWorkList, i, colBARCODE)
'                    If Trim(RS("ORD_YMD")) = strDate And Trim(RS("BCODE_NO")) = strBarcode Then
'                        blnSame = True
'                    End If
'
'                    For intCol = colState + 1 To vasWorkList.MaxCols
'                        If Trim(RS.Fields("ORD_CD")) = gArrEquip(intCol - colState, 3) Then
'                            vasWorkList.Row = .MaxRows
'                            vasWorkList.Col = intCol
'                            vasWorkList.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
'                Next
'                If blnSame = False Then
'                    .MaxRows = .MaxRows + 1
'
'                    SetText vasWorkList, txtSeq.Text, .MaxRows, colSpecNo
'                    SetText vasWorkList, "1", .MaxRows, colCheckBox
'                    SetText vasWorkList, Trim(RS.Fields("ORD_YMD")) & "", .MaxRows, colHOSPDATE
'                    SetText vasWorkList, Trim(RS.Fields("BCODE_NO")) & "", .MaxRows, colBARCODE
'                    SetText vasWorkList, Trim(RS.Fields("RECEPT_NO")) & "", .MaxRows, colCHARTNO
'                    SetText vasWorkList, Trim(RS.Fields("PTNT_NO")) & "", .MaxRows, colPID
'                    SetText vasWorkList, Trim(RS.Fields("PTNT_NM")) & "", .MaxRows, colPNAME
'                    SetText vasWorkList, Trim(RS.Fields("AGE")) & "", .MaxRows, colPAGE
'                    SetText vasWorkList, Trim(RS.Fields("SEX")) & "", .MaxRows, colPSEX
'
'                    txtSeq.Text = txtSeq.Text + 1
'
'                    For intCol = colState + 1 To vasWorkList.MaxCols
'                        If Trim(RS.Fields("ORD_CD")) = gArrEquip(intCol - colState, 3) Then
'                            vasWorkList.Row = .MaxRows
'                            vasWorkList.Col = intCol
'                            vasWorkList.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
'                End If
'                blnSame = False
'            End With
'
'            RS.MoveNext
'        Loop
'        chkAll.Value = "1"
'    Else
'        'StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
'        chkAll.Value = "0"
'    End If
'
'    RS.Close
'
'    '-- 프로그레스바 닫기
'    Unload frmProgress
'
'    vasWorkList.RowHeight(-1) = 12
'    vasWorkList.ReDraw = True
'
'End Sub
'
'
'Private Sub Form_Load()
'
'    dtpStartDt.Value = Now
'    dtpStopDt.Value = Now
'    txtSeq.Text = "1"
'
'    vasWorkList.MaxRows = 0
'
'End Sub
'
'
'
'Private Sub txtSeq_KeyPress(KeyAscii As Integer)
'    Dim intRow As Integer
'
'    If KeyAscii = vbKeyReturn Then
'
'        For intRow = vasWorkList.ActiveRow To vasWorkList.MaxRows
'            Call SetText(vasWorkList, Val(txtSeq.Text), intRow, colSpecNo)
'            txtSeq.Text = txtSeq.Text + 1
'        Next
'
'        txtSeq.Text = Format(txtSeq.Text, "0000")
'
'    End If
'
'
'End Sub
'
'Private Sub vasWorkList_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim pGrid_Point As Integer
'    Dim sBarcode As String
'    Dim sChartNo As String
'
'    If Row = 0 Then Exit Sub
'
'    With vasWorkList
'        '.Col = Col
'        '.Row = Row
'        '.Col = colBarcode
'        pGrid_Point = SeqSearch(frmInterface.vasID, GetText(vasWorkList, Row, colBARCODE), colBARCODE)
'
'        If pGrid_Point = 0 Then
'            pGrid_Point = SeqNullSearch(frmInterface.vasID, Trim(.Text), colBARCODE)
'            If pGrid_Point = 0 Then
'                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
'                pGrid_Point = frmInterface.vasID.MaxRows
'            End If
'            .RowHeight(-1) = 12
'        End If
'
''        .Row = Row: .Col = colBarcode
''        sBarcode = Trim(.Text)
'
'
''        Call frmInterface.vasworklist.SetText(colSpecNo, pGrid_Point, "1")
''        Call frmInterface.vasworklist.SetText(1, pGrid_Point, "1")
''        Call frmInterface.vasworklist.SetText(4, pGrid_Point, .Text)
'
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colSpecNo), pGrid_Point, colSpecNo)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colCheckBox), pGrid_Point, colCheckBox)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colHOSPDATE), pGrid_Point, colHOSPDATE)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colBARCODE), pGrid_Point, colBARCODE)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPID), pGrid_Point, colPID)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colCHARTNO), pGrid_Point, colCHARTNO)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPNAME), pGrid_Point, colPNAME)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPSEX), pGrid_Point, colPSEX)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPAGE), pGrid_Point, colPAGE)
'
'
'
''        .Row = Row: .Col = 5
''        Call vasworklist.SetText(5, pGrid_Point, .Text)
''        .Row = Row: .Col = 6
''        Call vasworklist.SetText(6, pGrid_Point, .Text)
''        .Row = Row: .Col = 7
''        Call vasworklist.SetText(7, pGrid_Point, .Text)
''        .Row = Row: .Col = 8
''        Call vasworklist.SetText(8, pGrid_Point, .Text)
'        frmInterface.vasID.RowHeight(-1) = 12
'
''''        '바코드번호로 환자정보 불러오기
''''              SQL = "SELECT DiSTINCT CHARTNO, PATNAME, PATSEX, PATAGE,COMPANY,HOSPCODE,PATJUMIN,PATNO,COMMDATE,EXAMNO,EXAMID,IOFLAG  "
''''        SQL = SQL & vbCrLf & "  FROM PAT_RES "
''''        SQL = SQL & vbCrLf & " WHERE EXAMTYPE = '" & gPart & "' "
''''        SQL = SQL & vbCrLf & "   AND BARCODE = '" & sBarcode & "'"
''''
''''
''''        Res = GetDBSelectColumn(gLocal, SQL)
''''
''''        If Res = 1 Then
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(1)), pGrid_Point, colPName  '6
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(2)), pGrid_Point, colSex    '7
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(3)), pGrid_Point, colAge    '8
''''            SetText frmInterface.vasworklist, Format(Trim(gReadBuf(8)), "####-##-##"), pGrid_Point, 2
''''
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(4)), pGrid_Point, 12
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(5)), pGrid_Point, 13
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(6)), pGrid_Point, 14
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(7)), pGrid_Point, 15
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(8)), pGrid_Point, 16
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(9)), pGrid_Point, 17
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(10)), pGrid_Point, 18
''''            SetText frmInterface.vasworklist, Trim(gReadBuf(11)), pGrid_Point, 19
''''            frmInterface.vasworklist.RowHeight(-1) = 12
''''        End If
'
'    End With
'End Sub
