VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "워크리스트 조회"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14490
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   14490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   900
      TabIndex        =   11
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
      TabIndex        =   9
      Text            =   "0001"
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
      TabIndex        =   8
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownClose 
      Caption         =   "Down >> Close"
      Enabled         =   0   'False
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
      TabIndex        =   7
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
      TabIndex        =   2
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownLoad 
      Caption         =   "Down"
      Enabled         =   0   'False
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
      TabIndex        =   1
      Top             =   150
      Width           =   1395
   End
   Begin FPSpread.vaSpread vasWorkList1 
      Height          =   6195
      Left            =   270
      TabIndex        =   0
      Top             =   3090
      Width           =   14055
      _Version        =   393216
      _ExtentX        =   24791
      _ExtentY        =   10927
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      ColsFrozen      =   1
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
      MaxCols         =   10
      MaxRows         =   20
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":000C
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
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
      Format          =   21299201
      CurrentDate     =   40457
   End
   Begin MSComCtl2.DTPicker dtpStopDt 
      Height          =   315
      Left            =   3030
      TabIndex        =   4
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
      Format          =   21299201
      CurrentDate     =   40457
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   7995
      Left            =   240
      TabIndex        =   12
      Top             =   630
      Width           =   13515
      _Version        =   393216
      _ExtentX        =   23839
      _ExtentY        =   14102
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
      SpreadDesigner  =   "frmWorkList.frx":0AC9
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
      TabIndex        =   10
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
      TabIndex        =   6
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
      TabIndex        =   5
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
'        For intRow = 1 To .DataRowCnt
'            .Row = intRow
'            .Col = colCheckBox
'            If .Value = 1 Then
'                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
'                intVasRow = frmInterface.vasID.MaxRows
'
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colSpecNo), intVasRow, colSpecNo)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colCheckBox), intVasRow, colCheckBox)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colHOSPDATE), intVasRow, colHOSPDATE)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colGubun), intVasRow, colGubun)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colBARCODE), intVasRow, colBARCODE)
'                'Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colRack), intVasRow, colRack)
'                'Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colPos), intVasRow, colPos)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colPID - 2), intVasRow, colPID)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colPNAME - 2), intVasRow, colPNAME)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colSex - 2), intVasRow, colSex)
'                Call SetText(frmInterface.vasID, GetText(vasWorkList, intVasRow, colAge - 2), intVasRow, colAge)
'
'                frmInterface.txtNum = frmInterface.txtNum + 1
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
''                Call SetText(frmInterface.vasID, Format(txtPos.Text, "0000"), i, 0)
''                txtPos.Text = Format(txtPos.Text + 1, "0000")
''            End If
''        Next
''    End If
'End Sub
'
'Private Sub cmdSearch_Click()
'
'    Call GetWorkList_JWINFO(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'
'End Sub
'
'Private Sub GetWorkList_JWINFO(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
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
'        vasID.MaxRows = 0
'        intRow = 0
'    End If
'
'    blnSame = False
'    vasID.ReDraw = False
'
'
'          SQL = "SELECT DISTINCT RECEIPTDATE as 접수일자, SPECIMENNUM as 바코드번호, RECEIPTNO as 챠트번호, IPDOPD, PTNO as 내원번호, SNAME as 이름, LABCODE as ITEM,ORDERCODE"
'    SQL = SQL & vbCrLf & "  FROM SLA_LabMaster "
'    SQL = SQL & vbCrLf & " WHERE RECEIPTDATE between '" & Format(pFrDt, "####-##-##") & "' and '" & Format(pToDt, "####-##-##") & "'"
'    SQL = SQL & vbCrLf & "   AND LABCODE IN (" & gAllExam & ") "
'    SQL = SQL & vbCrLf & "   AND JSTATUS < '3'" & vbLf
'    SQL = SQL & "  ORDER BY RECEIPTDATE "
'
'' Select distinct M.IpdOpd, M.ReceiptDate ADT, M.PTno PID, M.SName PNM, M.Age,
''        M.ReceiptTime , M.Sex, M.BI, M.DeptCode
''        ,M.WardCode, M.Roomcode, M.BillFlag, M.JStatus,  M.SPECIMENNUM
'' From SLA_LabMaster M, SLA_LABRESULT R
'' Where M.ReceiptNo = R.ReceiptNo
''   And M.ORDERCODE = R.ORDERCODE
''   And M.ReceiptDate Between '2012-04-24' And '2012-04-24'
''   And R.LABCODE in('C2200','C2210','C3730','C3750','C3720','C3780','C3711','C2411','B2570','B2580','B2710','B2602','B2590','C3790','C3794','CRP','C4903')
''   And M.JsTATUS < '3'
'' order by ADT, ReceiptTime, PID
'
'    Call SetSQLData("워크조회", SQL)
'
'    '-- Record Count 가져옴
'    cn_Ser.CursorLocation = adUseClient
'    Set RS = cn_Ser.Execute(SQL, , 1)
'    If Not RS.EOF = True And Not RS.BOF = True Then
'        frmProgress.Show
'        frmProgress.ZOrder 0
'        frmProgress.Xprog.Min = 1
'        frmProgress.Xprog.Max = RS.RecordCount + 1
'
'        Do Until RS.EOF
'            iCnt = iCnt + 1
'            With vasID
'                .ReDraw = False
'                For i = 1 To .DataRowCnt
'                    strDate = GetText(vasID, i, colHOSPDATE)
'                    strBarcode = GetText(vasID, i, colBARCODE)
'                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strBarcode Then
'                        blnSame = True
'                    End If
'
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
'                Next
'
'                If blnSame = False Then
'                    .MaxRows = .MaxRows + 1
'                    SetText vasID, "1", .MaxRows, colCheckBox
'                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
'                    'If Trim(RS.Fields("바코드번호")) & "" = "0" Then
'                    '    SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colBARCODE
'                    'Else
'                        SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
'                    'End If
'                    SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colCHARTNO
'                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colPID
'                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
'                    SetText vasID, IIf(Trim(RS.Fields("IPDOPD")) = 1, "입원", "외래"), .MaxRows, colINOUT
'                    SetText vasID, Trim(RS.Fields("ORDERCODE")) & "", .MaxRows, colPSEX
'
'
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
'
'                End If
'
'                blnSame = False
'            End With
'            '-- 프로그레스바 진행
'            frmProgress.Xprog.Value = iCnt
'            DoEvents
'
'            RS.MoveNext
'        Loop
'        chkWAll.Value = "1"
'    Else
'        StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
'        chkWAll.Value = "0"
'    End If
'
'    RS.Close
'
'    '-- 프로그레스바 닫기
'    Unload frmProgress
'
'    vasID.RowHeight(-1) = 12
'    vasID.ReDraw = True
'
'End Sub
'
'Private Sub Form_Load()
'
'    dtpStartDt.Value = Now
'    dtpStopDt.Value = Now
'    txtSeq.Text = "0001"
'
'    vasWorkList.MaxRows = 0
'
'End Sub
'
'
'
'Private Sub txtPos_KeyPress(KeyAscii As Integer)
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
''                Call SetText(frmInterface.vasID, Format(txtPos.Text, "0000"), i, 0)
''                txtPos.Text = Format(txtPos.Text + 1, "0000")
''            End If
''        Next
''    End If
'
'End Sub
'
'
'Private Sub txtSeq_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        txtSeq.Text = Format(txtSeq.Text, "0000")
'    End If
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
''        Call frmInterface.vasID.SetText(colSpecNo, pGrid_Point, "1")
''        Call frmInterface.vasID.SetText(1, pGrid_Point, "1")
''        Call frmInterface.vasID.SetText(4, pGrid_Point, .Text)
'
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colSpecNo), pGrid_Point, colSpecNo)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colCheckBox), pGrid_Point, colCheckBox)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colHOSPDATE), pGrid_Point, colHOSPDATE)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colGubun), pGrid_Point, colGubun)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colBARCODE), pGrid_Point, colBARCODE)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPID - 2), pGrid_Point, colPID)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPNAME - 2), pGrid_Point, colPNAME)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colSex - 2), pGrid_Point, colSex)
'        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colAge - 2), pGrid_Point, colAge)
'
''        .Row = Row: .Col = 5
''        Call vasID.SetText(5, pGrid_Point, .Text)
''        .Row = Row: .Col = 6
''        Call vasID.SetText(6, pGrid_Point, .Text)
''        .Row = Row: .Col = 7
''        Call vasID.SetText(7, pGrid_Point, .Text)
''        .Row = Row: .Col = 8
''        Call vasID.SetText(8, pGrid_Point, .Text)
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
''''            SetText frmInterface.vasID, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
''''            SetText frmInterface.vasID, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
''''            SetText frmInterface.vasID, Trim(gReadBuf(1)), pGrid_Point, colPName  '6
''''            SetText frmInterface.vasID, Trim(gReadBuf(2)), pGrid_Point, colSex    '7
''''            SetText frmInterface.vasID, Trim(gReadBuf(3)), pGrid_Point, colAge    '8
''''            SetText frmInterface.vasID, Format(Trim(gReadBuf(8)), "####-##-##"), pGrid_Point, 2
''''
''''            SetText frmInterface.vasID, Trim(gReadBuf(4)), pGrid_Point, 12
''''            SetText frmInterface.vasID, Trim(gReadBuf(5)), pGrid_Point, 13
''''            SetText frmInterface.vasID, Trim(gReadBuf(6)), pGrid_Point, 14
''''            SetText frmInterface.vasID, Trim(gReadBuf(7)), pGrid_Point, 15
''''            SetText frmInterface.vasID, Trim(gReadBuf(8)), pGrid_Point, 16
''''            SetText frmInterface.vasID, Trim(gReadBuf(9)), pGrid_Point, 17
''''            SetText frmInterface.vasID, Trim(gReadBuf(10)), pGrid_Point, 18
''''            SetText frmInterface.vasID, Trim(gReadBuf(11)), pGrid_Point, 19
''''            frmInterface.vasID.RowHeight(-1) = 12
''''        End If
'
'    End With
'End Sub
