VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrint 
   Caption         =   "frmPrint"
   ClientHeight    =   8100
   ClientLeft      =   10200
   ClientTop       =   6960
   ClientWidth     =   11265
   LinkTopic       =   "frmPrint"
   ScaleHeight     =   8100
   ScaleWidth      =   11265
   Begin FPSpread.vaSpread vasPRes 
      Height          =   2205
      Left            =   5100
      TabIndex        =   15
      Top             =   8460
      Visible         =   0   'False
      Width           =   7065
      _Version        =   393216
      _ExtentX        =   12462
      _ExtentY        =   3889
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
      SpreadDesigner  =   "frmPrint.frx":0000
   End
   Begin VB.CheckBox ChkAll 
      Height          =   255
      Left            =   660
      TabIndex        =   12
      Top             =   1020
      Width           =   165
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   150
      TabIndex        =   3
      Top             =   0
      Width           =   10905
      Begin VB.CommandButton Command1 
         Appearance      =   0  '평면
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9840
         TabIndex        =   14
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton txtResPrint 
         Appearance      =   0  '평면
         Caption         =   "결과 출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   8160
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   585
         Left            =   4770
         TabIndex        =   7
         Top             =   150
         Width           =   3195
         Begin VB.TextBox txtPE 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2310
            TabIndex        =   9
            Top             =   180
            Width           =   675
         End
         Begin VB.TextBox txtPS 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1170
            TabIndex        =   8
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label7 
            Caption         =   "~"
            Height          =   195
            Left            =   2040
            TabIndex        =   11
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "S.No"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdCall 
         Caption         =   "Local Data"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2940
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1290
         TabIndex        =   4
         Top             =   330
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67305473
         CurrentDate     =   39699
      End
      Begin VB.Label Label3 
         Caption         =   "접수일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   5
         Top             =   360
         Width           =   1035
      End
   End
   Begin FPSpread.vaSpread vasRes 
      Height          =   7065
      Left            =   5760
      TabIndex        =   0
      Top             =   930
      Width           =   5295
      _Version        =   393216
      _ExtentX        =   9340
      _ExtentY        =   12462
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   3
      Protect         =   0   'False
      SpreadDesigner  =   "frmPrint.frx":0267
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   7065
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5535
      _Version        =   393216
      _ExtentX        =   9763
      _ExtentY        =   12462
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   4
      Protect         =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmPrint.frx":3EEB
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   6915
      Left            =   11760
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   13215
      _Version        =   393216
      _ExtentX        =   23310
      _ExtentY        =   12197
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   30
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmPrint.frx":7C3A
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const colCheckBox = 1
Const colBARCODE = 2
Const colRack = 3
Const colPos = 4
Const colPID = 5
Const colPName = 6
Const colJumin = 7
Const colPSex = 8
Const colPAge = 9
Const colState = 10
Const colEXAMDATE = 11
Const colSlipNo1 = 12
Const colSlipNo2 = 13
Const colReqDate = 14

Const colEQUIPEXAM = 3
Const colExamCode = 4
Const colExamName = 5
Const colResult = 6
Const colRCheck = 7
Const colPCheck = 8
Const colDCheck = 9
Const colUnit = 10
Const colRef = 11
Const colPanic = 12
Const colResult1 = 13
Private Sub cmdCall_Click()
    If vasID.MaxRows > 0 Then vasID.MaxRows = 0
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    gstrQuy = "SELECT BARCODE, SEQNO, sendflag "
    gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EXAMDATE = '" & Format(DTPicker1.Value, "YYYYMMDD") & "' "
    gstrQuy = gstrQuy & vbCrLf & " GROUP BY BARCODE, SEQNO, sendflag "
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY SEQNO "
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If
    
    If Not ADR Is Nothing Then
        Do Until ADR.EOF
            vasID.MaxRows = vasID.MaxRows + 1: vasID.Row = vasID.MaxRows
            
            vasID.Col = 2: vasID.Text = Trim(ADR!barcode & "")
            vasID.Col = 3: vasID.Text = Trim(ADR!SEQNO & "")
            
            Select Case Trim(ADR!sendflag & "")
                Case "0"
                    vasID.Col = 4: vasID.Text = "W/S"
                    SetBackColor vasID, vasID.MaxRows, vasID.MaxRows, 1, 1, 255, 250, 205
                Case "1"
                    vasID.Col = 4: vasID.Text = "Result"
                    SetBackColor vasID, vasID.MaxRows, vasID.MaxRows, 1, 1, 255, 250, 205
                Case "2"
                    vasID.Col = 4: vasID.Text = "완료"
                    SetBackColor vasID, vasID.MaxRows, vasID.MaxRows, 1, 1, 202, 255, 112
            End Select
            
            ADR.MoveNext
        Loop
        ADR.Close: Set ADR = Nothing
    End If
    Call CloseDB

End Sub

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If ChkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            
'            If Trim(GetText(vasID, iRow, colState)) = "Result" And Trim(GetText(vasID, iRow, colBARCODE)) <> "" Then
                vasID.Row = iRow
                vasID.Col = 1
                vasID.Value = 1
'            End If
        Next iRow
    ElseIf ChkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub



Private Sub Command1_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    DTPicker1 = Format(Date, "yyyy-mm-dd")
End Sub

Private Sub txtPS_GotFocus()
    SELECTFocus txtPS
End Sub

Private Sub txtPS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtPS = "" Then
            txtPS.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtPS) = False Then
            txtPS.SetFocus
            Exit Sub
        End If
        
'        txtPS.Text = Format(Trim(txtPS.Text), "000#")
        
        txtPE.SetFocus
    End If
End Sub

Private Sub txtPE_GotFocus()
    SELECTFocus txtPE
End Sub

Private Sub txtPE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim lsBARCODE As String
    Dim lRow As Long
    
    
    If KeyCode = vbKeyReturn Then
        If txtPE = "" Then
            txtPE.SetFocus
            Exit Sub
        End If
        
        If IsNumeric(txtPS) = False Then
            txtPS.SetFocus
            Exit Sub
        End If
        
        If IsNumeric(txtPE) = False Then
            txtPE.SetFocus
            Exit Sub
        End If
'        txtPE.Text = Format(Trim(txtPE.Text), "000#")

        For i = CLng(txtPS) To CLng(txtPE)
            vasID.Col = 1
            vasID.Row = i
            vasID.Value = 1
               
        Next
        
        txtResPrint.SetFocus
    End If
End Sub

Private Sub txtResPrint_Click()
    Dim lngRow      As Long
    Dim lngLineCnt   As Long
    Dim strLine     As String
    Dim intGumCnt   As Integer
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub
    
''    Dim X As Printer
''    For Each X In Printers
''       If X.Orientation = vbPRORPortrait Then
''          ' 프린터를 시스템 기본값으로 설정합니다.
''          Set Printer = X
''          ' 프린터 찾기를 중지합니다.
''          Exit For
''       End If
''    Next

'    Set Printer = Printers
    
    Printer.Orientation = vbPRORLandscape
    
    lngLineCnt = 0
    
    GoSub PRN_HEADER
    
    For lngRow = 1 To vasID.DataRowCnt
        
        vasID.Col = 1
        vasID.Row = lngRow
        If vasID.Value = 1 Then
                    
            gstrQuy = "SELECT A.EXAMDATE, A.barcode, A.SEQNO, A.diskno, A.posno, A.ExamCode, B.ExamNAME, A.RESULT "
            gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES A, EQUIPEXAM B "
            gstrQuy = gstrQuy & vbCrLf & " WHERE A.examcode  = B.examcode "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPCODE = B.EQUIPCODE "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.EXAMDATE  = '" & Format(DTPicker1.Value, "YYYYMMDD") & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPNO   = '" & gtypREG_INFO.EQUIPCD & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.BARCODE   = '" & Trim(GET_CELL(vasID, 2, lngRow)) & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND A.SEQNO     = '" & Trim(GET_CELL(vasID, 3, lngRow)) & "' "
            gstrQuy = gstrQuy & vbCrLf & " ORDER BY B.SEQNO " '/검사순서
            If ReadSQL(gstrQuy, ADR) = False Then
                Call CloseDB
                Exit Sub
            End If
            
            If Not ADR Is Nothing Then
                lngLineCnt = lngLineCnt + 1
                If lngLineCnt > 60 Then
                    '''Printer.Print Printer.Page     ' 인쇄됩니다.
                    lngLineCnt = 1
                    Printer.NewPage
                    GoSub PRN_HEADER
                End If
                Printer.Print ""
                
                lngLineCnt = lngLineCnt + 1
                If lngLineCnt > 60 Then
                    '''Printer.Print Printer.Page     ' 인쇄됩니다.
                    lngLineCnt = 1
                    Printer.NewPage
                    GoSub PRN_HEADER
                End If
                
                Printer.Print Space(10) & "S.No " & TEXT_LSET(Trim(ADR!SEQNO & ""), 4) & Space(5) & Trim(ADR!diskno & "") & "-" & Trim(ADR!posno & "") & Space(10) & "S.ID " & Trim(ADR!barcode & "")
                
                Do Until ADR.EOF
                    intGumCnt = intGumCnt + 1
                    strLine = strLine & TEXT_LSET(Trim(ADR!examname & ""), 8) & TEXT_RSET(Trim(ADR!Result & ""), 8) & " / "
        
                    If intGumCnt = 10 Then
                        lngLineCnt = lngLineCnt + 1
                        
                        If lngLineCnt > 60 Then
                            '''Printer.Print Printer.Page     ' 인쇄됩니다.
                            Printer.NewPage
                            GoSub PRN_HEADER
                            lngLineCnt = 1
                        End If
                        
                        Printer.Print Space(10) & strLine
                        intGumCnt = 0
                        strLine = ""
                    End If
                    
                    ADR.MoveNext
                Loop
                
                If intGumCnt < 10 And intGumCnt > 0 Then
                    lngLineCnt = lngLineCnt + 1
                    
                    If lngLineCnt > 60 Then
                        '''Printer.Print Printer.Page     ' 인쇄됩니다.
                        Printer.NewPage
                        GoSub PRN_HEADER
                        lngLineCnt = 1
                    End If
                    
                    Printer.Print Space(10) & strLine
                    intGumCnt = 0
                    strLine = ""
                End If
                
                ADR.Close: Set ADR = Nothing
            End If
    
    
        End If
    Next lngRow
    
    Printer.EndDoc
    
    Call CloseDB
Exit Sub

PRN_HEADER:
    Printer.Print Space(10) & "검사일자: " & DTPicker1.Value & Space(10) & "Page: " & Printer.Page
    Printer.Print Space(10) & "-----------------------------------------------------------------------------------------------------------------------------"
Return
    
'''    Dim i As Long
'''    Dim j As Long
'''    Dim sResCnt As Integer
'''    Dim sBARCODE As String
'''    Dim sEQUIPCODE As String
'''    Dim sResult As String
'''    Dim iRow As Integer
'''
'''    Dim sCurDate As String
'''    Dim sSerDate As String
'''    Dim sHead As String
'''    Dim sFoot As String
'''
'''    ClearSpread vasPrint
'''    If IsNumeric(txtPS) = True And IsNumeric(txtPE) = True Then
'''    Else
'''        MsgBox "검사번호를 숫자로 입력하세요."
'''        Exit Sub
'''    End If
'''
'''    sResCnt = 1
'''
'''    For i = 1 To vasID.DataRowCnt
'''
'''        vasID.Col = 1
'''        vasID.Row = i
'''        If vasID.Value = 1 Then
'''
'''
'''            iRow = sResCnt
'''            sBARCODE = Trim(GetText(vasID, i, colBARCODE))
'''
'''            SetText vasPrint, sBARCODE, iRow, 1
'''            ClearSpread vasPRes
'''
'''            SQL = "SELECT EQUIPCODE, result from PAT_RES WHERE BARCODE = '" & sBARCODE & "'"
'''            res = db_SELECT_Vas(gLocal, SQL, vasPRes)
'''            For j = 1 To vasPRes.DataRowCnt
'''                sEQUIPCODE = Trim(GetText(vasPRes, j, 1))
'''                sResult = Trim(GetText(vasPRes, j, 2))
'''                SELECT Case sEQUIPCODE
'''                Case "!"
'''                    SetText vasPrint, sResult, iRow, 4
'''                Case "2"
'''                    SetText vasPrint, sResult, iRow, 6
'''                Case "3"
'''                    SetText vasPrint, sResult, iRow, 7
'''                Case "4"
'''                    SetText vasPrint, sResult, iRow, 8
'''                Case "5"
'''                    SetText vasPrint, sResult, iRow, 9
'''                Case "6"
'''                    SetText vasPrint, sResult, iRow, 10
'''                Case "7"
'''                    SetText vasPrint, sResult, iRow, 11
'''                Case "@"
'''                    SetText vasPrint, sResult, iRow, 5
'''                Case ")"
'''                    SetText vasPrint, sResult, iRow, 12
'''                Case "+"
'''                    SetText vasPrint, sResult, iRow, 15
'''                Case "-"
'''                    SetText vasPrint, sResult, iRow, 16
'''                Case "%"
'''                    SetText vasPrint, sResult, iRow, 14
'''                Case "#"
'''                    SetText vasPrint, sResult, iRow, 13
'''                End SELECT
'''
'''            Next
'''
'''            If sResCnt = 30 Then
'''
'''
'''
'''                sCurDate = Format(Date, "yyyy/mm/dd")
'''                sSerDate = Format(Date, "yyyy/mm/dd")
'''
'''                vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
'''                vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
'''                vasPrint.PrintJobName = "혈액학 결과 출력"
'''
'''                sHead = "/fn""굴림체"" /fz""13"" /fb1 /fi0 /fu0 " & "/l" & "  Result Review Chart" & "/n" & "/n" & "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "     " & Format(Time, "hh:mm:ss") & "     " & "Op. : abx" & "/n"
'''
'''                sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 울산산업보건센터 임상병리실" & "/n"
'''
'''                vasPrint.PrintHeader = sHead
'''                vasPrint.PrintFooter = sFoot
'''
'''                vasPrint.PrintMarginTop = 680
'''                vasPrint.PrintMarginBottom = 680
'''                vasPrint.PrintMarginLeft = 300
'''            '현재 SS가 비대칭으로 출력함
'''            '    vaslist.PrintMarginLeft = 720
'''    '            vasPrint.PrintMarginLeft = 0
'''                vasPrint.PrintMarginRight = 300
'''
'''                vasPrint.PrintColor = True
'''                vasPrint.PrintGrid = True
'''
'''            'Set printing range
'''                vasPrint.PrintType = 0  'SS_PRINT_ALL(default)
'''
'''                vasPrint.PrintShadows = True
'''
'''                vasPrint.Action = 13 'SS_ACTION_PRINT
'''
'''                sResCnt = 1
'''                ClearSpread vasPrint
'''            Else
'''                sResCnt = sResCnt + 1
'''            End If
'''        End If
'''    Next
'''
'''    If vasPrint.DataRowCnt > 0 Then
'''        sCurDate = Format(Date, "yyyy/mm/dd")
'''        sSerDate = Format(Date, "yyyy/mm/dd")
'''
'''        vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
'''        vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
'''        vasPrint.PrintJobName = "혈액학 결과 출력"
'''
'''        sHead = "/fn""굴림체"" /fz""13"" /fb1 /fi0 /fu0 " & "/l" & "  Result Review Chart" & "/n" & "/n" & "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "     " & Format(Time, "hh:mm:ss") & "     " & "Op. : abx" & "/n"
'''
'''        sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 울산산업보건센터 임상병리실"
'''
'''        vasPrint.PrintHeader = sHead
'''        vasPrint.PrintFooter = sFoot
'''
'''        vasPrint.PrintMarginTop = 680
'''        vasPrint.PrintMarginBottom = 680
'''        vasPrint.PrintMarginLeft = 300
'''    '현재 SS가 비대칭으로 출력함
'''    '    vaslist.PrintMarginLeft = 720
''''            vasPrint.PrintMarginLeft = 0
'''        vasPrint.PrintMarginRight = 300
'''
'''        vasPrint.PrintColor = True
'''        vasPrint.PrintGrid = True
'''
'''    'Set printing range
'''        vasPrint.PrintType = 0  'SS_PRINT_ALL(default)
'''
'''        vasPrint.PrintShadows = True
'''
'''        vasPrint.Action = 13 'SS_ACTION_PRINT
'''
'''        sResCnt = 1
'''        ClearSpread vasPrint
'''
'''    End If
'''    For i = 1 To vasID.DataRowCnt
'''        vasID.Col = 1
'''        vasID.Row = i
'''        vasID.Value = 0
'''
'''    Next
'''
'''
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    If vasRes.MaxRows > 0 Then vasRes.MaxRows = 0
    
    If Row < 1 Then Exit Sub

    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    gstrQuy = "SELECT A.ExamCode, B.ExamNAME, A.RESULT "
    gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES A, EQUIPEXAM B "
    gstrQuy = gstrQuy & vbCrLf & " WHERE A.examcode  = B.examcode "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPCODE = B.EQUIPCODE "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.EXAMDATE  = '" & Format(DTPicker1.Value, "YYYYMMDD") & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPNO   = '" & gtypREG_INFO.EQUIPCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.BARCODE   = '" & Trim(GET_CELL(vasID, 2, Row)) & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.SEQNO     = '" & Trim(GET_CELL(vasID, 3, Row)) & "' "
    gstrQuy = gstrQuy & vbCrLf & " ORDER BY B.SEQNO "
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If

    If Not ADR Is Nothing Then
        Do Until ADR.EOF
            vasRes.MaxRows = vasRes.MaxRows + 1: vasRes.Row = vasRes.MaxRows

            vasRes.Col = 1: vasRes.Text = Trim(ADR!ExamCode & "")
            vasRes.Col = 2: vasRes.Text = Trim(ADR!examname & "")
            vasRes.Col = 3: vasRes.Text = Trim(ADR!Result & "")

            ADR.MoveNext
        Loop
        ADR.Close: Set ADR = Nothing
    End If
    Call CloseDB
    
    
    
    
'''    Dim lsID As String
'''    Dim lsTmpID As String
'''
'''    Dim i As Integer
'''
'''    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
'''    If Row = 0 Then
'''        vasSort vasID, Col
'''    End If
'''
'''    If Row < 1 Or Row > vasID.DataRowCnt Then
'''        Exit Sub
'''    End If
'''
'''    lsID = Trim(GetText(vasID, Row, colBARCODE))
'''
'''    ClearSpread vasRes
'''    vasRes.MaxRows = 0
'''
'''    SQL = "SELECT '', a.BARCODE, a.EQUIPCODE,  a.examcode, a.examname, a.result, a.refflag, a.panicflag, a.deltaflag, a.unit, a.refvalue, a.panicvalue, a.result " & vbCrLf & _
'''          "FROM PAT_RES a, EQUIPEXAM b" & vbCrLf & _
'''          "WHERE a.EXAMDATE = '" & Format(DTPicker1, "yyyymmdd") & "' " & vbCrLf & _
'''          "  AND a.EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'''          "  AND a.BARCODE = '" & Trim(GetText(vasID, vasID.Row, colBARCODE)) & "' " & vbCrLf & _
'''          "  AND a.SEQNO = '" & Trim(GetText(vasID, vasID.Row, colRack)) & "' " & vbCrLf & _
'''          "  AND a.diskno = '" & Trim(GetText(vasID, vasID.Row, colPos)) & "' " & vbCrLf & _
'''          "  AND a.examcode = b.examcode and a.EQUIPCODE = b.EQUIPCODE " & vbCrLf & _
'''          "  ORDER BY b.SEQNO"
'''
'''    res = db_SELECT_Vas(gLocal, SQL, vasRes)
'''    If res = -1 Then
'''        SaveQuery SQL
'''        Exit Sub
'''    End If

End Sub

