VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "워크리스트"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   16470
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2925
      Left            =   8340
      TabIndex        =   10
      Top             =   3780
      Visible         =   0   'False
      Width           =   7185
      Begin VB.TextBox txtSeqNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1140
         TabIndex        =   13
         Text            =   "1"
         Top             =   960
         Width           =   645
      End
      Begin VB.CommandButton cmdSeq 
         Caption         =   "Seq 매치"
         Height          =   375
         Left            =   3690
         TabIndex        =   12
         Top             =   990
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox chkSave 
         Appearance      =   0  '평면
         BackColor       =   &H00800000&
         Caption         =   "저장포함"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   2280
         TabIndex        =   11
         Top             =   1110
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblSeqNo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   480
         TabIndex        =   14
         Top             =   1020
         Width           =   510
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   28560
      TabIndex        =   0
      Top             =   0
      Width           =   28560
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "화면정리"
         Height          =   375
         Left            =   11550
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   180
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00800000&
         Height          =   555
         Left            =   4650
         TabIndex        =   15
         Top             =   30
         Width           =   3345
         Begin VB.OptionButton optTest 
            Appearance      =   0  '평면
            BackColor       =   &H00800000&
            Caption         =   "COVID-19"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   120
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton optTest 
            Appearance      =   0  '평면
            BackColor       =   &H00800000&
            Caption         =   "RP+PB"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Index           =   3
            Left            =   2280
            TabIndex        =   19
            Top             =   120
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.OptionButton optTest 
            Appearance      =   0  '평면
            BackColor       =   &H00800000&
            Caption         =   "PB"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Index           =   2
            Left            =   1590
            TabIndex        =   18
            Top             =   120
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.OptionButton optTest 
            Appearance      =   0  '평면
            BackColor       =   &H00FF80FF&
            Caption         =   "RP"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Index           =   1
            Left            =   870
            TabIndex        =   17
            Top             =   120
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.CommandButton cmdSendClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "전송/닫기"
         Height          =   375
         Left            =   10410
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "닫기"
         Height          =   375
         Left            =   12690
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "화면전송"
         Height          =   375
         Left            =   9270
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "워크조회"
         Height          =   375
         Left            =   8130
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1410
         TabIndex        =   1
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   136052737
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   3090
         TabIndex        =   2
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   136052737
         CurrentDate     =   40457
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   465
         Left            =   8070
         Top             =   120
         Width           =   5805
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회기간 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   270
         Width           =   930
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   465
         Left            =   270
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   2880
         TabIndex        =   3
         Top             =   270
         Width           =   150
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   8835
      Left            =   30
      TabIndex        =   9
      Top             =   750
      Width           =   16395
      _Version        =   393216
      _ExtentX        =   28919
      _ExtentY        =   15584
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   23
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmWorkList.frx":06C2
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    
    Call CtlInitializing
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSearch_Click()
    
'    Call XmlSelect_Free

    If optTest(0).Value = True Then
        gTest = "MTB"
    ElseIf optTest(1).Value = True Then
        gTest = "RP"
    ElseIf optTest(2).Value = True Then
        gTest = "PB"
    ElseIf optTest(3).Value = True Then
        gTest = "RPPB"
    End If
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork)
    
'    Call SetSpreadSort(spdWork, 0)
    
'    Call SetSpreadSort(spdWork, 0)
    
''    SortKeys = Array(1, 3)
''    SortKeyOrder = Array(1, 1)
''    ' Sort data in first five columns and rows by column 1 and 3
''    spdWork.Sort Col, 1, 5, 5, SS_SORT_BY_ROW, SortKeys, SortKeyOrder
    
    
    spdWork.SortKey(1) = colCHARTNO
    spdWork.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
    spdWork.Sort -1, -1, -1, -1, SortByRow
    
    spdWork.RowHeight(-1) = 15

End Sub

Private Sub cmdSend_Click()
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                blnSame = False
                strBarno = GetText(spdWork, intWRow, colBARCODE)
                For intORow = 1 To frmMain.spdOrder.MaxRows
                    frmMain.spdOrder.Row = intORow
                    frmMain.spdOrder.Col = colBARCODE
                    If strBarno = GetText(frmMain.spdOrder, intORow, colBARCODE) Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
                    intRow = frmMain.spdOrder.MaxRows
                    For i = colCHECKBOX To colSTATE
                        Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                
                        varItems = GetText(spdWork, intWRow, colITEMS)
                        varItems = Split(varItems, "|")
                        For intItems = 0 To UBound(varItems)
                            For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                                frmMain.spdOrder.Row = 0
                                frmMain.spdOrder.Col = intOCol
                                If Trim(varItems(intItems)) = Trim(frmMain.spdOrder.Text) Then
                                    .Row = frmMain.spdOrder.MaxRows
                                    Call SetText(frmMain.spdOrder, "◇", frmMain.spdOrder.MaxRows, intOCol)
'                                    GoTo RST
                                End If
                            Next
                        Next
                    Next
                    
                    frmMain.spdOrder.RowHeight(-1) = 15
                End If
            End If
        Next
    End With
    
End Sub

Private Sub cmdSendClose_Click()
    
    Call cmdSend_Click
    
    Call cmdClose_Click
    
End Sub

Private Sub cmdSeq_Click()
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim strSeq          As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                For intORow = 1 To frmMain.spdOrder.MaxRows
                    If GetText(spdWork, intWRow, colSEQNO) = GetText(frmMain.spdOrder, intORow, colSEQNO) Then
                        
                        Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), intORow, colBARCODE)
                        DoEvents
                        If GetSampleInfo(intORow, frmMain.spdOrder) = -1 Then
                            'MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                        Else
                            '정보수정
                            SQL = ""
                            SQL = SQL & "UPDATE PATRESULT SET "
                            SQL = SQL & "  BARCODE       = '" & Trim(GetText(frmMain.spdOrder, intORow, colBARCODE)) & "'" & vbCr
                            SQL = SQL & " ,INOUT         = '" & Trim(GetText(frmMain.spdOrder, intORow, colINOUT)) & "'" & vbCr
                            SQL = SQL & " ,CHARTNO       = '" & Trim(GetText(frmMain.spdOrder, intORow, colCHARTNO)) & "'" & vbCr
                            SQL = SQL & " ,PID           = '" & Trim(GetText(frmMain.spdOrder, intORow, colPID)) & "'" & vbCr
                            SQL = SQL & " ,PNAME         = '" & Trim(GetText(frmMain.spdOrder, intORow, colPNAME)) & "'" & vbCr
                            SQL = SQL & " ,PSEX          = '" & Trim(GetText(frmMain.spdOrder, intORow, colPSEX)) & "'" & vbCr
                            SQL = SQL & " ,PAGE          = '" & Trim(GetText(frmMain.spdOrder, intORow, colPAGE)) & "'" & vbCr
''                            SQL = SQL & " ,PJUMIN        = '" & Trim(GetText(frmMain.spdOrder, intORow, colPJUMIN)) & "'" & vbCr
'                            SQL = SQL & " ,PANICVALUE    = '" & Trim(GetText(frmMain.spdOrder, intORow, colKEY1)) & "'" & vbCr
                            SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(frmMain.spdOrder, intORow, colEXAMDATE)) & "'" & vbCr
                            SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(frmMain.spdOrder, intORow, colSAVESEQ)) & vbCr
                            SQL = SQL & "   AND EQUIPNO  = '" & gHOSP.MACHCD & "' & vbCr"
                            
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- 성공
                            End If
                        End If
                        Exit For
                    End If
                Next intORow
            End If
        Next intWRow
    End With
End Sub

Private Sub Form_Load()
    
    Call CtlInitializing

    '-- 컬럼보이기설정
'    Call SetColumnView
    
    '-- 검사명 보이기
'    Call SetExamCode(spdResult)

End Sub


Private Sub SetColumnView()
    Dim i       As Integer
    Dim varSize As Variant
    
    varSize = Split(gCOLSIZE, "|")
    
    For i = 0 To UBound(varSize) - 1
        '워크리스트
        If i >= 2 Then
            spdWork.Col = i + 2
            If Mid(gCOLVIEW, i + 1, 1) = 1 Then
                spdWork.ColHidden = False
            Else
                spdWork.ColHidden = True
            End If
            spdWork.ColWidth(i + 2) = varSize(i)
        End If
    Next

End Sub

Private Sub CtlInitializing()
    
    spdWork.MaxRows = 0
    
    dtpFrom.Value = Now '- 6
    dtpTo.Value = Now
    
    txtSeqNo.Text = "1"
    
    '순번사용
    If gHOSP.RSTTYPE = "1" Then
        lblSeqNo.Visible = True
        txtSeqNo.Visible = True
    Else
        lblSeqNo.Visible = False
        txtSeqNo.Visible = False
    End If
    
End Sub


Private Sub Form_Resize()
    
    If Me.ScaleHeight = 0 Then Exit Sub

    spdWork.Width = Me.ScaleWidth - 300
    spdWork.Height = Me.ScaleHeight - picHeader.Height - 300

    spdWork.ColWidth(colSTATE + 1) = 50 '(spdWork.Width / 40) * intColSum

End Sub

Private Sub optTest_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 3
        optTest(i).BackColor = &H800000
    Next
    
    optTest(Index).BackColor = &HFF80FF
    
    frmMain.optTest(Index).Value = True

End Sub

Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer

    If Row = 0 And Col <> colCHECKBOX Then
        Call SetSpreadSort(spdWork, 0)
        Exit Sub
    End If
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdWork, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "1", i, colCHECKBOX)
            Next
        End If
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
            Call SetText(spdWork, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdWork, "1", Row, colCHECKBOX)
        End If
    End If
    
'    txtQuery.Visible = True
'    txtQuery.Text = GetText(spdWork, Row, colITEMS)
    
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    Dim strBarno_Work   As String
    
    If Row = 0 Then Exit Sub
    If Col <> colBARCODE Then
        Exit Sub
    End If
    
    intWRow = Row
    spdWork.Row = Row
    spdWork.Col = colBARCODE
    strBarno_Work = Trim(spdWork.Text)
    
    With frmMain.spdOrder
        blnSame = False
        For intORow = 1 To .MaxRows
            .Row = intORow
            .Col = colBARCODE
            If strBarno_Work = Trim(.Text) Then
                blnSame = True
                Exit For
            End If
        Next
        
        If blnSame = False Then
            frmMain.spdOrder.MaxRows = frmMain.spdOrder.MaxRows + 1
            intRow = frmMain.spdOrder.MaxRows
            
            For i = colCHECKBOX To colSTATE
                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, i), intRow, i)

'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSPECNO), intRow, colSPECNO)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHECKBOX), intRow, colCHECKBOX)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colHOSPDATE), intRow, colHOSPDATE)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colBARCODE), intRow, colBARCODE)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colSEQNO), intRow, colSEQNO)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colCHARTNO), intRow, colCHARTNO)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPID), intRow, colPID)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colINOUT), intRow, colINOUT)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPNAME), intRow, colPNAME)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPSEX), intRow, colPSEX)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPAGE), intRow, colPAGE)
'    '            Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colPJUMIN), introw, colPJUMIN)
'                Call SetText(frmMain.spdOrder, GetText(spdWork, intWRow, colOCNT), intRow, colOCNT)
                
                varItems = GetText(spdWork, intWRow, colITEMS)
                varItems = Split(varItems, "/")
                For intItems = 0 To UBound(varItems)
                    For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                        frmMain.spdOrder.Row = 0
                        frmMain.spdOrder.Col = intOCol
                        If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
                            .Row = intRow
                            Call SetText(frmMain.spdOrder, "◇", intRow, intOCol)
                        End If
                    Next
                Next
            Next
            
            frmMain.spdOrder.RowHeight(-1) = 15
        End If
    
    End With
    
End Sub

Private Sub spdWork_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    Dim strSeq As String
    
    If KeyAscii = vbKeyReturn Then
        With spdWork
            If .ActiveCol = colSEQNO Then
                strSeq = GetText(spdWork, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "숫자만 입력이 가능합니다"
                    Exit Sub
                End If
                For intRow = .ActiveRow + 1 To .MaxRows
                    Call SetText(spdWork, strSeq + 1, intRow, colSEQNO)
                Next
            End If
        End With
    End If
    
End Sub
