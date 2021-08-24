VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmWorkList 
   BackColor       =   &H00BF8B59&
   Caption         =   "워크리스트"
   ClientHeight    =   9885
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   16725
   Icon            =   "frmWorkList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   16725
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   945
      Left            =   17370
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   3375
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
         Left            =   780
         TabIndex        =   12
         Text            =   "1"
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton cmdSeq 
         Caption         =   "Seq 매치"
         Height          =   375
         Left            =   1620
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblSeqNo 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "시작 Seq"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   150
         TabIndex        =   13
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   16725
      TabIndex        =   0
      Top             =   0
      Width           =   16725
      Begin VB.CheckBox chkSave 
         Appearance      =   0  '평면
         BackColor       =   &H00AE8B59&
         Caption         =   "저장 결과     포함 조회"
         ForeColor       =   &H00C0FFFF&
         Height          =   510
         Left            =   3480
         TabIndex        =   16
         Top             =   210
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   120
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
         Format          =   136904705
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   1290
         TabIndex        =   2
         Top             =   450
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
         Format          =   136904705
         CurrentDate     =   40457
      End
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   555
         Left            =   5010
         TabIndex        =   6
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "워크조회"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmWorkList.frx":06C2
         TransparentPicture=   "frmWorkList.frx":081C
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSend 
         Height          =   555
         Left            =   6480
         TabIndex        =   7
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "화면전송"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmWorkList.frx":320E
         TransparentPicture=   "frmWorkList.frx":3368
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSendClose 
         Height          =   555
         Left            =   7950
         TabIndex        =   8
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "전송/닫기"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmWorkList.frx":5D5A
         TransparentPicture=   "frmWorkList.frx":5EB4
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdWorkPrint 
         Height          =   555
         Left            =   9420
         TabIndex        =   9
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "워크출력"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmWorkList.frx":88A6
         TransparentPicture=   "frmWorkList.frx":8A00
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   555
         Left            =   10890
         TabIndex        =   14
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         Caption         =   "닫기"
         CaptionChecked  =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmWorkList.frx":B3F2
         TransparentPicture=   "frmWorkList.frx":B98C
         ButtonAttrib    =   2
         ButtonTransparent=   1
         ForeColor       =   16777215
         BackColor       =   16777215
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "까지"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   2760
         TabIndex        =   15
         Top             =   540
         Width           =   360
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   795
         Left            =   90
         Top             =   60
         Width           =   12405
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "조회기간 :"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   4
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "에서"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   2760
         TabIndex        =   3
         Top             =   210
         Width           =   360
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   8835
      Left            =   120
      TabIndex        =   5
      Top             =   930
      Width           =   16455
      _Version        =   393216
      _ExtentX        =   29025
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
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmWorkList.frx":E37E
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

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSearch_Click()
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork)
    
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
                For intORow = 1 To frmInterface.spdOrder.MaxRows
                    frmInterface.spdOrder.Row = intORow
                    frmInterface.spdOrder.Col = colBARCODE
                    If strBarno = GetText(frmInterface.spdOrder, intORow, colBARCODE) Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    frmInterface.spdOrder.MaxRows = frmInterface.spdOrder.MaxRows + 1
                    intRow = frmInterface.spdOrder.MaxRows
                    For i = colCHECKBOX To colSTATE
                        Call SetText(frmInterface.spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                
                        varItems = GetText(spdWork, intWRow, colITEMS)
                        varItems = Split(varItems, "/")
                        For intItems = 0 To UBound(varItems)
                            For intOCol = colSTATE + 1 To frmInterface.spdOrder.MaxCols
                                frmInterface.spdOrder.Row = 0
                                frmInterface.spdOrder.Col = intOCol
                                If varItems(intItems) = Trim(frmInterface.spdOrder.Text) Then
                                    .Row = frmInterface.spdOrder.MaxRows
                                    Call SetText(frmInterface.spdOrder, "◇", frmInterface.spdOrder.MaxRows, intOCol)
'                                    GoTo RST
                                End If
                            Next
                        Next
                    Next
                    
                    'Call SetText(frmInterface.spdOrder, GetText(spdWork, intWRow, colITEMS), intRow, colITEMS)
                    frmInterface.spdOrder.RowHeight(-1) = 15
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
                For intORow = 1 To frmInterface.spdOrder.MaxRows
                    If GetText(spdWork, intWRow, colSEQNO) = GetText(frmInterface.spdOrder, intORow, colSEQNO) Then
                        
                        Call SetText(frmInterface.spdOrder, GetText(spdWork, intWRow, colBARCODE), intORow, colBARCODE)
                        DoEvents
                        If GetSampleInfo(intORow, frmInterface.spdOrder) = -1 Then
                            'MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
                        Else
                            '정보수정
                            SQL = ""
                            SQL = SQL & "UPDATE PATRESULT SET "
                            SQL = SQL & "  BARCODE       = '" & Trim(GetText(frmInterface.spdOrder, intORow, colBARCODE)) & "'" & vbCr
                            SQL = SQL & " ,INOUT         = '" & Trim(GetText(frmInterface.spdOrder, intORow, colINOUT)) & "'" & vbCr
                            SQL = SQL & " ,CHARTNO       = '" & Trim(GetText(frmInterface.spdOrder, intORow, colCHARTNO)) & "'" & vbCr
                            SQL = SQL & " ,PID           = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPID)) & "'" & vbCr
                            SQL = SQL & " ,PNAME         = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPNAME)) & "'" & vbCr
                            SQL = SQL & " ,PSEX          = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPSEX)) & "'" & vbCr
                            SQL = SQL & " ,PAGE          = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPAGE)) & "'" & vbCr
''                            SQL = SQL & " ,PJUMIN        = '" & Trim(GetText(frmInterface.spdOrder, intORow, colPJUMIN)) & "'" & vbCr
'                            SQL = SQL & " ,PANICVALUE    = '" & Trim(GetText(frmInterface.spdOrder, intORow, colKEY1)) & "'" & vbCr
                            SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(frmInterface.spdOrder, intORow, colEXAMDATE)) & "'" & vbCr
                            SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(frmInterface.spdOrder, intORow, colSAVESEQ)) & vbCr
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

Private Sub cmdWorkPrint_Click()
    
    If spdWork.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        spdWork.PrintOrientation = PrintOrientationPortrait
        spdWork.Action = 13
    End If
    

End Sub

Private Sub Form_Load()
    Dim intCol      As Integer
    Dim intColWidth As Integer
    
    Call CtlInitializing

    '-- 컬럼보이기설정
    Call SetColumnView(spdWork)
    
    spdWork.ColWidth(spdWork.MaxCols) = 30
    
'''    intColWidth = 0
'''    With spdWork
'''        'For intCol = 1 To colSTATE
'''        For intCol = 1 To .MaxCols
'''            intColWidth = intColWidth + .ColWidth(intCol)
'''        Next
'''        spdWork.ColWidth(colITEMS) = (((intColWidth * Screen.TwipsPerPixelX) - (spdWork.WIDTH / Screen.TwipsPerPixelX)) / Screen.TwipsPerPixelX) - 7
'''
'''    End With
        
'pixels = twips / Screen.TwipsPerPixelX

'twips = pixels * Screen.TwipsPerPixelX
        
'
'To get width of an object in Pixel from Twips:
'
'widthinpixels = object.WIDTH / Screen.TwipsPerPixelX
'
'To get height of an object in Pixel from Twips:
'
'heigthinpixels = object.HEIGHT / Screen.TwipsPerPixelY
'
'To get width of an object in Twips from Pixel:
'
'widthinpixels = object.WIDTH * Screen.TwipsPerPixelX
'
'To get height of an object in Twips from Pixel:
'
'heigthinpixels = object.HEIGHT * Screen.TwipsPerPixelY
'    spdWork.MaxRows = 10
    
    
'    Dim i As Integer
'
'    For i = 1 To 10
'        Call SetText(spdWork, i, i, colBARCODE)
'        Call SetText(spdWork, i * 10, i, colITEMS)
'
'    Next
    '-- 검사명 보이기
'    Call SetExamCode(spdWork)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub CtlInitializing()
    spdWork.MaxRows = 0
    
    dtpFrom.Value = Now
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
'    Dim intColWidth As Integer
    
    If Me.ScaleHeight = 0 Then Exit Sub

    spdWork.WIDTH = Me.ScaleWidth - 300
    spdWork.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 300
    
'    intColWidth = 0
'    With spdWork
'        For intCol = 1 To .MaxCols
'            intColWidth = intColWidth + .ColWidth(intCol)
'        Next
'        spdWork.ColWidth(colITEMS) = (((intColWidth * Screen.TwipsPerPixelX) - (spdWork.WIDTH / Screen.TwipsPerPixelX)) / Screen.TwipsPerPixelX) - 7
'
'    End With

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
    
    With frmInterface.spdOrder
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
            frmInterface.spdOrder.MaxRows = frmInterface.spdOrder.MaxRows + 1
            intRow = frmInterface.spdOrder.MaxRows
            
            For i = colCHECKBOX To colSTATE
                Call SetText(frmInterface.spdOrder, GetText(spdWork, intWRow, i), intRow, i)

                varItems = GetText(spdWork, intWRow, colITEMS)
                varItems = Split(varItems, "/")
                For intItems = 0 To UBound(varItems)
                    For intOCol = colSTATE + 1 To frmInterface.spdOrder.MaxCols
                        frmInterface.spdOrder.Row = 0
                        frmInterface.spdOrder.Col = intOCol
                        If varItems(intItems) = Trim(frmInterface.spdOrder.Text) Then
                            .Row = intRow
                            Call SetText(frmInterface.spdOrder, "◇", intRow, intOCol)
                        End If
                    Next
                Next
            Next
            
            frmInterface.spdOrder.RowHeight(-1) = 15
        End If
    
    End With
    
End Sub

Private Sub spdWork_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
'    Dim intRow      As Integer
'    Dim strSeq      As String
    
    
    sRow = spdWork.ActiveRow
    sCol = colPNAME
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdWork, sRow, sCol)
    
    If KeyCode = vbKeyDelete Then
        
        If MsgBox(strNewBarNo & " 를 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdWork, sRow, sRow
        spdWork.MaxRows = spdWork.MaxRows - 1
    
    End If

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

