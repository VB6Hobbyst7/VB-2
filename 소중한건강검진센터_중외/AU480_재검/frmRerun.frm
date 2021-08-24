VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmRerun 
   BackColor       =   &H00FFFFFF&
   Caption         =   "검사재검"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14100
   Icon            =   "frmRerun.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00400040&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   14100
      TabIndex        =   1
      Top             =   0
      Width           =   14100
      Begin VB.CheckBox chkAll 
         Appearance      =   0  '평면
         BackColor       =   &H00400040&
         Caption         =   "전체기간"
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3420
         TabIndex        =   12
         Top             =   150
         Width           =   675
      End
      Begin VB.TextBox txtBarNum 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1110
         TabIndex        =   0
         Text            =   "123456789012345"
         Top             =   150
         Width           =   2175
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "선택제외"
         Height          =   375
         Left            =   10020
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "화면정리"
         Height          =   375
         Left            =   11160
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "결과조회"
         Height          =   375
         Left            =   8880
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "닫기"
         Height          =   375
         Left            =   12300
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   150
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   15930
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   5490
         TabIndex        =   6
         Top             =   150
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
         Format          =   156368897
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   7170
         TabIndex        =   7
         Top             =   150
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
         Format          =   156368897
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "바코드:"
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
         Index           =   2
         Left            =   330
         TabIndex        =   11
         Top             =   240
         Width           =   660
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   465
         Left            =   150
         Top             =   90
         Width           =   4005
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
         Left            =   6960
         TabIndex        =   9
         Top             =   240
         Width           =   150
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
         Left            =   4380
         TabIndex        =   8
         Top             =   240
         Width           =   930
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   465
         Left            =   4260
         Top             =   90
         Width           =   4425
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   465
         Left            =   8760
         Top             =   90
         Width           =   4815
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   8835
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   21195
      _Version        =   393216
      _ExtentX        =   37386
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
      MaxCols         =   22
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmRerun.frx":058A
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmRerun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnCheck    As Boolean

Private Sub cmdClear_Click()
    
    spdResult.MaxRows = 0
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()
    Dim lRow As Long
    
    With spdResult
        For lRow = .DataRowCnt To 1 Step -1
            .Row = lRow
            .Col = 1
            If .Value = 1 Then
                Call DeleteRow(spdResult, lRow, lRow)
                spdResult.MaxRows = spdResult.MaxRows - 1
            End If
        Next lRow
    End With
    
    
End Sub

Private Sub cmdExcel_Click()

'    Call spdResult.ExportToExcel(App.PATH & "\" & Format(Now, "yyyy-mm-dd") & "_결과대장.xls", "결과대장", "Log.Text")
    
    Dim sFileName As String
            
On Error GoTo ErrHandler

    If spdResult.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        With CommonDialog1
            .CancelError = True
            .Flags = cdlOFNHideReadOnly
            .InitDir = App.PATH
            .Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
            .Filename = App.PATH & "\" & Format(Now, "yyyy-mm-dd") & "_결과대장.xls"
            .ShowSave
            sFileName = CommonDialog1.Filename
            SaveExcel sFileName, spdResult
            MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
        End With
    End If

Exit Sub
  
ErrHandler:
      
    ' 사용자가 [취소] 단추를 눌렀습니다.
    Exit Sub

End Sub

Private Sub cmdSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    If MsgBox("선택한 결과를 재전송하시겠습니까?", vbYesNo + vbCritical, "결과전송") = vbYes Then
        With spdResult
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = 1
                If .Value = 1 Then
                    
                    Res = SaveTransData(lRow, spdResult)
                
                    If Res = -1 Then
                        SetForeColor spdResult, lRow, lRow, 1, colSTATE, 255, 0, 0
                        SetText spdResult, "저장실패", lRow, colSTATE
                              
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '1' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdResult, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                    Else
                        SetBackColor spdResult, lRow, lRow, 1, colSTATE, 202, 255, 112
                        SetText spdResult, "저장완료", lRow, colSTATE
                        
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '2' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdResult, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- 성공
                        End If
                        
                    End If
                    spdResult.Row = lRow
                    spdResult.Col = 1
                    spdResult.Value = 0
                End If
            Next lRow
        End With
    End If
    
End Sub

Private Sub cmdSearch_Click()

    'spdResult.MaxRows = 0
    
    Call GetResultList(Format(dtpFrom.Value, "yyyy-mm-dd"), Format(dtpTo.Value, "yyyy-mm-dd"), spdResult)

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
    
    With spdResult
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                blnSame = False
                strBarno = GetText(spdResult, intWRow, colBARCODE)
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
                        Call SetText(frmMain.spdOrder, GetText(spdResult, intWRow, i), intRow, i)
                
                        varItems = GetText(spdResult, intWRow, colITEMS)
                        varItems = Split(varItems, "/")
                        For intItems = 0 To UBound(varItems)
                            For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                                frmMain.spdOrder.Row = 0
                                frmMain.spdOrder.Col = intOCol
                                If varItems(intItems) = Trim(frmMain.spdOrder.Text) Then
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

Private Sub Form_Load()
    
    dtpFrom.Value = Now - 1
    dtpTo.Value = Now

    blnCheck = True
    
    gRerunMode = True
    
    txtBarNum.Text = ""
    
    spdResult.MaxRows = 0
    spdResult.UserColAction = UserColActionSort
    
    '-- 컬럼보이기설정
    Call SetColumnView(spdResult)
    
    '-- 검사명 보이기
    Call SetExamCodeRerun(spdResult)
    
End Sub

Private Sub SetExamCodeRerun(ByVal SPD As Object)


    Dim i As Integer
    
    With SPD
        .MaxCols = colSTATE + UBound(gArrEQP)
        For i = 0 To UBound(gArrEQP) - 1
            .Col = colSTATE + (i + 1)
            .Row = -1
            'CellType = CellTypeStaticText
            '.TypeHAlign = TypeHAlignCenter
            '.TypeVAlign = TypeVAlignCenter

            .CellType = CellTypeCheckBox
            .TypeCheckCenter = True
            
            '.TypeCheckPicture(0) = LoadPicture(App.PATH & "\ICON\Uncheck.bmp")
            '.TypeCheckPicture(1) = LoadPicture(App.PATH & "\ICON\check.bmp")

            .TypeCheckPicture(0) = LoadPicture(App.PATH & "\ICON\task_uncheck.bmp")
            .TypeCheckPicture(1) = LoadPicture(App.PATH & "\ICON\task_check.bmp")

            
            Call SetText(SPD, Trim(gArrEQP(i + 1, 6)), 0, colSTATE + (i + 1))    '-- 5 : 약어명
            .ColWidth(colSTATE + (i + 1)) = 5 'gCOLWIDTH
        Next
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    frmMain.lblHospInfo(2).Caption = "-- 일반모드 --"
    frmMain.shpStatus.BackColor = &HACFFEF
    gRerunMode = False

End Sub

Private Sub Form_Resize()

    If Me.ScaleHeight = 0 Then Exit Sub

    spdResult.WIDTH = Me.ScaleWidth - 300
    spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 300
    
End Sub

Private Sub spdResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim iRow As Long
    
    If Row = 0 Then
        If Col = colCHECKBOX Then
            With spdResult
                If blnCheck = False Then
                    For iRow = 1 To .DataRowCnt
                        .Row = iRow
                        .Col = 1
                        
                        .Value = 1
                    Next iRow
                    blnCheck = True
                Else
                    For iRow = 1 To .DataRowCnt
                        .Row = iRow
                        .Col = 1
                        
                        .Value = 0
                    Next iRow
                    blnCheck = False
                End If
            End With
        'ElseIf Col = colPNAME Or Col = colPID Then
            '-- 정렬 추가
        '    Call SetSpreadSort(spdResult, Col)

        ElseIf Col > colSTATE Then
            With spdResult
                If GetText(spdResult, 1, Col) = "1" Then
                    For iRow = 1 To .DataRowCnt
                        .Row = iRow
                        .Col = Col
                        .Value = 0
                    Next iRow
                Else
                    For iRow = 1 To .DataRowCnt
                        .Row = iRow
                        .Col = Col
                        .Value = 1
                    Next iRow
                End If
            End With
        End If
    End If


'    If Row > 0 And Col = colCHECKBOX Then
'        If GetText(spdResult, Row, colCHECKBOX) = "1" Then
'            Call SetText(spdResult, "0", Row, colCHECKBOX)
'        Else
'            Call SetText(spdResult, "1", Row, colCHECKBOX)
'        End If
'    End If


End Sub

Private Sub txtBarNum_GotFocus()
        
    txtBarNum.SelStart = 0
    txtBarNum.SelLength = Len(txtBarNum.Text)

End Sub

Private Sub txtBarNum_KeyDown(KeyCode As Integer, Shift As Integer)

    
    If Trim(txtBarNum.Text) <> "" And KeyCode = vbKeyReturn Then
        Call GetResultList_Barcode(Format(dtpFrom.Value, "yyyy-mm-dd"), Format(dtpTo.Value, "yyyy-mm-dd"), spdResult, Trim(txtBarNum.Text), chkAll.Value)
    
        txtBarNum.SelStart = 0
        txtBarNum.SelLength = Len(txtBarNum.Text)
    
    End If
    
    
End Sub
