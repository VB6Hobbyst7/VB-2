VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmResult 
   BackColor       =   &H00BF8B59&
   Caption         =   "결과조회"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15705
   Icon            =   "frmResult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   15705
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   15705
      TabIndex        =   0
      Top             =   0
      Width           =   15705
      Begin VB.OptionButton optPrtOri 
         BackColor       =   &H00AE8B59&
         Caption         =   "가로"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   9240
         TabIndex        =   15
         Top             =   480
         Width           =   675
      End
      Begin VB.OptionButton optPrtOri 
         BackColor       =   &H00AE8B59&
         Caption         =   "세로"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   9240
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   675
      End
      Begin HSCotrol.CButton cmdSave 
         Height          =   495
         Left            =   4860
         TabIndex        =   8
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   15698777
         Caption         =   " 선택저장"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmResult.frx":030A
      End
      Begin HSCotrol.CButton cmdClose 
         Height          =   495
         Left            =   12720
         TabIndex        =   9
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   15698777
         Caption         =   " 닫    기"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmResult.frx":0464
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   14910
         Top             =   150
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1290
         TabIndex        =   2
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
         Format          =   131268609
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   1290
         TabIndex        =   3
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
         Format          =   131268609
         CurrentDate     =   40457
      End
      Begin HSCotrol.CButton cmdSearch 
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   15698777
         Caption         =   " 결과조회"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmResult.frx":09FE
      End
      Begin HSCotrol.CButton cmdExcel 
         Height          =   495
         Left            =   7620
         TabIndex        =   11
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   15698777
         Caption         =   " 엑셀저장"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmResult.frx":0B58
      End
      Begin HSCotrol.CButton cmdRsltPrint 
         Height          =   495
         Left            =   9960
         TabIndex        =   12
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   15698777
         Caption         =   " 미리보기"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmResult.frx":0CB2
      End
      Begin HSCotrol.CButton cmdDelete 
         Height          =   495
         Left            =   6240
         TabIndex        =   13
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   15698777
         Caption         =   " 결과삭제"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmResult.frx":0E0C
      End
      Begin HSCotrol.CButton cmdClear 
         Height          =   495
         Left            =   11340
         TabIndex        =   7
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BackColor       =   15698777
         Caption         =   " 화면정리"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   65535
         HoverPicture    =   "frmResult.frx":0F66
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
         TabIndex        =   6
         Top             =   210
         Width           =   360
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
         TabIndex        =   5
         Top             =   210
         Width           =   840
      End
      Begin VB.Shape shpTopMenu 
         BorderColor     =   &H00FFFFFF&
         Height          =   795
         Left            =   90
         Top             =   60
         Width           =   15465
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
         TabIndex        =   4
         Top             =   540
         Width           =   360
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   8835
      Left            =   90
      TabIndex        =   1
      Top             =   900
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
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmResult.frx":10C0
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
End
Attribute VB_Name = "frmResult"
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
    
    If MsgBox("선택한 결과를 삭제하시겠습니까?", vbYesNo + vbCritical, "결과전송") = vbYes Then
        With spdResult
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = 1
                If .Value = 1 Then
                          SQL = " DELETE From PATRESULT " & vbCrLf
                    SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                    SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(spdResult, lRow, colSAVESEQ))
                    SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdResult, lRow, colBARCODE)) & "' "
                        
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- 성공
                    End If
                    
                    spdResult.Row = lRow
                    spdResult.Col = 1
                    spdResult.Value = 0
                End If
            Next lRow
        End With
        
        Call cmdSearch_Click
        
    End If
    
End Sub

Private Sub cmdExcel_Click()
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

Private Sub cmdRsltPrint_Click()
    
    If spdResult.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    Else
        If optPrtOri(0).Value = True Then
            spdResult.PrintOrientation = PrintOrientationPortrait       '세로출력
        Else
            spdResult.PrintOrientation = PrintOrientationLandscape      '가로출력
        End If
'        spdResult.Action = 13
        
        frmPreview.spdResultPreview.hWndSpread = spdResult.hWnd
        frmPreview.Show vbModal
                
    End If
    
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

    spdResult.MaxRows = 0
    
    Call GetResultList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdResult)

End Sub

Private Sub Form_Load()
    
    dtpFrom.Value = Now
    dtpTo.Value = Now

    blnCheck = True
    
    spdResult.MaxRows = 0
    
    '-- 컬럼보이기설정
    Call SetColumnView(spdResult)
    
    '-- 검사명 보이기
    Call SetExamCode(spdResult)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Resize()

    If Me.ScaleHeight = 0 Then Exit Sub

    spdResult.WIDTH = Me.ScaleWidth - 300
    spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 300
    
End Sub

Private Sub spdResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim iRow As Long
    
    If Row = 0 And Col = colCHECKBOX Then
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
    End If


    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdResult, Row, colCHECKBOX) = "1" Then
            Call SetText(spdResult, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdResult, "1", Row, colCHECKBOX)
        End If
    End If


End Sub

Private Sub spdResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    
    sRow = spdResult.ActiveRow
    sCol = colPNAME
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdResult, sRow, sCol)
    
    If KeyCode = vbKeyDelete Then
        
        If MsgBox(strNewBarNo & " 를 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdResult, sRow, sRow
        spdResult.MaxRows = spdResult.MaxRows - 1
    
    End If

End Sub
