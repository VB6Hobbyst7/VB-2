VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmResult 
   BackColor       =   &H00FFFFFF&
   Caption         =   "결과조회"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14880
   Icon            =   "frmResult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   14880
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14880
      TabIndex        =   0
      Top             =   0
      Width           =   14880
      Begin VB.OptionButton optPrtOri 
         BackColor       =   &H00F8E4D8&
         Caption         =   "가로"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   8280
         TabIndex        =   12
         Top             =   300
         Width           =   675
      End
      Begin VB.OptionButton optPrtOri 
         BackColor       =   &H00F8E4D8&
         Caption         =   "세로"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   8280
         TabIndex        =   11
         Top             =   90
         Value           =   -1  'True
         Width           =   675
      End
      Begin HSCotrol.CButton cmdSave 
         Height          =   375
         Left            =   4890
         TabIndex        =   5
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "선택저장"
         ForeColor       =   16744576
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmResult.frx":030A
      End
      Begin HSCotrol.CButton cmdClose 
         Height          =   375
         Left            =   11250
         TabIndex        =   6
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "닫기"
         ForeColor       =   0
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmResult.frx":0464
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12540
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
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
         Format          =   145948673
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2310
         TabIndex        =   2
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
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
         Format          =   145883137
         CurrentDate     =   40457
      End
      Begin HSCotrol.CButton cmdSearch 
         Height          =   375
         Left            =   3780
         TabIndex        =   7
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   15698777
         Caption         =   "결과조회"
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmResult.frx":09FE
      End
      Begin HSCotrol.CButton cmdExcel 
         Height          =   375
         Left            =   7110
         TabIndex        =   8
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "엑셀저장"
         ForeColor       =   16744576
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmResult.frx":0B58
      End
      Begin HSCotrol.CButton cmdRsltPrint 
         Height          =   375
         Left            =   9030
         TabIndex        =   9
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "미리보기"
         ForeColor       =   16744576
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmResult.frx":0CB2
      End
      Begin HSCotrol.CButton cmdDelete 
         Height          =   375
         Left            =   6000
         TabIndex        =   10
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "결과삭제"
         ForeColor       =   16744576
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmResult.frx":0E0C
      End
      Begin HSCotrol.CButton cmdClear 
         Height          =   375
         Left            =   10140
         TabIndex        =   4
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   16777215
         Caption         =   "화면정리"
         ForeColor       =   16744576
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
         BorderStyle     =   1
         BorderColor     =   0
         HoverColor      =   4210752
         HoverPicture    =   "frmResult.frx":0F66
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "~"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   2130
         TabIndex        =   13
         Top             =   180
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "기간"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   180
         Width           =   360
      End
   End
   Begin FPSpreadADO.fpSpread spdResult 
      CausesValidation=   0   'False
      Height          =   8895
      Left            =   150
      TabIndex        =   14
      Tag             =   "20001"
      Top             =   720
      Width           =   16020
      _Version        =   524288
      _ExtentX        =   28258
      _ExtentY        =   15690
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   22
      MaxRows         =   489
      Protect         =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmResult.frx":10C0
      VisibleCols     =   3
      VisibleRows     =   10
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
      CellNoteIndicatorColor=   16576
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
        
        frmPreview.spdResultPreview.hWndSpread = spdResult.hwnd
        
        
        frmPreview.spdResultPreview.hWndSpread = spdResult.hwnd
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
    
    Call GetResultList(Format(dtpFrom.Value, "yyyy-mm-dd"), Format(dtpTo.Value, "yyyy-mm-dd"), spdResult)

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

Private Sub spdResult_DblClick(ByVal Col As Long, ByVal Row As Long)

'    frmPatReport.Show

End Sub

Private Sub spdResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    
    '검사결과 수정
    Dim strSaveSeq  As String
    Dim strRsltDate As String
    Dim strRsltTime As String
    Dim strTestNm   As String
    Dim strIntBase  As String
    Dim strResult   As String
    Dim strTC       As String
    Dim strTG       As String
    Dim strHDL      As String
    Dim strBUN      As String
    Dim strCREA     As String
    Dim intCol      As Integer

    
    sRow = spdResult.ActiveRow
    sCol = colPNAME
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdResult, sRow, sCol)
    
    If KeyCode = vbKeyDelete Then
        
        If MsgBox(strNewBarNo & " 를 화면에서 지우시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdResult, sRow, sRow
        spdResult.MaxRows = spdResult.MaxRows - 1
    
    ElseIf KeyCode = vbKeyReturn Then
        '검사결과 수정
'        If sCol > colSTATE And sCol <= spdResult.MaxCols Then
'            If strNewBarNo = "" Then
'                Exit Sub
'            End If
'            With spdResult
'                With mResult
'                    .BarNo = strNewBarNo
'                    If strSaveSeq = "" Then
'                        .RsltDate = Format(Now, "yyyy-mm-dd")
'                        .RsltTime = Format(Now, "hh:mm:ss")
'                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'
'                        Call SetText(spdResult, .RsltDate, sRow, colEXAMDATE)
'                        Call SetText(spdResult, .RsltTime, sRow, colEXAMTIME)
'                        Call SetText(spdResult, .RsltSeq, sRow, colSAVESEQ)
'                    Else
'                        .RsltDate = GetText(spdResult, sRow, colEXAMDATE)
'                        .RsltTime = GetText(spdResult, sRow, colEXAMTIME)
'                        .RsltSeq = GetText(spdResult, sRow, colSAVESEQ)
'                    End If
'                End With
'
'                '-- 결과환자정보
'                Call GetSampleInfo(sRow, spdResult)
'
'                gRow = sRow
'                strTestNm = GetText(spdResult, 0, sCol)      '입력한 컬럼의 검사명 찾기
'                strIntBase = GetChannel(strTestNm)          '검사명으로 채널찾기
'                strResult = GetText(spdResult, sRow, sCol)   '결과값
'
'                '-- 검사결과처리 프로세스
'                If strIntBase <> "" And strResult <> "" Then
'                    If ResultProcess(mResult.BarNo, strIntBase, strResult, strResult) = True Then
'                        'strState = "R"
'                    End If
'                End If
'
'                If gHOSP.PARTNM = "LFT" Then
'                    strTC = "":     strTG = "": strHDL = ""
'                    strBUN = "":    strCREA = ""
'
'                    For intCol = colSTATE + 1 To spdResult.MaxCols
'                        With spdResult
'                            .Row = 0
'                            .Col = intCol
'                            Select Case .Text
'                                Case "TC":  strTC = GetText(spdResult, sRow, intCol)
'                                Case "TG":  strTG = GetText(spdResult, sRow, intCol)
'                                Case "HDL": strHDL = GetText(spdResult, sRow, intCol)
'                                Case "BUN": strBUN = GetText(spdResult, sRow, intCol)
'                                Case "CRE": strCREA = GetText(spdResult, sRow, intCol)
'                            End Select
'                        End With
'                    Next
'
'                    'LDL 계산
'                    If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
'                        strIntBase = "99"
'                        strResult = strTC - ((strTG / 5) + strHDL)
'                        If strResult < 0 Then
'                            strResult = "0"
'                        End If
'                        strTC = ""
'                        strTG = ""
'                        strHDL = ""
'                        '-- 검사결과처리 프로세스
'                        If strIntBase <> "" And strResult <> "" Then
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strResult) = True Then
'                                'strState = "R"
'                            End If
'                        End If
'                    End If
'
'                    'B/C 바율
'                    If strBUN <> "" And strCREA <> "" And IsNumeric(strBUN) And IsNumeric(strCREA) Then
'                        strIntBase = "98"
'                        strResult = strBUN / strCREA
'                        If strResult < 0 Then
'                            strResult = "0"
'                        End If
'                        strBUN = ""
'                        strCREA = ""
'                        '-- 검사결과처리 프로세스
'                        If strIntBase <> "" And strResult <> "" Then
'                            If ResultProcess(mResult.BarNo, strIntBase, strResult, strResult) = True Then
'                                strState = "R"
'                            End If
'                        End If
'                    End If
'                End If
'            End With
'        End If
    End If

End Sub
