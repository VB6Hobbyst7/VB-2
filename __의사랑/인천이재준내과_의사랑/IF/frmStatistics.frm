VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmStatistics 
   Caption         =   "검사통계"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14880
   Icon            =   "frmStatistics.frx":0000
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
      Begin VB.CheckBox chkMon 
         BackColor       =   &H00F8E4D8&
         Caption         =   "월통계"
         Height          =   315
         Left            =   3720
         TabIndex        =   13
         Top             =   120
         Width           =   855
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
         Left            =   6900
         TabIndex        =   2
         Top             =   90
         Value           =   -1  'True
         Width           =   675
      End
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
         Left            =   6900
         TabIndex        =   1
         Top             =   300
         Width           =   675
      End
      Begin HSCotrol.CButton cmdClose 
         Height          =   375
         Left            =   9870
         TabIndex        =   3
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
         HoverPicture    =   "frmStatistics.frx":1084A
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
         TabIndex        =   4
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
         Format          =   146800641
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2310
         TabIndex        =   5
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
         Format          =   146800641
         CurrentDate     =   40457
      End
      Begin HSCotrol.CButton cmdSearch 
         Height          =   375
         Left            =   4620
         TabIndex        =   6
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BackColor       =   15698777
         Caption         =   "조회"
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
         HoverPicture    =   "frmStatistics.frx":10DE4
      End
      Begin HSCotrol.CButton cmdExcel 
         Height          =   375
         Left            =   5730
         TabIndex        =   7
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
         HoverPicture    =   "frmStatistics.frx":10F3E
      End
      Begin HSCotrol.CButton cmdRsltPrint 
         Height          =   375
         Left            =   7650
         TabIndex        =   8
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
         HoverPicture    =   "frmStatistics.frx":11098
      End
      Begin HSCotrol.CButton cmdClear 
         Height          =   375
         Left            =   8760
         TabIndex        =   9
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
         HoverPicture    =   "frmStatistics.frx":111F2
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
         TabIndex        =   11
         Top             =   180
         Width           =   360
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
         TabIndex        =   10
         Top             =   180
         Width           =   150
      End
   End
   Begin FPSpreadADO.fpSpread spdStatistics 
      CausesValidation=   0   'False
      Height          =   8895
      Left            =   150
      TabIndex        =   12
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
      MaxCols         =   1
      MaxRows         =   489
      Protect         =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmStatistics.frx":1134C
      VisibleCols     =   1
      VisibleRows     =   10
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
      CellNoteIndicatorColor=   16576
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnCheck    As Boolean

Private Sub cmdClear_Click()
    
    spdStatistics.MaxRows = 0
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub


Private Sub cmdExcel_Click()
    Dim sFilename As String
            
On Error GoTo ErrHandler

    If spdStatistics.DataRowCnt < 1 Then
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
            sFilename = CommonDialog1.Filename
            SaveExcel sFilename, spdStatistics
            MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
        End With
    End If

Exit Sub
  
ErrHandler:
    ' 사용자가 [취소] 단추를 눌렀습니다.
    Exit Sub

End Sub

Private Sub cmdRsltPrint_Click()
    
    If spdStatistics.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    Else
        If optPrtOri(0).Value = True Then
            spdStatistics.PrintOrientation = PrintOrientationPortrait       '세로출력
        Else
            spdStatistics.PrintOrientation = PrintOrientationLandscape      '가로출력
        End If
'        spdResult.Action = 13
        
        frmPreview.spdResultPreview.hWndSpread = spdStatistics.hwnd
        
        
        frmPreview.spdResultPreview.hWndSpread = spdStatistics.hwnd
        frmPreview.Show vbModal
                
    End If
    
End Sub


Private Sub cmdSearch_Click()

    spdStatistics.MaxRows = 0
    
    Call GetResultStatistics(Format(dtpFrom.Value, "yyyy-mm-dd"), Format(dtpTo.Value, "yyyy-mm-dd"), IIf(chkMon.Value = "1", True, False), spdStatistics)

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    dtpFrom.Value = Format(Now, "YYYY-MM-01")
    dtpTo.Value = Now

    blnCheck = True
    
    spdStatistics.MaxRows = 0
    
    '-- 컬럼보이기설정
    'Call SetColumnView(spdStatistics)
    
    '-- 검사명 보이기
    'Call SetExamCode(spdStatistics)
    
    
    With spdStatistics
        .MaxCols = 1 + UBound(gArrEQP)
        For i = 0 To UBound(gArrEQP) - 1
            .Col = 1 + (i + 1)
            .Row = -1
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .ColWidth(1 + (i + 1)) = gCOLWIDTH
            .FontBold = False
            
            '5 : 약어명으로 표시한다.
            Call SetText(spdStatistics, Trim(gArrEQP(i + 1, 6)), 0, 1 + (i + 1))
            
            '계산식 여부
            If gArrEQP((i + 1), 14) = "" Or gArrEQP((i + 1), 14) = "0" Then
                .FontBold = False
            Else
                .FontBold = True    '계산식 검사일 경우 굵게 표현한다.
            End If
        Next
    End With
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Resize()

    If Me.ScaleHeight = 0 Then Exit Sub

    spdStatistics.WIDTH = Me.ScaleWidth - 300
    spdStatistics.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - 300
    
End Sub

Private Sub spdStatistics_Click(ByVal Col As Long, ByVal Row As Long)
    Dim iRow As Long
    
    If Row = 0 And Col = colCHECKBOX Then
        With spdStatistics
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
        If GetText(spdStatistics, Row, colCHECKBOX) = "1" Then
            Call SetText(spdStatistics, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdStatistics, "1", Row, colCHECKBOX)
        End If
    End If


End Sub


