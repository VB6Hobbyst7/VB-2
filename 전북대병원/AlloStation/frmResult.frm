VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResult 
   Caption         =   "결과조회"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21315
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   21315
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox picHeader 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00004000&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   21315
      TabIndex        =   7
      Top             =   0
      Width           =   21315
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "결과삭제"
         Height          =   375
         Left            =   8250
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   150
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12600
         Top             =   150
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "화면정리"
         Height          =   375
         Left            =   9390
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "엑셀저장"
         Height          =   375
         Left            =   7110
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "선택저장"
         Height          =   375
         Left            =   5970
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "결과조회"
         Height          =   375
         Left            =   4830
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "닫기"
         Height          =   375
         Left            =   10530
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   150
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
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
         Format          =   130875393
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   3150
         TabIndex        =   1
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
         Format          =   130875393
         CurrentDate     =   40457
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   465
         Left            =   4770
         Top             =   90
         Width           =   6885
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   465
         Left            =   240
         Top             =   90
         Width           =   4485
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
         Left            =   330
         TabIndex        =   9
         Top             =   240
         Width           =   930
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
         Left            =   2940
         TabIndex        =   8
         Top             =   240
         Width           =   150
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   8835
      Left            =   60
      TabIndex        =   10
      Top             =   690
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
      SpreadDesigner  =   "frmResult.frx":030A
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

    spdResult.MaxRows = 0
    
    Call GetResultList(Format(dtpFrom.Value, "yyyy-mm-dd"), Format(dtpTo.Value, "yyyy-mm-dd"), spdResult)

End Sub

Private Sub Form_Load()
    
    dtpFrom.Value = Now
    dtpTo.Value = Now

    blnCheck = True
    
    spdResult.MaxRows = 0
    
    'Call SetColumnView(spdResult)
    
    '-- 검사명 보이기
    Call SetExamCode(spdResult)
    
End Sub

Private Sub Form_Resize()

    If Me.ScaleHeight = 0 Then Exit Sub

    spdResult.Width = Me.ScaleWidth - 300
    spdResult.Height = Me.ScaleHeight - picHeader.Height - 300
    
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
