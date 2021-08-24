VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm601MachHistory 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Lis601.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00F4F0F2&
      Caption         =   "삭제(&D)"
      Height          =   510
      Left            =   7050
      Style           =   1  '그래픽
      TabIndex        =   25
      Top             =   8070
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Cancel          =   -1  'True
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9675
      Style           =   1  '그래픽
      TabIndex        =   24
      Top             =   8070
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   4410
      Style           =   1  '그래픽
      TabIndex        =   23
      Top             =   8070
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   5730
      Style           =   1  '그래픽
      TabIndex        =   22
      Top             =   8070
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8355
      Style           =   1  '그래픽
      TabIndex        =   21
      Top             =   8070
      Width           =   1320
   End
   Begin Crystal.CrystalReport crtReport 
      Left            =   2955
      Top             =   8070
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox lstInstrument 
      BackColor       =   &H00F7FFF7&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6885
      Left            =   135
      TabIndex        =   4
      Top             =   255
      Width           =   2640
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   2895
      TabIndex        =   5
      Top             =   255
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 장비 정보"
      LeftGab         =   100
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   930
      Left            =   2895
      TabIndex        =   6
      Top             =   480
      Width           =   8130
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   4020
         TabIndex        =   33
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Model No"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   4020
         TabIndex        =   34
         Top             =   510
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "최종 상태"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDRefNm 
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   120
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   635
         BackColor       =   16773606
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblModelNm 
         Height          =   360
         Left            =   5415
         TabIndex        =   8
         Top             =   120
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   635
         BackColor       =   16773606
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblFinalDt 
         Height          =   360
         Left            =   1440
         TabIndex        =   9
         Top             =   510
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   635
         BackColor       =   16773606
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblFinalStatus 
         Height          =   360
         Left            =   5415
         TabIndex        =   10
         Top             =   510
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   635
         BackColor       =   16773606
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   60
         TabIndex        =   31
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "선택 장비"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   60
         TabIndex        =   32
         Top             =   510
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "최종점검일"
         Appearance      =   0
      End
      Begin VB.Label lblDRefCd 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Caption         =   "P007"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1695
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Left            =   2910
      TabIndex        =   13
      Top             =   1485
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 점검 사항"
      LeftGab         =   100
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Height          =   2595
      Left            =   2925
      TabIndex        =   12
      Top             =   1710
      Width           =   8100
      Begin MSComctlLib.ListView lvwAction 
         Height          =   1680
         Left            =   1425
         TabIndex        =   14
         Top             =   135
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   2963
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "코드"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "조치사항"
            Object.Width           =   4128
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfRemark 
         Height          =   630
         Left            =   1425
         TabIndex        =   15
         Top             =   1845
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1111
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis601.frx":000C
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   45
         TabIndex        =   28
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "점검 상태"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   30
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "비       고"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel lblPrgBar 
      Height          =   300
      Left            =   2925
      TabIndex        =   26
      Top             =   4365
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "◈ 점검 상태조회"
      LeftGab         =   100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   3345
      Left            =   2925
      TabIndex        =   0
      Top             =   4590
      Width           =   8100
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00FCEFE9&
         Caption         =   "Display"
         Height          =   390
         Left            =   4635
         Style           =   1  '그래픽
         TabIndex        =   27
         Tag             =   "126"
         Top             =   195
         Width           =   915
      End
      Begin FPSpread.vaSpread ssHistory 
         Height          =   2610
         Left            =   90
         TabIndex        =   16
         Top             =   645
         Width           =   7920
         _Version        =   196608
         _ExtentX        =   13970
         _ExtentY        =   4604
         _StockProps     =   64
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColsFrozen      =   3
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   50
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis601.frx":0220
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DBE6E6&
         Height          =   480
         Left            =   1485
         TabIndex        =   17
         Top             =   105
         Width           =   3030
         Begin VB.OptionButton optSel 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체"
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   135
            Width           =   660
         End
         Begin VB.OptionButton optSel 
            BackColor       =   &H00DBE6E6&
            Caption         =   "날짜지정"
            Height          =   315
            Index           =   1
            Left            =   750
            TabIndex        =   18
            Top             =   135
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker dtpSDate 
            Height          =   315
            Left            =   1905
            TabIndex        =   20
            Top             =   135
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy/MM"
            Format          =   66256899
            UpDown          =   -1  'True
            CurrentDate     =   36328
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   375
         Index           =   1
         Left            =   90
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "조회 조건"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   885
      Left            =   150
      TabIndex        =   1
      Top             =   7050
      Width           =   2640
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00CDE7FA&
         Caption         =   "<<Previous"
         Height          =   405
         Left            =   165
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   285
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00CDE7FA&
         Caption         =   "Next     >>"
         Height          =   405
         Left            =   1395
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm601MachHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tInsertData
    sEqpCd    As String
    sCalibDt  As String
    sExpTm  As String
    sCalibEmp As String
    sStatusFg As String
    sAction   As String
    sRemark   As String
End Type
Public Event LastFormUnload()



Private Sub cmdClear_Click()
    Dim i As Integer
    
    lblDRefNm.Caption = ""
    lblDRefCd.Caption = ""
    lblModelNm.Caption = ""
    lblFinalDt.Caption = ""
    lblFinalStatus.Caption = ""
    rtfRemark.Text = ""
    optSel(1).Value = True
    dtpSDate.Value = Format(GetSystemDate, "yyyy/mm")
    
    For i = 1 To lvwAction.ListItems.Count
        lvwAction.ListItems.Item(i).Checked = False
    Next
    
    With ssHistory
        Call medClearTable(ssHistory)
        .MaxRows = 9
    End With
End Sub

Private Sub ClearlstInstrumentContent()
    lstInstrument.Clear
End Sub

Private Sub cmdDelete_Click()
    
    Dim sMsg  As String
    Dim sRes  As Integer, sStyle As Integer
    Dim iRow  As Integer
    Dim bFlag As Boolean
    
    If Trim(lblDRefCd.Caption) = "" Then Exit Sub
    
    With ssHistory
        bFlag = False
        For iRow = 1 To .DataRowCnt
            .Row = iRow
            
            .Col = 2
            If .Value <> "" Then
                .Col = 1
                If .Value = 1 Then
                    bFlag = True
                    Exit For
                End If
            End If
        Next
    End With
    
    If bFlag = False Then
        MsgBox "선택한 항목이 없습니다.", vbCritical, "오류"
        For iRow = 1 To ssHistory.MaxRows
            ssHistory.Row = iRow: ssHistory.Col = 1
            ssHistory.Value = 0
        Next
        Exit Sub
    End If
    
    sMsg = "선택된 항목이 모두 삭제됩니다." & Chr(13) & "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2
    
    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        If DeleteEquipInfo = False Then
            Exit Sub
        End If
        
        'medMain.stsBar.Panels(2).Text = "정상적으로 삭제 처리 되었습니다. 다음 작업을 처리하세요"
        
        Call InitCollection
        Call DspInstrumentStatus
    Else
        Exit Sub
    End If
    
End Sub
    
Private Function DeleteEquipInfo() As Boolean
    
    Dim sSqlDel  As String
    Dim sEqpCd   As String
    Dim sCalibDt As String
    Dim sCalibTm As String
    Dim iRow     As Integer
    
On Error GoTo DBExecError
    
    dbconn.BeginTrans
    
    sEqpCd = lblDRefCd.Caption
    
    With ssHistory
        For iRow = 1 To .DataRowCnt
            .Row = iRow: .Col = 1
            
            If .Value = 1 Then
                
                .Col = 2 '장비상태일자
                sCalibDt = Format(.Value, "yyyymmdd")
                
                .Col = 2 '장비상태시간
                sCalibTm = Format(Replace(medGetP(.Value, 2, " "), ":", ""), "0000")
                
                sSqlDel = "delete from " & T_LAB601 _
                        & " where " & DBW("eqpcd =", sEqpCd) _
                        & "   and " & DBW("calibdt =", sCalibDt) _
                        & "   and " & DBW("exptm =", sCalibTm)
                          
                dbconn.Execute (sSqlDel)
            End If
        Next
    End With
    
    dbconn.CommitTrans
    
    DeleteEquipInfo = True
    
    Exit Function
    
DBExecError:
    MsgBox "오류:" & Err.Description, vbCritical, "삭제오류"
    dbconn.RollbackTrans
    DeleteEquipInfo = False
End Function

Private Sub cmdDisplay_Click()
    If lblDRefCd.Caption <> "" Then
        Call DspInstrumentStatus
    Else
        MsgBox "장비를 선택하세요!", vbExclamation, "조회확인"
    End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
   If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdNext_Click()
    Dim i%
    For i = 0 To lstInstrument.ListCount - 1
        If lstInstrument.Selected(i) = True And (i <> lstInstrument.ListCount - 1) Then
            lstInstrument.Selected(i + 1) = True
            Exit For
        End If
    Next i
    
End Sub

Private Sub cmdPrevious_Click()
    Dim i%
    For i = 0 To lstInstrument.ListCount - 1
        If lstInstrument.Selected(i) = True And i <> 0 Then
            lstInstrument.Selected(i - 1) = True
            Exit For
        End If
    Next i
End Sub

Private Sub cmdPrint_Click()
    Dim i As Long
    Dim strTmp As String
    Dim strFileNm As String
    Dim strRptNm As String
    Dim lngFNum As Long
    Dim strMyFile As String
    Dim lngCnt As Long
    
    With ssHistory
        If .DataRowCnt < 1 Then
            MsgBox "출력할 내용이 없습니다.", vbExclamation, "확인"
            Exit Sub
        End If
    End With
    
    strRptNm = installdir & "lis\rpt\EqpHistoryReport.rpt"
    If Dir(strRptNm) = "" Then
        MsgBox """EqpHistoryReport.rpt"" 파일이 없습니다.", vbExclamation
        Exit Sub
    End If

    strFileNm = installdir & "lis\rpt\CrystalReport.txt"
    
    If Dir(strFileNm) = "" Then
        MsgBox """EqpHistoryReport.rpt"" 파일이 없습니다.", vbExclamation
        Exit Sub
    End If
    
    lngFNum = FreeFile

On Error GoTo ErrPrint
    
    Open strFileNm For Output As #lngFNum
    
    With ssHistory
        strTmp = ""
        For i = 1 To .DataRowCnt
            '장비명
            strTmp = strTmp & lblDRefNm.Caption & vbTab
            
            '모델명
            strTmp = strTmp & lblModelNm.Caption & vbTab
            
            '최종상태
            strTmp = strTmp & lblFinalStatus.Caption & vbTab
            
            .Row = i
            
            .Col = 2
            strTmp = strTmp & .Value & vbTab
            
            .Col = 3
            strTmp = strTmp & .Value & vbTab
            
            .Col = 4
            strTmp = strTmp & .Value & vbTab
            
            .Col = 5
            strTmp = strTmp & .Value & vbNewLine
        Next
        
        Print #lngFNum, Mid(strTmp, 1, Len(strTmp) - 1)
    End With
    
    Close #lngFNum
    With crtReport
        .ReportFileName = strRptNm
        .ParameterFields(0) = "hostnm;" & P_HOSPITALNAME & ";true"
        .RetrieveDataFiles
        .Destination = crptToWindow
        .WindowState = 2 ' crptMaximized
        .Action = 1
        .Reset
    End With
    
    Exit Sub
    
ErrPrint:
    MsgBox "출력하는 도중 에러가 발생했습니다." & vbNewLine & Err.Description, vbExclamation
    
End Sub

Private Sub cmdSave_Click()
    
    Dim sSqlDel As String
    Dim sSqlInsert As String
    Dim sSqlInsert_New As String
    Dim busefg As Boolean
    Dim sCalibDt As String
    Dim sCalibTm As String
    Dim vInsertData As tInsertData
    'Dim objSql As New clsLISSqlStatement
    
On Error GoTo DBExecError
    
    If lstInstrument.SelCount <= 0 Then
        MsgBox "점검을 위한 장비를 선택하세요.", vbExclamation
        Exit Sub
    End If
    
    sCalibDt = Format(GetSystemDate, "yyyyMMdd")
    sCalibTm = Format(GetSystemDate, "hhmm")
    
    dbconn.BeginTrans
    
    sSqlDel = " delete " & T_LAB601 & _
              "  where " & DBW("eqpcd = ", lblDRefCd.Caption) & _
              "    and " & DBW("calibdt = ", sCalibDt) & _
              "    and " & DBW("exptm = ", sCalibTm)

    dbconn.Execute (sSqlDel)
    
    vInsertData = MakeInsertData
                                
    With vInsertData
        sSqlInsert_New = " Insert into " & T_LAB601 & _
                         " values( " & _
                           DBV("eqpcd    ", .sEqpCd) & " ," & _
                           DBV("calibdt  ", .sCalibDt) & " ," & _
                           DBV("exptm  ", .sExpTm) & " ," & _
                           DBV("calibemp ", .sCalibEmp) & " , " & _
                           DBV("statusfg ", .sStatusFg) & " , " & _
                           DBV("descdx ", .sAction) & " , " & _
                           DBV("remark   ", .sRemark) & _
                           ")"
    End With
    
    dbconn.Execute (sSqlInsert_New)
    dbconn.CommitTrans
    
    Call InitCollection
    Call DspInstrumentStatus
    
    Exit Sub

DBExecError:
    dbconn.RollbackTrans
    MsgBox "오류:" & Err.Description, vbCritical, "저장오류"
End Sub

Private Function MakeInsertData() As tInsertData
    Dim iRow As Integer
    Dim strItem As String
    Dim LvwItem As ListItem
    Dim lngSelCnt As Long
    Dim i As Long
    
    With MakeInsertData
        .sEqpCd = Trim(lblDRefCd.Caption)
        .sCalibDt = Format(GetSystemDate, "yyyymmdd")
        .sExpTm = Format(GetSystemDate, "hhmm")
        
        With lvwAction
            For i = 1 To .ListItems.Count
                If .ListItems.Item(i).Checked Then
                    MakeInsertData.sAction = MakeInsertData.sAction & .ListItems.Item(i).Text & COL_DIV
                    lngSelCnt = lngSelCnt + 1
                End If
            Next
        End With
        
        .sStatusFg = IIf(lngSelCnt = 0, "0", "1")
        
        .sRemark = Trim(rtfRemark.Text)
        .sCalibEmp = ObjSysInfo.EmpId
    End With
End Function

Private Sub dtpStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpStatusTm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
   
   Call InitCollection
   Call DsplstInstrument
   '-- 조치사항Display
   Call DspAction
End Sub

Private Sub InitCollection()
    Dim i As Integer
    
    lblDRefNm.Caption = ""
    lblDRefCd.Caption = ""
    lblModelNm.Caption = ""
    lblFinalDt.Caption = ""
    lblFinalStatus.Caption = ""
    rtfRemark.Text = ""
    optSel(1).Value = True
    dtpSDate.Value = Format(GetSystemDate, "yyyy/mm")
    
    For i = 1 To lvwAction.ListItems.Count
        lvwAction.ListItems.Item(i).Checked = False
    Next
    
    With ssHistory
        Call medClearTable(ssHistory)
        .MaxRows = 9
    End With
End Sub

Private Sub DsplstInstrument()
    Dim rsGetInstrument As Recordset
    Dim SqlInstrument_New As String
    Dim i%
    
    SqlInstrument_New = " SELECT * " & _
                        " FROM  " & T_LAB006 & _
                        " ORDER BY eqpcd, eqpnm"

    Set rsGetInstrument = New Recordset
    rsGetInstrument.Open SqlInstrument_New, dbconn
    
    If rsGetInstrument Is Nothing Then Exit Sub
    
    lstInstrument.Clear
    rsGetInstrument.MoveFirst
    Do Until rsGetInstrument.EOF
        lstInstrument.addItem rsGetInstrument.Fields("eqpcd").Value & "" & vbTab & _
                              rsGetInstrument.Fields("eqpnm").Value & ""
        rsGetInstrument.MoveNext
    Loop
    
    Set rsGetInstrument = Nothing
End Sub

Private Sub DspAction()
    Dim strSQL  As String
    Dim Rs      As Recordset
    Dim LvwItem As ListItem
    
    strSQL = "select * from " & T_LAB032 _
           & " where " & DBW("cdindex", "C252", 2) _
           & " order by cdval1"
    
    Set Rs = New Recordset
    Rs.Open strSQL, dbconn
    
    lvwAction.ListItems.Clear
    If Rs.BOF = False Then
        With lvwAction
            
            Do Until Rs.EOF = True
                Set LvwItem = .ListItems.Add()
                
                LvwItem.Text = Rs.Fields("cdval1").Value & ""
                LvwItem.SubItems(1) = Rs.Fields("field1").Value & ""
                
                Rs.MoveNext
            Loop
        End With
    End If
    
    Set Rs = Nothing
End Sub

Private Sub lstInstrument_Click()
    
    Call InitCollection
    Call DspInstrumentStatus
    
End Sub

Private Sub DspInstrumentStatus()
    Dim strSQL As String
    Dim RsInfo As Recordset
    Dim sEqpCd As String
    Dim strDt  As String
    Dim strSDt As String
    Dim strEDt As String
    Dim iRow   As Integer
    Dim varCalDt As Variant
    
    sEqpCd = Mid(lstInstrument.Text, 1, _
                 InStr(1, lstInstrument.Text, vbTab, vbTextCompare) - 1)
    
    If optSel(1).Value Then
    
        strDt = Format(DateAdd("m", 1, dtpSDate), "yyyy-MM") & "-01"
        strSDt = Format(dtpSDate.Value, "yyyymm") & "01"
        strEDt = Format(DateAdd("d", -1, strDt), "yyyymmdd")
        
        strSQL = " SELECT a.eqpcd, a.eqpnm, a.modelnm, b.*, c.empnm " & _
                 " FROM  " & T_LAB006 & " a " & "," & T_LAB601 & " b " & _
                 "," & T_COM006 & " c " & _
                 " where " & DBW("a.eqpcd =", sEqpCd) & _
                 "   and " & DBJ("b.eqpcd = a.eqpcd") & _
                 "   and " & DBJ("c.empid = b.calibemp") & _
                 "   and b.calibdt(+) between '" & strSDt & "' and '" & strEDt & "'" & _
                 " order by b.calibdt,b.exptm desc "
    Else
        strSQL = " SELECT a.eqpcd, a.eqpnm, a.modelnm, b.*, c.empnm " & _
                 " FROM  " & T_LAB006 & " a " & "," & T_LAB601 & " b " & _
                 "," & T_COM006 & " c " & _
                 " where " & DBW("a.eqpcd =", sEqpCd) & _
                 "   and " & DBJ("b.eqpcd = a.eqpcd") & _
                 "   and " & DBJ("c.empid = b.calibemp") & _
                 " order by b.calibdt,b.exptm  desc "
    End If
    
    Set RsInfo = New Recordset
    RsInfo.Open strSQL, dbconn
    
    If RsInfo Is Nothing Then Exit Sub
    
    On Error GoTo ErrTrap
    
    lblDRefCd.Caption = sEqpCd
    lblDRefNm.Caption = "" & RsInfo.Fields("eqpnm").Value & ""
    lblModelNm.Caption = "" & RsInfo.Fields("modelnm").Value & ""
    lblFinalDt.Caption = Format(RsInfo.Fields("calibdt").Value & "", "####/##/##")
    Select Case RsInfo.Fields("statusfg").Value & ""
        Case ""
            lblFinalStatus.Caption = ""
        Case "0"
            lblFinalStatus.Caption = "정상"
        Case "1"
            lblFinalStatus.Caption = "점검"
    End Select
    
       
    With ssHistory
        Call medClearTable(ssHistory)
        RsInfo.MoveFirst
        Do Until RsInfo.EOF
            If .MaxRows <= .DataRowCnt Then
                .MaxRows = .MaxRows + 1
                .RowHeight(-1) = 10
            End If
                        
            .Row = .DataRowCnt + 1
            .Col = 2 '등록일자
            .Value = Trim(Format(RsInfo.Fields("calibdt").Value & "", "####/##/##") & "" _
                   & " " & Format(RsInfo.Fields("exptm").Value & "", "##:##") & "")
                
            Call .GetText(2, 1, varCalDt)
            
            If varCalDt <> "" Then
                .Col = 3 '장비상태
                .Value = IIf(RsInfo.Fields("statusfg").Value & "" = "1", "점검", "정상")
                
                .Col = 4 'Remark
                .Value = RsInfo.Fields("remark").Value & ""
                
                .Col = 5 '등록자
                .Value = RsInfo.Fields("empnm").Value & ""
                
                '-- Hidden=========================================
                .Col = 6 '등록자ID
                .Value = RsInfo.Fields("calibemp").Value & ""
                
                .Col = 7 '최종상태코드
                
                .Col = 8 '조치사항코드
                .Value = RsInfo.Fields("descdx").Value & ""
                '==================================================
            End If
            
            RsInfo.MoveNext
        Loop
    End With
    
    Set RsInfo = Nothing
    Exit Sub
ErrTrap:
    MsgBox Err.Description
    Set RsInfo = Nothing
End Sub

Private Sub lvwAction_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer
        
    '-- 정렬
    With lvwAction
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub optSel_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpSDate.Enabled = False
        Case 1
            dtpSDate.Enabled = True
    End Select
End Sub

Private Sub rtfRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub ssHistory_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strStatus As String
    Dim aryTemp() As String
    Dim i, j      As Long
    Dim itmFound As ListItem
    
    With ssHistory
        If Row < 1 Or Row > .DataRowCnt Then
            Exit Sub
        End If
    
        .Row = Row
        
        .Col = 4: rtfRemark.Text = .Value
        .Col = 8: aryTemp = Split(.Value, COL_DIV)
        
        For i = 1 To lvwAction.ListItems.Count
            lvwAction.ListItems.Item(i).Checked = False
        Next
        
        For i = LBound(aryTemp) To UBound(aryTemp)
            Set itmFound = lvwAction.FindItem(aryTemp(i), lvwText)
            
            If Not itmFound Is Nothing Then
                itmFound.Checked = True
            End If
        Next
    End With
End Sub
