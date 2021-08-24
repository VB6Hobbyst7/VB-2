VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm601MachHistory 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "장비이력 관리"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstInstrument 
      BackColor       =   &H00F7FFF7&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7485
      Left            =   210
      TabIndex        =   19
      Top             =   165
      Width           =   2640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7680
      Left            =   2880
      TabIndex        =   11
      Top             =   105
      Width           =   8010
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00FCEFE9&
         Caption         =   "Display"
         Height          =   315
         Left            =   2880
         Style           =   1  '그래픽
         TabIndex        =   36
         Tag             =   "126"
         Top             =   3030
         Width           =   855
      End
      Begin VB.CheckBox chkSDate 
         BackColor       =   &H00DBE6E6&
         Height          =   210
         Left            =   1620
         TabIndex        =   4
         Top             =   3090
         Width           =   195
      End
      Begin MSComctlLib.ListView lvwAction 
         Height          =   2085
         Left            =   4530
         TabIndex        =   7
         Top             =   495
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   3678
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
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00F4F0F2&
         Caption         =   "출력(&P)"
         Height          =   510
         Left            =   3840
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   7035
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   1170
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   7035
         Width           =   1320
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00F4F0F2&
         Caption         =   "화면지움(&C)"
         Height          =   510
         Left            =   5175
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   7035
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Cancel          =   -1  'True
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   6510
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   7035
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00F4F0F2&
         Caption         =   "삭제(&D)"
         Height          =   510
         Left            =   2505
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   7035
         Width           =   1320
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1605
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   2235
         Width           =   2790
      End
      Begin RichTextLib.RichTextBox rtfRemark 
         Height          =   315
         Left            =   1605
         TabIndex        =   3
         Top             =   2640
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   556
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis601.frx":0000
      End
      Begin MSComCtl2.DTPicker dtpStatus 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy""년"" MM""월"" dd""일"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1605
         TabIndex        =   0
         Top             =   1845
         Width           =   1290
         _ExtentX        =   2275
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
         CustomFormat    =   "yyyy/MM/dd"
         Format          =   64028675
         CurrentDate     =   72937
      End
      Begin MedControls1.LisLabel lblPrgBar 
         Height          =   330
         Left            =   165
         TabIndex        =   27
         Top             =   3435
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   582
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
         Caption         =   "◈ 장비상태내역"
         LeftGab         =   100
      End
      Begin FPSpread.vaSpread ssHistory 
         Height          =   3165
         Left            =   165
         TabIndex        =   30
         Top             =   3795
         Width           =   7695
         _Version        =   196608
         _ExtentX        =   13573
         _ExtentY        =   5583
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
         SpreadDesigner  =   "Lis601.frx":009D
      End
      Begin MedControls1.LisLabel lblAction 
         Height          =   330
         Left            =   4545
         TabIndex        =   31
         Top             =   150
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   582
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
         Caption         =   "◈ 조치사항선택"
         LeftGab         =   100
      End
      Begin MSComCtl2.DTPicker dtpStatusTm 
         Height          =   300
         Left            =   2910
         TabIndex        =   1
         Top             =   1845
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
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
         CustomFormat    =   "HH:mm"
         Format          =   64028675
         UpDown          =   -1  'True
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   315
         Left            =   1845
         TabIndex        =   5
         Top             =   3045
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
         Format          =   64028675
         UpDown          =   -1  'True
         CurrentDate     =   36328
      End
      Begin VB.Label lblFrDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Display Date :"
         Height          =   270
         Left            =   210
         TabIndex        =   35
         Tag             =   "30509"
         Top             =   3075
         Width           =   1245
      End
      Begin VB.Label lblEmpID 
         Appearance      =   0  '평면
         BackColor       =   &H00FFF1E6&
         BorderStyle     =   1  '단일 고정
         Caption         =   "양호"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6690
         TabIndex        =   34
         Top             =   3060
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblName 
         Appearance      =   0  '평면
         BackColor       =   &H00FFF1E6&
         BorderStyle     =   1  '단일 고정
         Caption         =   "양호"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5115
         TabIndex        =   33
         Top             =   3060
         Width           =   1560
      End
      Begin VB.Label 등록자 
         BackColor       =   &H00DBE6E6&
         Caption         =   "등   록   자 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3870
         TabIndex        =   32
         Top             =   3075
         Width           =   1110
      End
      Begin VB.Label lblFinalDt 
         Appearance      =   0  '평면
         BackColor       =   &H00FFF1E6&
         BorderStyle     =   1  '단일 고정
         Caption         =   "2002/02/02 13:45"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1605
         TabIndex        =   29
         Top             =   1035
         Width           =   2085
      End
      Begin VB.Label lblFinalStatus 
         Appearance      =   0  '평면
         BackColor       =   &H00FFF1E6&
         BorderStyle     =   1  '단일 고정
         Caption         =   "양호"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1605
         TabIndex        =   28
         Top             =   1425
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "등 록 일 자 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   26
         Top             =   1845
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "최 종 상 태 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   25
         Top             =   1425
         Width           =   1200
      End
      Begin VB.Label lblModelNm 
         Appearance      =   0  '평면
         BackColor       =   &H00FFF1E6&
         BorderStyle     =   1  '단일 고정
         Caption         =   "D6000-0211"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1605
         TabIndex        =   24
         Top             =   645
         Width           =   2775
      End
      Begin VB.Label lblEquipCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "선 택 장 비 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   23
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label lblDRefNm 
         Appearance      =   0  '평면
         BackColor       =   &H00FFF1E6&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Demension"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1605
         TabIndex        =   22
         Top             =   255
         Width           =   2775
      End
      Begin VB.Label lblDRefCd 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
         Caption         =   "P007"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1605
         TabIndex        =   21
         Top             =   135
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblSerialNo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Model No :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   18
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label lblVendor 
         BackColor       =   &H00DBE6E6&
         Caption         =   "장 비 상 태 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   17
         Top             =   2235
         Width           =   1275
      End
      Begin VB.Label lblPuchaseDate 
         BackColor       =   &H00DBE6E6&
         Caption         =   "최종등록일 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   16
         Top             =   1035
         Width           =   1290
      End
      Begin VB.Label lblNotes 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Remark :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   15
         Top             =   2640
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   855
      Left            =   225
      TabIndex        =   12
      Top             =   7785
      Width           =   10650
      Begin Crystal.CrystalReport crtReport 
         Left            =   180
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00CDE7FA&
         Caption         =   "Next      >>"
         Height          =   510
         Left            =   5505
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   225
         Width           =   1320
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00CDE7FA&
         Caption         =   "<< Previous"
         Height          =   510
         Left            =   4140
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   225
         Width           =   1320
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


Private Sub cboStatus_Click()
    If cboStatus.ListIndex = 1 Then
        lblAction.Visible = True
        lvwAction.Visible = True
    Else
        lblAction.Visible = False
        lvwAction.Visible = False
    End If
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub chkSDate_Click()
    If chkSDate.Value = 0 Then
        dtpSDate.Visible = False
    Else
        dtpSDate.Visible = True
    End If
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    
    dtpStatus = GetSystemDate
    dtpStatusTm.Value = Format(GetSystemDate, "hh:mm")
    cboStatus.ListIndex = 0
    rtfRemark.Text = ""
    dtpStatus.SetFocus
    For i = 1 To lvwAction.ListItems.Count
        If lvwAction.ListItems.Item(i).Checked = True Then
            lvwAction.ListItems.Item(i).Checked = False
        End If
    Next
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
                sCalibTm = medGetP(.Value, 2, " ")
                
                sSqlDel = "delete from " & T_LAB601 _
                        & " where " & dbw("eqpcd =", sEqpCd) _
                        & "   and " & dbw("calibdt =", sCalibDt) _
                        & "   and " & dbw("calibtm =", sCalibTm)
                          
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
    
    strRptNm = installdir & "LIS\Rpt\EqpHistoryReport.rpt"

    strFileNm = installdir & "LIS\Rpt\CrystalReport.txt"

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
    'Err.Description
End Sub

Private Sub cmdSave_Click()
    
    Dim sSqlDel As String
    Dim sSqlInsert As String
    Dim sSqlInsert_New As String
    Dim busefg As Boolean
    Dim sCalibDt As String
    Dim sCalibTm As String
    Dim vInsertData As tInsertData
    Dim objSql As New clsLISSqlStatement
    
On Error GoTo DBExecError
    
    If lblDRefCd = "" Then
        MsgBox "작업을 수행할 장비를 선택하세요!", vbCritical, "오류"
        Exit Sub
    End If
    
    If cboStatus.ListIndex = -1 Then
        MsgBox "장비상태를 선택하세요!", vbCritical, "오류"
        cboStatus.SetFocus
        Exit Sub
    End If
    
    sCalibDt = Format(dtpStatus.Value, "yyyymmdd")
    sCalibTm = Format(dtpStatusTm.Value, "hhmm")
    
    dbconn.BeginTrans
    
    sSqlDel = " delete " & T_LAB601 & _
              "  where " & dbw("eqpcd = ", lblDRefCd.Caption) & _
              "    and " & dbw("calibdt = ", sCalibDt) & _
              "    and " & dbw("exptm = ", sCalibTm)

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
    'medMain.stsBar.Panels(2).Text = "정상적으로 처리 되었습니다."
    
    Call InitCollection
    Call DspInstrumentStatus
    
    Exit Sub

DBExecError:
    MsgBox "오류:" & Err.Description, vbCritical, "저장오류"
    dbconn.RollbackTrans
End Sub

Private Function MakeInsertData() As tInsertData
    Dim iRow As Integer
    Dim strItem As String
    Dim LvwItem As ListItem
    
    With MakeInsertData
        .sEqpCd = Trim(lblDRefCd.Caption)
        .sCalibDt = Format(dtpStatus.Value, "yyyymmdd")
        .sExpTm = Format(dtpStatusTm.Value, "hhmm")
        If cboStatus.ListIndex = 1 Then
            .sStatusFg = "1"
        Else
            .sStatusFg = "0"
        End If
        With lvwAction
            'Set LvwItem = .ListItems(.SelectedItem.Index)
            For iRow = 1 To .ListItems.Count
                If .ListItems.Item(iRow).Checked = True Then
                    MakeInsertData.sAction = MakeInsertData.sAction & .ListItems.Item(iRow).Text & COL_DIV
                End If
            Next
        End With
        .sRemark = Trim(rtfRemark.Text)
        .sCalibEmp = lblEmpID.Caption
    End With
End Function

Private Sub dtpStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub dtpStatusTm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
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
    dtpStatus.Value = GetSystemDate
    dtpStatusTm.Value = Format(GetSystemDate, "hh:mm")
    rtfRemark.Text = ""
    lblEmpID.Caption = ObjMyUser.EmpId
    lblName.Caption = ObjMyUser.emplngnm
    dtpSDate.Value = Format(GetSystemDate, "yyyy/mm")
    chkSDate.Value = 1
   
    With cboStatus
        .Clear
        .AddItem "정상"
        .AddItem "점검"
    End With
    cboStatus.ListIndex = 0
    
    For i = 1 To lvwAction.ListItems.Count
        If lvwAction.ListItems.Item(i).Checked = True Then
            lvwAction.ListItems.Item(i).Checked = False
        End If
    Next
    
    With ssHistory
        Call medClearTable(ssHistory)
        .MaxRows = 12
    End With
End Sub

Private Sub DsplstInstrument()
    Dim sSqlGetInstrument As String
    Dim rsGetInstrument As Recordset
    Dim objSql As New clsLISSqlStatement
    Dim i%
    
    Set rsGetInstrument = New Recordset
    rsGetInstrument.Open objSql.SqlInstrument_New, dbconn
    
    If rsGetInstrument.EOF = True Then Exit Sub
    
    lstInstrument.Clear
    For i = 1 To rsGetInstrument.RecordCount
        lstInstrument.AddItem rsGetInstrument.Fields("eqpcd").Value & vbTab & _
                              rsGetInstrument.Fields("eqpnm").Value
                              
        rsGetInstrument.MoveNext
    Next i
    
    Set rsGetInstrument = Nothing
    Set objSql = Nothing
End Sub

Private Sub DspAction()
    Dim strSQL  As String
    Dim Rs      As Recordset
    Dim LvwItem As ListItem
    
    strSQL = "select * from " & T_LAB032 _
           & " where " & dbw("cdindex", "C252", 2) _
           & " order by cdval1"
    
    Set Rs = New Recordset
    Rs.Open strSQL, dbconn
    
    lvwAction.ListItems.Clear
    If Rs.BOF = False Then
        With lvwAction
            
            Do Until Rs.EOF = True
                Set LvwItem = .ListItems.Add()
                
                LvwItem.Text = Rs.Fields("cdval1")
                LvwItem.SubItems(1) = Rs.Fields("field1") & ""
                
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
    Dim sSqlGetInstrumentInfo As String
    Dim rsGetInstrumentInfo As Recordset
    Dim sEqpCd As String
    Dim strDt  As String
    Dim strSDt As String
    Dim strEDt As String
    Dim iRow   As Integer
    
    sEqpCd = Mid(lstInstrument.Text, 1, _
                 InStr(1, lstInstrument.Text, vbTab, vbTextCompare) - 1)
    
    If chkSDate.Value = 1 Then
        strDt = Format(DateAdd("m", 1, dtpSDate), "yyyy-MM") & "-01"
        
        strSDt = Format(dtpSDate.Value, "yyyymm") & "01"
        strEDt = Format(DateAdd("d", -1, strDt), "yyyymmdd")
        
        sSqlGetInstrumentInfo = " SELECT a.eqpcd, a.eqpnm, a.modelnm, b.*, c.empnm " & _
                                " FROM  " & T_LAB006 & " a " & "," & T_LAB601 & " b " & _
                                "," & T_COM006 & " c " & _
                                " where " & dbw("a.eqpcd =", sEqpCd) & _
                                "   and " & DBJ("b.eqpcd = a.eqpcd") & _
                                "   and " & DBJ("c.empid = b.calibemp") & _
                                "   and b.calibdt(+) between '" & strSDt & "' and '" & strEDt & "'" & _
                                " order by b.calibdt desc "
    Else
        sSqlGetInstrumentInfo = " SELECT a.eqpcd, a.eqpnm, a.modelnm, b.*, c.empnm " & _
                                " FROM  " & T_LAB006 & " a " & "," & T_LAB601 & " b " & _
                                "," & T_COM006 & " c " & _
                                " where " & dbw("a.eqpcd =", sEqpCd) & _
                                "   and " & DBJ("b.eqpcd = a.eqpcd") & _
                                "   and " & DBJ("c.empid = b.calibemp") & _
                                " order by b.calibdt desc "
    End If
    
    Set rsGetInstrumentInfo = New Recordset
    rsGetInstrumentInfo.Open sSqlGetInstrumentInfo, dbconn
    
    With ssHistory
        Call medClearTable(ssHistory)
        .MaxRows = 12
        
        If rsGetInstrumentInfo.BOF = False Then
            lblDRefCd.Caption = sEqpCd
            lblDRefNm.Caption = "" & rsGetInstrumentInfo.Fields("eqpnm").Value
            lblModelNm.Caption = "" & rsGetInstrumentInfo.Fields("modelnm").Value
            
            '-- 최종장비상태/최종상태일자
            If IsNull(rsGetInstrumentInfo.Fields("calibdt")) = False Then
                lblFinalDt.Caption = Format(rsGetInstrumentInfo.Fields("calibdt"), "####/##/##")
            Else
                lblFinalDt.Caption = ""
            End If
            If IsNull(rsGetInstrumentInfo.Fields("statusfg")) = False Then
                If rsGetInstrumentInfo.Fields("statusfg") = "0" Then
                    lblFinalStatus.Caption = "정상"
                Else
                    lblFinalStatus.Caption = "점검"
                End If
            End If
            
            'lblPrgBar.Max = iCnt: lblPrgBar.Value = 0
            
            iRow = 1
            Do Until rsGetInstrumentInfo.EOF = True
                If .MaxRows < iRow Then
                    .MaxRows = iRow
                End If
                
                .Row = iRow
                
                .Col = 2 '등록일자
                .Value = Trim(Format(rsGetInstrumentInfo.Fields("calibdt"), "####/##/##") & "" _
                       & " " & Format(rsGetInstrumentInfo.Fields("exptm"), "##:##") & "")
                
                .Col = 3 '장비상태
                If IsNull(rsGetInstrumentInfo.Fields("statusfg")) = False Then
                    If rsGetInstrumentInfo.Fields("statusfg") & "" = "1" Then
                        .Value = "점검"
                    Else
                        .Value = "정상"
                    End If
                End If
                
                .Col = 4 'Remark
                .Value = rsGetInstrumentInfo.Fields("remark") & ""
                
                .Col = 5 '등록자
                .Value = rsGetInstrumentInfo.Fields("empnm") & ""
                
                '-- Hidden=========================================
                .Col = 6 '등록자ID
                .Value = rsGetInstrumentInfo.Fields("calibemp") & ""
                
                .Col = 7 '최종상태코드
                
                .Col = 8 '조치사항코드
                .Value = rsGetInstrumentInfo.Fields("descdx") & ""
                '==================================================
                
                iRow = iRow + 1
                rsGetInstrumentInfo.MoveNext
            Loop
        End If
    End With
    
    Set rsGetInstrumentInfo = Nothing
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

Private Sub rtfRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub ssHistory_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strStatus As String
    Dim aryTemp() As String
    Dim i, j      As Integer
    
    With ssHistory
        If Row < 1 Or Row > .DataRowCnt Then
            Exit Sub
        End If
    
        .Row = Row
        
        .Col = 6 '장비상태코드
        strStatus = .Value
        
        .Col = 2
        If Trim(.Text) <> "" Then
            dtpStatus.Value = medGetP(.Text, 1, " ")
            dtpStatusTm.Value = medGetP(.Text, 2, " ")
        Else
            Exit Sub
        End If
        
        .Col = 3
        If .Value <> "" Then
            cboStatus.Text = .Value
        End If
        
        .Col = 4
        rtfRemark.Text = .Value
        
        .Col = 8
        aryTemp = Split(.Value, COL_DIV)
        
        If cboStatus.ListIndex = 1 Then
            For j = 1 To lvwAction.ListItems.Count
                lvwAction.ListItems.Item(j).Checked = False
            Next
            
            For i = LBound(aryTemp) To UBound(aryTemp)
               For j = 1 To lvwAction.ListItems.Count
                    If lvwAction.ListItems.Item(j).Text = aryTemp(i) Then
                        lvwAction.ListItems.Item(j).Checked = True
                    End If
                Next
            Next
        End If
        
    End With
End Sub
