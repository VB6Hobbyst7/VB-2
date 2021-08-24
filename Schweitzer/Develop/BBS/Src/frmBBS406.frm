VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS406 
   BackColor       =   &H00DBE6E6&
   Caption         =   "적격/부적격 판정"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   Icon            =   "frmBBS406.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   18
      Top             =   45
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  조회 조건"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   480
      Left            =   10485
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   10
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   480
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   8
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   480
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   9
      Tag             =   "15101"
      Top             =   8535
      Visible         =   0   'False
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblOkNot 
      Height          =   7320
      Left            =   75
      TabIndex        =   6
      Top             =   1080
      Width           =   7455
      _Version        =   196608
      _ExtentX        =   13150
      _ExtentY        =   12912
      _StockProps     =   64
      BackColorStyle  =   1
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
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   22
      MaxRows         =   27
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS406.frx":076A
      TextTip         =   4
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   7560
      TabIndex        =   11
      Top             =   1080
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "검 사 세 부 내 역"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   7560
      TabIndex        =   12
      Top             =   5085
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "결 과 판 정"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblDetail 
      Height          =   3660
      Left            =   7560
      TabIndex        =   13
      Top             =   1410
      Width           =   6900
      _Version        =   196608
      _ExtentX        =   12171
      _ExtentY        =   6456
      _StockProps     =   64
      BackColorStyle  =   1
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
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   9
      MaxRows         =   13
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS406.frx":1270
      TextTip         =   4
   End
   Begin VB.Frame fraQuery 
      BackColor       =   &H00DBE6E6&
      Height          =   795
      Index           =   0
      Left            =   75
      TabIndex        =   14
      Top             =   285
      Width           =   14400
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   480
         Left            =   12480
         Style           =   1  '그래픽
         TabIndex        =   3
         Tag             =   "15101"
         Top             =   210
         Width           =   1245
      End
      Begin VB.TextBox txtReservedID 
         Height          =   345
         Left            =   6270
         TabIndex        =   2
         Top             =   285
         Width           =   1695
      End
      Begin VB.CommandButton cmdReserved 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11070
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   270
         Width           =   360
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   345
         Left            =   3165
         TabIndex        =   1
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   59965443
         CurrentDate     =   36797
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   345
         Left            =   1530
         TabIndex        =   0
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   59965443
         CurrentDate     =   36797
      End
      Begin MedControls1.LisLabel lblReservedNm 
         Height          =   330
         Left            =   7980
         TabIndex        =   17
         Top             =   285
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   582
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   14
         Left            =   285
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   270
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "검사의뢰일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   0
         Left            =   5190
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   285
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "지정환자"
         Appearance      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2970
         TabIndex        =   15
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   3090
      Left            =   7560
      TabIndex        =   19
      Top             =   5325
      Width           =   6915
      Begin VB.ComboBox cboReason 
         Height          =   300
         ItemData        =   "frmBBS406.frx":181E
         Left            =   4260
         List            =   "frmBBS406.frx":1820
         TabIndex        =   4
         Top             =   210
         Width           =   2535
      End
      Begin VB.ComboBox cboResult 
         Height          =   300
         ItemData        =   "frmBBS406.frx":1822
         Left            =   1140
         List            =   "frmBBS406.frx":1832
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   210
         Width           =   2295
      End
      Begin VB.TextBox txtrmk 
         Height          =   1035
         Left            =   1140
         ScrollBars      =   2  '수직
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1635
         Width           =   5670
      End
      Begin MSComctlLib.ListView lvwReason 
         Height          =   975
         Left            =   1140
         TabIndex        =   21
         Top             =   630
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1720
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   60
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   210
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "결과 판정"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   60
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   630
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "세부 사유"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   60
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1635
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         Caption         =   "참고 사항"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   3465
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
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
         Caption         =   "사유"
         Appearance      =   0
      End
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "적용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   6285
         TabIndex        =   22
         Top             =   2730
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmBBS406"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Enum TblColumn
'    tcORDDT = 1     '검사의뢰일
'    tcDONORNM       '헌혈자명
'    tcDONORID       '헌혈자id
'    tcBLDNO         '혈액번호
'    TcABO           '혈액형         '5
'
'    tcCOMPO         '혈액제제
'    TcVOLUMN        '볼륨
'    tcTESTNM        '검사명
'    tcTESTCD        '검사코드
'    TcRESULT        '결과           '10
'
'    TcRSTUNIT       '결과단위
'    TcJudge         '판단
'    tcREASON        '사유
'    TcLastJudge     '적격여부
'    tcSPCCD         '검체코드       '15
'
'    TcRSTTYPE       '결과타입
'    TcDAY           '일령
'    tcRmk           '결과참고치
'    TcMainReason    '주사유
'    TcSubReason     '세부사유       '20
'
'    tcREMARK        '비고
'    tcACCDT         '접수일자
'End Enum
Private Enum TblCol
    tcOORDDT = 1
    tcDONORNM
    tcDONORID
    tcBLDNO
    tcABO
    tcCompo
    TcVOLUMN
    tcACCDT
    tcOKDIV
    tcRSN
    tcSUBRSN
    tcRMK
    tcSFG
    tcCompocd
End Enum

Private Enum TblColumn1
    tcTESTNM = 1    '검사명
    tcTESTCD        '검사코드
    TcRESULT        '결과
    TcRSTUNIT       '결과단위
    TcJudge         '판단
    tcSPCCD         '검체코드
    TcRSTTYPE       '결과타입
    TcDAY           '일령
    tcRMK           '결과참고치
End Enum
Private blnApply As Boolean

Private WithEvents GetPtInfo As frmPtInfo
Attribute GetPtInfo.VB_VarHelpID = -1

Private Sub cboReason_Click()
    If isSELECT = False Then
        cboReason.ListIndex = 0
        Exit Sub
    End If
    
End Sub

Private Sub cboResult_Click()
    If isSELECT = False Then
        cboResult.ListIndex = 0
        Exit Sub
    End If
    If cboResult.Text = "적격" Then
        lvwReason.Enabled = False
        cboReason.Enabled = False
    Else
        lvwReason.Enabled = True
        cboReason.Enabled = True
    End If
End Sub

Private Sub cmdClear_Click()
    Data_Clear
    Reason_Clear
End Sub
Private Sub Data_Clear()
'화면상의 정보 Clear
    
    dtpFrom = DateAdd("d", -7, GetSystemDate)
    dtpTo = GetSystemDate

    txtrmk.Text = ""
    txtReservedID.Text = ""
    lblReservedNm.Caption = ""
    tblOkNot.MaxRows = 0
    tblDetail.MaxRows = 0
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdReserved_Click()
    Set GetPtInfo = New frmPtInfo
    GetPtInfo.Show 1
End Sub

Private Sub Form_Activate()
    'Data_Clear
    medMain.lblSubMenu.Caption = Me.Caption
End Sub
Private Sub cmdQuery_Click()
    Dim objDonor        As clsDonorOkNot    '적격 부적격변수
    Dim Rs              As Recordset         '조회대상 레코드셋
    Dim ReasonRS        As Recordset         '사유코드 레코드셋
    Dim BRs             As Recordset
    
    Dim FrDt            As String              '검사의뢰일시작
    Dim ToDt            As String              '검사의뢰일 종료
    Dim strTmp          As String              '동일환자 처리하기위한 임시변수
    Dim strPtid         As String
    Dim strname         As String
    Dim sexage          As String
    Dim DonorId         As String
    Dim donoraccdt      As String
    Dim strRsn          As String
    Dim strSRsn         As String
    Dim ii              As Integer
    
    Me.MousePointer = 11
    tblOkNot.MaxRows = 0
    FrDt = Format(dtpFrom.value, PRESENTDATE_FORMAT)
    ToDt = Format(dtpTo.value, PRESENTDATE_FORMAT)
    tblDetail.MaxRows = 0
    Set objDonor = New clsDonorOkNot
    Set Rs = objDonor.QueryJudgeList(FrDt, ToDt, txtReservedID)
    
    
    If Not Rs.EOF Then

        With tblOkNot
            .ReDraw = False
            Do Until Rs.EOF
                If .DataRowCnt + 1 > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                .Row = .DataRowCnt + 1
                
                If strTmp = Rs.Fields("orddt").value & "" Then
                    .Col = TblCol.tcOORDDT: .value = Format(Rs.Fields("orddt").value & "", "0###-##-##"): .ForeColor = .BackColor
                Else
                    .Col = TblCol.tcOORDDT: .value = Format(Rs.Fields("orddt").value & "", "0###-##-##")
                End If

                .Col = TblCol.tcDONORNM:    .value = Rs.Fields("donornm").value & ""
                .Col = TblCol.tcDONORID:    .value = Rs.Fields("donorid").value & ""
                .Col = TblCol.tcACCDT:      .value = Rs.Fields("donoraccdt").value & ""
                Select Case Rs.Fields("okdiv3").value & ""
                    Case "1":   .Col = TblCol.tcOKDIV: .value = "적격"
                    Case "0":   .Col = TblCol.tcOKDIV: .value = "부적격"
                    Case Else:  .Col = TblCol.tcOKDIV: .value = "보류"
                End Select
                
                Set BRs = New Recordset
                BRs.Open objDonor.GetDonorBlood(Rs.Fields("donorid").value & "", Rs.Fields("donoraccdt").value & ""), DBConn
                If Not BRs.EOF Then
                    Do Until BRs.EOF
                        .Col = TblCol.tcDONORNM:    .value = Rs.Fields("donornm").value & ""
                        .Col = TblCol.tcDONORID:    .value = Rs.Fields("donorid").value & ""
                        
                        .Col = TblCol.tcBLDNO: .value = BRs.Fields("bldsrc").value & "" & "-" & _
                                                     BRs.Fields("bldyy").value & "" & "-" & _
                                                     Format(BRs.Fields("bldno").value & "", "000000")
                        .Col = TblCol.tcCompo:      .value = Get_CompNm(BRs.Fields("compocd").value & "")
                        .Col = TblCol.tcABO:        .value = BRs.Fields("abo").value & "" & Rs.Fields("rh").value & ""
                        .Col = TblCol.TcVOLUMN:     .value = BRs.Fields("volumn").value & ""
                        .Col = TblCol.tcRMK:        .value = Rs.Fields("rmk3").value & ""
                        .Col = TblCol.tcCompocd: .value = BRs.Fields("compocd").value & ""
                        If .value = "0" Then .value = ""
                        .Col = TblCol.tcACCDT:      .value = Rs.Fields("donoraccdt").value & ""
                        Select Case Rs.Fields("okdiv3").value & ""
                            Case "1":   .Col = TblCol.tcOKDIV: .value = "적격"
                            Case "0":   .Col = TblCol.tcOKDIV: .value = "부적격"
                            Case Else:  .Col = TblCol.tcOKDIV: .value = "보류"
                        End Select
                        If .DataRowCnt + 1 > .MaxRows Then
                            .MaxRows = .MaxRows + 1
                        End If
                        .Row = .DataRowCnt + 1
                        BRs.MoveNext
                    Loop
                End If
                
'                .Col = TblCol.tcBLDNO:      .value = rs.Fields("bldsrc").value & "" & "-" & _
'                                                     rs.Fields("bldyy").value & "" & "-" & _
'                                                     Format(rs.Fields("bldno").value & "", "000000")
'                If .value = "--000000" Then .value = ""
''                .Col = TblCol.tcCOMPO:      .value = Get_CompNm(rs.Fields("compocd").value & "")
''                .Col = TblCol.TcABO:        .value = rs.Fields("abo").value & rs.Fields("rh").value
''                .Col = TblCol.TcVOLUMN:     .value = rs.Fields("volumn").value & ""
'                If .value = "0" Then .value = ""
'                .Col = TblCol.tcACCDT:      .value = rs.Fields("donoraccdt").value & ""
'                Select Case rs.Fields("okdiv3").value & ""
'                    Case "1":   .Col = TblCol.tcOKDIV: .value = "적격"
'                    Case "0":   .Col = TblCol.tcOKDIV: .value = "부적격"
'                    Case Else:  .Col = TblCol.tcOKDIV: .value = "보류"
'                End Select
'                .Col = TblCol.tcRMK:        .value = rs.Fields("rmk3").value & ""
'                .Col = TblCol.tcCompocd: .value = rs.Fields("compocd").value & ""
                
                strTmp = Rs.Fields("orddt").value & ""
'                strPtid = rs.Fields("donorid").value & ""
'                strname = rs.Fields("donornm").value & ""
                Rs.MoveNext
            Loop
            
            For ii = 1 To .MaxRows
                .Row = ii:
                .Col = TblCol.tcDONORID: DonorId = .value
                .Col = TblCol.tcACCDT:   donoraccdt = .value
                Set ReasonRS = objDonor.GetReason(DonorId, donoraccdt)
                strRsn = "": strSRsn = ""
                If Not ReasonRS.EOF Then
                    Do Until ReasonRS.EOF
                        If ReasonRS.Fields("seq").value & "" = "1" Then
                            strRsn = ReasonRS.Fields("rsncd").value & ""
                        Else
                            strSRsn = strSRsn & ReasonRS.Fields("rsncd").value & "" & COL_DIV
                        End If
                        ReasonRS.MoveNext
                    Loop
                    .Col = TblCol.tcRSN: .value = strRsn
                    If strSRsn <> "" Then
                        strSRsn = Mid(strSRsn, 1, Len(strSRsn) - 1)
                        .Col = TblCol.tcSUBRSN: .value = strSRsn
                    End If
                End If
            Next
            .ReDraw = True
        End With
    End If
    

    Reason_Clear
    Me.MousePointer = 0
    Set Rs = Nothing
    Set BRs = Nothing
    Set objDonor = Nothing
    Set ReasonRS = Nothing
'    Set objProgressBar = Nothing
    
End Sub




Private Sub Form_Load()
    Resaon_Setting
    Data_Clear
End Sub






Private Sub GetPtInfo_Click(ByVal isSELECT As Boolean, ByVal ptInfo As S2BBS_Library.clsPtInformation)
    If isSELECT = False Then Exit Sub
    
    txtReservedID.Text = "": lblReservedNm.Caption = ""
    
    With ptInfo
        txtReservedID = .PtId
        lblReservedNm.Caption = .ptnm
    End With
End Sub

Private Sub lvwReason_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If isSELECT = False Then
        Item.Checked = False
        Exit Sub
    End If
    If cboReason.ListIndex = 0 Then
        MsgBox "주판정사유를 선택하신후 세부사유를 선택하실수 있습니다.", vbInformation + vbOKOnly, "판정사유선택"
        Item.Checked = False
        Exit Sub
    End If
    If Item.tag = medGetP(cboReason.Text, 1, vbTab) Then
        If Item.Checked = True Then
            Call medSleep(10)
            DoEvents
            Item.Checked = False
        End If
    End If

End Sub

Private Function isSELECT() As Boolean
    If tblOkNot.MaxRows > 0 Then isSELECT = True
End Function

Private Sub tblOkNot_Click(ByVal Col As Long, ByVal Row As Long)
    Dim DonorId    As String
    Dim donoraccdt As String
    
    Dim ResultRS   As Recordset
    Dim Rs         As Recordset
    
    Dim objSql     As clsDonorOkNot
    Dim objDetail  As clsDictionary
    
    Dim sWorkArea  As String
    Dim sAccDt     As String
    Dim sAccSeq    As String
    Dim ii         As Integer
    Dim strTmp     As String
    
    Dim strTmpTest As String
    
    
    If isSELECT = False Then Exit Sub
    If Row < 1 Then Exit Sub
    
    Set objSql = New clsDonorOkNot
    Set objDetail = New clsDictionary
    objDetail.Clear
    
    '결과 display 시 구분하기 위해서,필드추가(panelfg): 2001/10/26
    objDetail.FieldInialize "seq", "testnm,testcd,result,rstunit,resultfg,panelfg"
    
    
    
    With tblOkNot
        .Row = Row:
        .Col = TblCol.tcDONORID: DonorId = .value
        .Col = TblCol.tcACCDT:   donoraccdt = .value
        objDetail.Sort = False
        Set Rs = objSql.GetWorkareaAccdtAccSeq(DonorId, donoraccdt)
        If Not Rs.EOF Then
            Do Until Rs.EOF
                sWorkArea = Rs.Fields("workarea").value & ""
                sAccDt = Rs.Fields("accdt").value & ""
                sAccSeq = Rs.Fields("accseq").value & ""
                'Query 문 안에도 b.panelfg추가
                'Dictionary AddNew문에도 추가
                
                Set ResultRS = objSql.GetTestResult(sWorkArea, sAccDt, sAccSeq)
                If Not ResultRS.EOF Then
                    strTmpTest = objSql.GetResultCdNm(sWorkArea, sAccDt, sAccSeq)
                    If strTmpTest <> "" Then
                        ii = ii + 1
                        '                                                                               요거("")추가:2001/10/26
                        objDetail.AddNew ii, Join(Array(medGetP(strTmpTest, 2, COL_DIV), "", "", "", "1", ""), COL_DIV)
                    End If
                    
                    Do Until ResultRS.EOF
                        ii = ii + 1
                        strTmp = ResultRS.Fields("testnm").value & "" & COL_DIV & ResultRS.Fields("testcd").value & COL_DIV & _
                                 ResultRS.Fields("rstcd").value & "" & COL_DIV & ResultRS.Fields("rstunit").value & COL_DIV & "0" & COL_DIV & _
                                 ResultRS.Fields("panelfg").value & ""
                                '위에꺼추가(ResultRS.Fields("panelfg").value & "")
                        objDetail.AddNew ii, strTmp
                        ResultRS.MoveNext
                    Loop
                End If
                strTmpTest = ""
                Rs.MoveNext
            Loop
        End If
        Set Rs = Nothing
        Set ResultRS = Nothing
    End With
    If objDetail.RecordCount > 0 Then
        With tblDetail
            .MaxRows = 0: .MaxRows = objDetail.RecordCount: ii = 1
            .ReDraw = False
            objDetail.MoveFirst
            Do Until objDetail.EOF
                .Row = ii
                '아래부분 수정(if 문):2001/10/26
                If objDetail.Fields("resultfg") = "1" Then
                    .Col = TblColumn1.tcTESTNM:   .value = objDetail.Fields("testnm"): .ForeColor = DCM_LightRed: .FontBold = True: .FontSize = 10
            
                ElseIf objDetail.Fields("panelfg") & "" = "D" Then '
                    .Col = TblColumn1.tcTESTNM:   .value = objDetail.Fields("testnm"):  .BackColor = DCM_LightGray: .ForeColor = DCM_LightBlue: .FontBold = True
                Else
                    .Col = TblColumn1.tcTESTNM:   .value = objDetail.Fields("testnm")
                End If
                
                .Col = TblColumn1.tcTESTCD:   .value = objDetail.Fields("testcd")
                .Col = TblColumn1.TcRESULT:   .value = objDetail.Fields("result")
                .Col = TblColumn1.TcRSTUNIT:  .value = objDetail.Fields("rstunit")
                ii = ii + 1
                objDetail.MoveNext
            Loop
        End With
    Else
        tblDetail.MaxRows = 0
    End If
   

    Dim itmX     As ListItem
    Dim reason() As String
   
    
    With tblOkNot
        .Row = Row: .Col = TblCol.tcOKDIV
        Select Case .value
            Case "적격":   cboResult.ListIndex = 1
            Case "부적격": cboResult.ListIndex = 2
            Case "보류":   cboResult.ListIndex = 3
            Case "미등록"
'                MsgBox "미등록 헌혈자에 대해서는 적격/부적격 판정을 할수 없습니다.", vbInformation + vbOKOnly, "적격/부적격판정"
        End Select
        
        
        .Col = TblCol.tcRMK:      txtrmk = .value
        .Col = TblCol.tcSUBRSN:   reason() = Split(.value, COL_DIV)
        .Col = TblCol.tcRSN
        If .value = "" Then
            cboReason.ListIndex = 0
        Else
            For ii = 0 To cboReason.ListCount
                If medGetP(cboReason.List(ii), 1, vbTab) = .value Then
                    cboReason.ListIndex = ii
                    Exit For
                End If
            Next
        End If
        
        If UBound(reason) < 0 Then
            For Each itmX In lvwReason.ListItems
                itmX.Checked = False
            Next itmX
            Exit Sub
        End If
        For ii = 0 To UBound(reason)
            For Each itmX In lvwReason.ListItems
                If itmX.tag = reason(ii) Then
                    itmX.Checked = True
                End If
            Next itmX
        Next ii
    End With
    Set objDetail = Nothing
End Sub

Private Sub lblApply_Click()
    Dim itmX   As ListItem
    Dim strTmp As String
    
    If isSELECT = False Then Exit Sub
    
    If cboResult.ListIndex = 0 Then
        MsgBox "결과 판정을 하신후 변경사항을 적용하세요.", vbInformation + vbOKOnly, "결과 판정"
        Exit Sub
    End If
    
    
    With tblOkNot
                
        .Row = .ActiveRow
        .Col = TblCol.tcOKDIV
        
        If .value <> cboReason.Text Then
            If cboResult.ListIndex > 1 Then
                If cboReason.ListIndex = 0 Then
                    MsgBox "사유를 선택하신후 진행하세요.", vbCritical + vbOKOnly, "판정사유선택"
                    Exit Sub
                End If
            End If
        End If
        
        .Col = TblCol.tcOKDIV:   .value = cboResult.Text
        If cboReason.ListIndex = 0 Then
            .Col = TblCol.tcRSN:    .value = ""
        Else
            .Col = TblCol.tcRSN: .value = medGetP(cboReason.Text, 1, vbTab)
        End If
        .Col = TblCol.tcRMK:      .value = txtrmk
        
        For Each itmX In lvwReason.ListItems
            If itmX.Checked = True Then
                strTmp = strTmp & COL_DIV & itmX.tag
                itmX.Checked = False
            End If
        Next itmX
        strTmp = Mid(strTmp, 2)
        .Col = TblCol.tcSUBRSN:   .value = "": .value = strTmp
        txtrmk = ""
        cboReason.ListIndex = 0
        cboResult.ListIndex = 0
        .Col = TblCol.tcSFG: .value = 1
        
    End With
    blnApply = True
End Sub

Private Sub Resaon_Setting()
    Dim objOkNot As New clsDonorOkNot
    
    objOkNot.Reason_List cboReason
    cboReason.ListIndex = 0
    lvw_Display
End Sub
Private Sub lvw_Display()
    Dim Rs       As New Recordset
    Dim objOkNot As New clsDonorOkNot
    Dim itmX As ListItem
    
    With objOkNot
        Set Rs = .Get_Judge_Reason_List
    End With
        
    lvwReason.ListItems.Clear
    Do Until Rs.EOF
        Set itmX = lvwReason.ListItems.Add(, , Rs.Fields("field1").value & "")
            itmX.tag = Rs.Fields("cdval1").value & ""
        Rs.MoveNext
    Loop
    
    Set Rs = Nothing
    Set objOkNot = Nothing
End Sub

Private Sub Reason_Clear()
    Dim itmX As ListItem
    
    txtrmk = ""
    cboReason.ListIndex = 0
    cboResult.ListIndex = 0
    
    For Each itmX In lvwReason.ListItems
        itmX.Checked = False
    Next itmX
    
End Sub

Private Sub cmdSave_Click()
'저장한다...헌혈
    Dim objOkNot   As clsDonorOkNot
    Dim arySql()   As String
    Dim DonorId    As String
    Dim donoraccdt As String
    Dim rmk        As String
    Dim okfg       As String
    Dim okdt       As String
    Dim MainReason As String
    Dim SubRsn()   As String
    Dim sBldNum    As String
    Dim sCompocd   As String
    Dim SSQL       As String
    
    Dim lngsubrsn  As Long
    Dim Cnt        As Long
    Dim ii         As Long
    Dim jj         As Long
    
    If blnApply = False Then
        MsgBox "적용버튼을 선택하여 변경사항을 적용한후 저장작업을 진행하세요.", vbInformation + vbOKOnly, "결과 적용"
        Exit Sub
    End If
    
    Set objOkNot = New clsDonorOkNot
    
    okdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    
    '헌혈자 적격여부문진내역 작성
    With tblOkNot
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblCol.tcSFG
            If .value = "1" Then
                .Col = TblCol.tcACCDT
                If .value <> "" Then
                    donoraccdt = Trim(.value)
                    
                    .Col = TblCol.tcDONORID: DonorId = .value
                    .Col = TblCol.tcBLDNO: sBldNum = .value
                    .Col = TblCol.tcCompocd: sCompocd = .value
                    
                    .Col = TblCol.tcOKDIV
                    Select Case .value
                        Case "적격":   okfg = "1"
                        Case "부적격": okfg = "0"
                        Case Else:     okfg = ""
                    End Select
                    .Col = TblCol.tcRMK: rmk = .value
                    
                    Cnt = Cnt + 1: ReDim Preserve arySql(Cnt - 1)
                    arySql(Cnt - 1) = objOkNot.Set_DonorTestSave(DonorId, donoraccdt, okfg, okdt, rmk)
                    
                    
                    '적격판정이 나면 아래 문장은 탈필요가 없겠지....
                    '2001/10/05 울산동강병원
                    '부적격 판정이 났다가 적격으로 바뀌면, 부적격 사유내역에서 DELETE해주어야 한다..
                    
                    Cnt = Cnt + 1: ReDim Preserve arySql(Cnt - 1)
                    arySql(Cnt - 1) = objOkNot.Delete_Rsncd(DonorId, donoraccdt)
                    If okfg = "0" Then
                        If sBldNum <> "" Then
                            Cnt = Cnt + 1: ReDim Preserve arySql(Cnt - 1)
                            SSQL = " Update " & T_BBS401 & " set " & _
                                         DBW("stscd", BBSBloodStatus.stsEXPIRE, 3) & _
                                         DBW("realexpdt", Format(GetSystemDate, "YYYYMMDD"), 3) & _
                                         DBW("realexptm", Format(GetSystemDate, "HHMMSS"), 3) & _
                                         DBW("expid", ObjSysInfo.EmpId, 3) & _
                                         DBW("exprcvid", ObjSysInfo.EmpId, 2) & _
                                " WHERE " & _
                                          DBW("donorid=", DonorId) & _
                                " AND " & DBW("donoraccdt=", donoraccdt)
                            arySql(Cnt - 1) = SSQL
                        End If
                        
                        .Col = TblCol.tcRSN: MainReason = .value
                        Cnt = Cnt + 1: ReDim Preserve arySql(Cnt - 1)
                        arySql(Cnt - 1) = objOkNot.Set_MainRsncd(DonorId, donoraccdt, MainReason)
                        
                        .Col = TblCol.tcSUBRSN
                        SubRsn() = Split(.value, COL_DIV)
                        If UBound(SubRsn) > -1 Then
                            lngsubrsn = 2
                            For jj = 0 To UBound(SubRsn)
                                Cnt = Cnt + 1: ReDim Preserve arySql(Cnt - 1)
                                arySql(Cnt - 1) = objOkNot.Set_SubRsncd(DonorId, donoraccdt, SubRsn(jj), lngsubrsn)
                                lngsubrsn = lngsubrsn + 1
                            Next
                        End If
                    Else
                        Cnt = Cnt + 1: ReDim Preserve arySql(Cnt - 1)
                        SSQL = " Update " & T_BBS401 & " set " & _
                                     DBW("stscd", BBSBloodStatus.stsENTER, 3) & _
                                     DBW("realexpdt", "", 3) & _
                                     DBW("realexptm", "", 3) & _
                                     DBW("expid", "", 3) & _
                                     DBW("exprcvid", "", 2) & _
                                " WHERE " & _
                                          DBW("donorid=", DonorId) & _
                                " AND " & DBW("donoraccdt=", donoraccdt)
                        arySql(Cnt - 1) = SSQL
                    End If
                
                
                End If
            End If
        Next
    End With
    
    ReDim Preserve arySql(Cnt)
    
    If Cnt > 0 Then
        If InsertData(arySql) Then
            MsgBox "헌혈자 검사내역의 적격/부적격여부가 등록되었습니다.", vbInformation + vbOKOnly, "적격/부적격 등록"
        End If
    End If
    blnApply = False
    
    Set objOkNot = Nothing
End Sub

Private Sub tblOkNot_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    Dim strDonorid    As String
    Dim strDonoraccdt As String
    Dim strTmp        As String
    
    
    If Row < 1 Then Exit Sub
    Dim objSql As New clsGetSqlStatement
    Dim Rs     As Recordset
    
    With tblOkNot
        Call .SetTextTipAppearance("굴림체", 10, False, False, &HFFFFC0, vbBlack)
        .Row = Row
        .Col = TblCol.tcDONORID:         strDonorid = .value
        .Col = TblCol.tcACCDT:           strDonoraccdt = .value
        Set Rs = objSql.Get_DonorBlood(strDonorid, strDonoraccdt)
        If Not Rs.EOF Then
            Do Until Rs.EOF
                strTmp = strTmp & "    " & Rs.Fields("bldsrc").value & "" & "-" & Rs.Fields("bldyy").value & "" & "-" & Format(Rs.Fields("bldno").value & "", "000000") & _
                        "     " & Rs.Fields("abbrnm").value & "" & "    " & Rs.Fields("volumn").value & "" & "cc" & vbNewLine
                Rs.MoveNext
            Loop
        End If
        If strTmp <> "" Then
            strTmp = "  입고혈액 " & vbNewLine & strTmp
            TipWidth = 5000
            MultiLine = 1
            TipText = vbNewLine & strTmp & vbNewLine
            ShowTip = True
        End If

    End With
    
    Set Rs = Nothing
    Set objSql = Nothing
End Sub



Private Sub txtReservedID_Change()
    lblReservedNm.Caption = ""
End Sub

Private Sub txtReservedID_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtReservedID.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtReservedID_LostFocus()
    Dim Rs As New Recordset
    Dim objMySQL As clsBBSSQLStatement
    Dim objPtInfo As clsPtInformation
    Dim Reserved As clsBBSSQLStatement
    
    If Trim(txtReservedID.Text) = "" Or Trim(txtReservedID.Text) = 0 Then Exit Sub
    
    Set objPtInfo = New clsPtInformation
    Set Reserved = New clsBBSSQLStatement
    
    
    Set objMySQL = New clsBBSSQLStatement
    
    Set Rs = New Recordset
    Rs.Open objPtInfo.GetPtInfo(Trim(txtReservedID.Text), True, GetSystemDate), DBConn
    
    If Rs.EOF Then
        MsgBox "등록된 환자가 아니거나 잘못된 문장입니다.", vbInformation, "정보확인"
        With txtReservedID
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
        Set Reserved = Nothing
        Set objPtInfo = Nothing
        Set objMySQL = Nothing
        Exit Sub
    End If
    
    txtReservedID.Text = Rs.Fields("ptid").value & ""
    lblReservedNm.Caption = Rs.Fields("ptnm").value & ""
    
    Set Reserved = Nothing
    Set objPtInfo = Nothing
    Set objMySQL = Nothing
End Sub

Private Sub txtrmk_Change()
    If isSELECT = False Then
        txtrmk = ""
        Exit Sub
    End If
   
End Sub

