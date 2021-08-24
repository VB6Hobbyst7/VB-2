VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS206 
   BackColor       =   &H00DBE6E6&
   Caption         =   "처방상태조회"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14550
   Icon            =   "frmBBS206.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14550
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   2
      Left            =   5175
      TabIndex        =   37
      Top             =   2235
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   2
      Caption         =   "※ 환자ID 조건 이외의 조건으로 조회 시에는 많은 시간이 소요될 수 있습니다."
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   45
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   30
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "조회 조건"
      Appearance      =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1950
      Left            =   45
      TabIndex        =   18
      Top             =   270
      Width           =   14400
      Begin VB.TextBox txtDoct 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6450
         MaxLength       =   10
         TabIndex        =   4
         Top             =   990
         Width           =   1365
      End
      Begin VB.CommandButton cmdPop 
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
         Height          =   315
         Index           =   3
         Left            =   7845
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   990
         Width           =   300
      End
      Begin VB.TextBox txtDeptCd 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6450
         MaxLength       =   10
         TabIndex        =   3
         Top             =   630
         Width           =   1365
      End
      Begin VB.CommandButton cmdPop 
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
         Height          =   315
         Index           =   2
         Left            =   7845
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   630
         Width           =   300
      End
      Begin VB.TextBox txtPtID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6450
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1350
         Width           =   1365
      End
      Begin VB.CommandButton cmdPop 
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
         Height          =   315
         Index           =   1
         Left            =   7845
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdPop 
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
         Height          =   315
         Index           =   0
         Left            =   7845
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   270
         Width           =   300
      End
      Begin VB.TextBox txtWardID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6450
         MaxLength       =   10
         TabIndex        =   2
         Top             =   270
         Width           =   1365
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   480
         Left            =   12780
         Style           =   1  '그래픽
         TabIndex        =   6
         Tag             =   "15101"
         Top             =   1155
         Width           =   1245
      End
      Begin VB.CheckBox chkTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         Height          =   240
         Left            =   360
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1380
         Width           =   675
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   495
         Left            =   1140
         TabIndex        =   19
         Top             =   1200
         Width           =   3975
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "처방"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   180
            Width           =   675
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈"
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   180
            Width           =   675
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "접수"
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Width           =   675
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "검사중"
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   180
            Width           =   855
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "완료"
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   180
            Width           =   735
         End
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   330
         Left            =   1500
         TabIndex        =   0
         Top             =   465
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   69533699
         CurrentDate     =   36803
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   330
         Left            =   3180
         TabIndex        =   1
         Top             =   465
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   69533699
         CurrentDate     =   36803
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   330
         Left            =   8175
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   270
         Width           =   3675
         _ExtentX        =   6482
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   330
         Left            =   8175
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   630
         Width           =   3675
         _ExtentX        =   6482
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   330
         Left            =   8175
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   990
         Width           =   3675
         _ExtentX        =   6482
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   8175
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1350
         Width           =   3675
         _ExtentX        =   6482
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   5220
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
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
         Caption         =   "병  동"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   5220
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   630
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
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
         Caption         =   "진료과 "
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   5220
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   990
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
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
         Caption         =   "처방의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   5220
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   285
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   465
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
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
         Caption         =   "예정일자"
         Appearance      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3000
         TabIndex        =   35
         Tag             =   "103"
         Top             =   525
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   10470
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "15101"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   11790
      Style           =   1  '그래픽
      TabIndex        =   8
      Tag             =   "124"
      Top             =   8520
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   13110
      Style           =   1  '그래픽
      TabIndex        =   9
      Tag             =   "124"
      Top             =   8520
      Width           =   1320
   End
   Begin VB.TextBox txtRemark 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   1470
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   8085
      Visible         =   0   'False
      Width           =   8145
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   5835
      Left            =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2580
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   10292
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
      MaxCols         =   18
      MaxRows         =   21
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS206.frx":076A
      TextTip         =   4
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   990
      Top             =   8505
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   45
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2235
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "조회 리스트"
      Appearance      =   0
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방 Remark"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   270
      TabIndex        =   36
      Top             =   8145
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmBBS206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    tcPTID = 1
    tcPTNM
    tcABO
    tcORDDT
    tcORDNM
    
    tcREQDT
    TcTRANSRSN
    tcWARD
    tcDEPT
    
    tcDOCTNM
    TcSTS
    tcUNIT
    TcA_Cnt
    TcD_CNT
    tcR_Cnt
    tcE_Cnt
    tcDCFG
    
    TcMESG
End Enum

Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
    
Private Sub chkTot_Click()
    If chkTot.value = 1 Then
        chkQue(0).value = 1
        chkQue(1).value = 1
        chkQue(2).value = 1
        chkQue(3).value = 1
        chkQue(4).value = 1
    Else
        chkQue(0).value = 0
        chkQue(1).value = 0
        chkQue(2).value = 0
        chkQue(3).value = 0
        chkQue(4).value = 0
    End If
End Sub

Private Sub cmdClear_Click()
    
    txtDeptCd.Text = "": txtDoct.Text = ""
    txtPtid.Text = "": txtWardID.Text = ""
    
    Call medClearTable(tblList)
    dtpFrom.value = DateAdd("d", -3, GetSystemDate)
    dtpTo.value = DateAdd("d", 3, GetSystemDate)
    
    cmdPrint.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
'출력하자.....크리스탈
    Dim strTmp As String
    Dim strRfile As String
    Dim strRptPath As String
    Dim intFNum As Integer
    Dim ii As Integer
    Dim jj As Integer
    
    Me.MousePointer = 11
    With tblList
        For ii = 1 To .MaxRows
            .Row = ii
            For jj = TblColumn.tcPTID To TblColumn.tcDCFG
                .Col = jj
                strTmp = strTmp & Trim(.value) & vbTab
            Next jj
            strTmp = strTmp & vbCr
        Next ii
    End With
        
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    
    strRfile = App.Path & "\Rpt\CrystalReport.txt"
    strRptPath = App.Path & "\Rpt\frmBBS206.rpt"
    
    Crystal_Print CReport, strTmp, strRfile, strRptPath
    cmdPrint.Enabled = False
    medClearTable tblList
    Me.MousePointer = 0
End Sub

Private Sub cmdQuery_Click()
    Dim objSql          As clsGetSqlStatement
    Dim objTransReason  As clsQueryOrder
    Dim RS              As Recordset
    Dim strStatus       As String
    Dim strReason       As String
    Dim strFDt          As String
    Dim strTDt          As String
    Dim strPtid         As String
    Dim strWID          As String
    Dim strDID          As String
    Dim strDoct         As String
    Dim strTmp          As String
    Dim blnOpt          As Boolean
'상태 조회조건.
    Dim strQue          As String
    Dim ii              As Integer
    
    Dim Ord_Cnt         As Integer
    Dim Del_Cnt         As Integer
    Dim Ret_Cnt         As Integer
    Dim Exp_Cnt         As Integer
    Dim Ass_Cnt         As Integer
    Dim Can_Cnt As Integer
    
    Dim onlyComplete    As Boolean
    Dim onlyTest        As Boolean
    Dim objPro As clsProgress
    
    If txtWardID.Text = "" And txtDeptCd.Text = "" And txtDoct.Text = "" And txtPtid.Text = "" Then
        If MsgBox("병동, 진료과, 처방의, 환자ID 중 하나의 조건은 입력하셔야 합니다." & vbNewLine & vbNewLine & _
                  "조건을 입력하지 않는 경우 속도가 느릴 수 있습니다." & vbNewLine & "계속하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
'
'    If txtWardId.Text = "" And txtDeptcd.Text = "" And txtDoct.Text = "" And txtPtId.Text = "" Then
'        MsgBox "병동, 진료과, 처방의, 환자ID 중 하나의 조건은 입력하셔야 합니다.", vbExclamation
'        Exit Sub
'    End If
    
    Screen.MousePointer = vbHourglass
    
    Set objPro = New clsProgress
    
    With objPro
        .Container = Me
        .Left = tblList.Left
        .Top = tblList.Top
        .Width = tblList.Width
        .Height = .Height * 2
        .DisplayPercent = False
        .Message = "자료를 읽기 위해 준비중입니다..."
    End With
    
    Set objSql = New clsGetSqlStatement
    Set objTransReason = New clsQueryOrder
    
    For ii = 0 To 4
        If chkQue(ii).value = 1 Then
            strQue = strQue & ii & COL_DIV
        Else
            strQue = strQue & "" & COL_DIV
        End If
    Next ii
    
    strQue = Mid(strQue, 1, Len(strQue) - 1)
    
    If chkTot.value = 1 Then
        strQue = "0" & COL_DIV & "1" & COL_DIV & "2" & COL_DIV & "3" & COL_DIV & "4"
    End If
    
    '오직 완결처방만....

    If strQue = "4" Then onlyComplete = True
    If strQue = "3" Then onlyTest = True
    
    strPtid = txtPtid
    strWID = UCase(txtWardID)
    strDID = UCase(txtDeptCd)
    strDoct = txtDoct
    
    strFDt = Format(dtpFrom.value, PRESENTDATE_FORMAT)
    strTDt = Format(dtpTo.value, PRESENTDATE_FORMAT)
    
    '무조건 예정일자로...
'    blnOpt = True
    
    '2005/06/01 modify by legends
    '예수병원은 처방일과 예정일이 같으므로 속도개선을 위해서 처방일을 조건으로 처리하도록 변경함
    blnOpt = False
    
    '테이블 초기화
    tblList.MaxRows = 0
    With objSql
        Set RS = New Recordset
        
        DoEvents
        
        RS.Open .Order_Status_LIst(strFDt, strTDt, blnOpt, strQue, strPtid, strWID, _
                                                     strDID, strDoct), DBConn
    End With
    
    If RS Is Nothing Then GoTo SKIP2
    
    If RS.EOF = False Then
        objPro.DisplayPercent = True
        objPro.Message = "자료를 읽고 있습니다..."
        objPro.Max = RS.RecordCount
        
        With tblList
            ii = 0
            .ReDraw = False
            Do Until RS.EOF = True
                ii = ii + 1
                
                objPro.value = ii
                If onlyComplete = True Then
                    Ass_Cnt = Val(RS.Fields("assigncnt").value & "") - Val(RS.Fields("assigncancelcnt").value & "")
                    If RS.Fields("unitqty").value & "" > Ass_Cnt Then GoTo Skip
                    
                End If
                If onlyTest = True Then
                    Ass_Cnt = Val(RS.Fields("assigncnt").value & "") - Val(RS.Fields("assigncancelcnt").value & "")
                    If RS.Fields("unitqty").value & "" <= Ass_Cnt Then GoTo Skip
                End If
                
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                
                If strTmp <> RS.Fields("ptid").value & "" Then
                    Dim ObjABO As New clsABO
                    
                    .Col = TblColumn.tcPTID: .value = RS.Fields("ptid").value & ""
                    .Col = TblColumn.tcPTNM: .value = GetPtNm(RS.Fields("ptid").value & "")
                    ObjABO.PtId = RS.Fields("ptid").value & ""      '혈액형을 구하자.
                    ObjABO.GetABO
                    .Col = TblColumn.tcABO:  .value = ObjABO.ABO & ObjABO.Rh
                    
                    Set ObjABO = Nothing
                End If
                
                .Col = TblColumn.tcORDDT: .value = Format(RS.Fields("orddt").value & "", "####-##-##")
                .Col = TblColumn.tcORDNM: .value = RS.Fields("testnm").value & ""
                '처방수량
                .Col = TblColumn.tcUNIT:  Ord_Cnt = Val(RS.Fields("unitqty").value & ""): .value = Ord_Cnt

                
                .Col = TblColumn.tcREQDT: .value = Format(RS.Fields("reqdt").value & "", "####-##-##")
                '수혈사유
                strReason = objTransReason.GetTransReason(RS.Fields("ptid").value & "", _
                                                                  RS.Fields("orddt").value & "", _
                                                                  RS.Fields("ordno").value & "")
                .Col = TblColumn.TcTRANSRSN: .value = strReason
                .Col = TblColumn.tcWARD:     .value = RS.Fields("wardid").value & ""
                .Col = TblColumn.tcDEPT:     .value = RS.Fields("deptcd").value & ""
                .Col = TblColumn.tcDOCTNM:   .value = GetEmpNm(RS.Fields("orddoct").value & "")
                .Col = TblColumn.TcMESG:     .value = RS.Fields("mesg").value & ""
                '갯수구하기
                .Col = TblColumn.TcA_Cnt:    .value = Val(RS.Fields("assigncnt").value & ""):      Ass_Cnt = Val(.value)
                .Col = TblColumn.TcD_CNT:    Del_Cnt = Val(RS.Fields("deliverycnt").value & ""):   .value = Del_Cnt: .ForeColor = vbRed
                .Col = TblColumn.tcR_Cnt:    Ret_Cnt = Val(RS.Fields("retcnt").value & ""):        .value = Ret_Cnt
                .Col = TblColumn.tcE_Cnt:    Exp_Cnt = Val(RS.Fields("expcnt").value & ""):        .value = Exp_Cnt
                .Col = TblColumn.tcDCFG:    .value = IIf(RS.Fields("dcfg").value & "" = "1", "√", ""): .ForeColor = vbRed
                Can_Cnt = Val(RS.Fields("assigncancelcnt").value & "")
                'TRANS_REQUIRE_USED
                .Col = TblColumn.TcSTS
                If TRANS_REQUIRE_USED Then
                        Select Case RS.Fields("stscd").value & ""
                            Case BBSOrdStatus.stsORDER: 'BBSOrdStatus
                                .value = STS_NM_ORDER '"처방"
                                .ForeColor = DCM_Gray
                            Case BBSOrdStatus.stsCOLLECT:
                                .value = STS_NM_COLLECT '"채혈"
                                .ForeColor = DCM_Green
                            Case BBSOrdStatus.stsACCESS:
                                .value = STS_NM_ACCESS '"접수"
                                .ForeColor = DCM_Brown
                            Case BBSOrdStatus.stsINPROCESS:
                                If Ord_Cnt = (Ass_Cnt - Can_Cnt) Then '완료(처방=(어싸인-취소))
                                    If Del_Cnt >= 1 Then
'                                        If Ord_Cnt = (Del_Cnt - (Ret_Cnt + Exp_Cnt)) Then  '처방=실제출고 종결
                                        If Ord_Cnt = (Del_Cnt - Ret_Cnt) Then    '처방=실제출고 종결
                                            .value = STS_NM_END '"종결"
                                            .ForeColor = DCM_Title_Blue
                                        ElseIf Ord_Cnt = Del_Cnt And Del_Cnt = Exp_Cnt Then
                                            .value = STS_NM_END '"종결"
                                            .ForeColor = DCM_Title_Blue
                                        Else '출고가 진행중이지만 Assign은 모두되었으니 완료로 그냥 냅둠
                                            .value = STS_NM_DONE '"완료"
                                            .ForeColor = DCM_LightBlue
                                        End If
                                    Else '출고를 하나도 안했어도 완료
                                        .value = STS_NM_DONE '"완료"
                                        .ForeColor = DCM_LightBlue
                                    End If
                                Else '어싸인도 제대로 안했으니 당연히 진행중
                                    .value = STS_NM_INPROGRESS '"검사중"
                                    .ForeColor = DCM_LightRed
                                End If
                            
'                            '+ Del_Cnt - (Ret_Cnt + Exp_Cnt)
'                                If Ord_Cnt >= Ass_Cnt Then
'                                    '출고-반환=0 or 출고-폐기=0이면 검사중
'                                    If (Del_Cnt - Ret_Cnt = 0) Or (Del_Cnt - Exp_Cnt = 0) Then
'                                        .value = STS_NM_INPROGRESS '"검사중"
'                                        .ForeColor = DCM_LightRed
'                                    Else
'                                        .value = STS_NM_DONE '"완료"
'                                        .ForeColor = DCM_LightBlue
'                                    End If
'                                Else
'                                    .value = STS_NM_INPROGRESS '"검사중"
'                                    .ForeColor = DCM_LightRed
'                                End If
'                                '2005/05/31 modify by legends
'                                If Ord_Cnt = Del_Cnt - (Ret_Cnt + Exp_Cnt) Then
'                                    .value = STS_NM_END '"종결"
'                                    .ForeColor = DCM_Title_Blue
'                                End If
                            Case BBSOrdStatus.stsEnd
'                                If Ord_Cnt = Del_Cnt - (Ret_Cnt + Exp_Cnt) Then
                                If Ord_Cnt = Del_Cnt - Ret_Cnt Then
                                    .value = STS_NM_END '"종결"
                                ElseIf Ord_Cnt = Del_Cnt And Del_Cnt = Exp_Cnt Then
                                    .value = STS_NM_END '"종결"
                                Else
                                    .value = STS_NM_DONE '"완료"
                                End If
                                .ForeColor = BBSOrdStatusColor.cIrEND
                        End Select
                Else
                        Select Case RS.Fields("stscd").value & ""
                            Case BBSOrdStatus.stsORDER:
                                .value = STS_NM_ORDER '"처방"
                                .ForeColor = DCM_Gray
                            Case BBSOrdStatus.stsCOLLECT:
                                .value = STS_NM_COLLECT '"채혈"
                                .ForeColor = DCM_Green
                            Case BBSOrdStatus.stsACCESS:
                                .value = STS_NM_ACCESS '"접수"
                                .ForeColor = DCM_Brown
                            Case BBSOrdStatus.stsREQUEST:
'                                If Ord_Cnt = (Ass_Cnt - Can_Cnt) Then '완료(처방=(어싸인-취소))
'                                    If Del_Cnt >= 1 Then
'                                        If Ord_Cnt = (Del_Cnt - (Ret_Cnt + Exp_Cnt)) Then  '처방=실제출고 종결
                                If Ord_Cnt <= (Ass_Cnt - Can_Cnt) Then '완료(처방=(어싸인-취소))
                                    If Del_Cnt >= 1 Then
'                                        If Ord_Cnt = (Del_Cnt - (Ret_Cnt + Exp_Cnt)) Then  '처방=실제출고 종결
                                        If Ord_Cnt = (Del_Cnt - Ret_Cnt) Then     '처방=실제출고 종결
                                            .value = STS_NM_END '"종결"
                                            .ForeColor = DCM_Title_Blue
                                        ElseIf Ord_Cnt = Del_Cnt And Del_Cnt = Exp_Cnt Then
                                        
                                            .value = STS_NM_END '"종결"
                                            .ForeColor = DCM_Title_Blue
                                        Else '출고가 진행중이지만 Assign은 모두되었으니 완료로 그냥 냅둠
                                            .value = STS_NM_DONE '"완료"
                                            .ForeColor = DCM_LightBlue
                                        End If
                                    Else '출고를 하나도 안했어도 완료
                                        .value = STS_NM_DONE '"완료"
                                        .ForeColor = DCM_LightBlue
                                    End If
                                Else '어싸인도 제대로 안했으니 당연히 진행중
                                    .value = STS_NM_INPROGRESS '"검사중"
                                    .ForeColor = DCM_LightRed
                                End If
                                
'                            '+ Del_Cnt - (Ret_Cnt + Exp_Cnt)
'                                If Ord_Cnt >= Ass_Cnt Then
'                                    '출고-반환=0 or 출고-폐기=0이면 검사중
'                                    If (Del_Cnt - Ret_Cnt = 0) Or (Del_Cnt - Exp_Cnt = 0) Then
'                                        .value = STS_NM_INPROGRESS '"검사중"
'                                        .ForeColor = DCM_LightRed
'                                    Else
'                                        .value = STS_NM_DONE '"완료"
'                                        .ForeColor = DCM_LightBlue
'                                    End If
'                                Else
'                                    .value = STS_NM_INPROGRESS '"검사중"
'                                    .ForeColor = DCM_LightRed
'                                End If
'                                '2005/05/31 modify by legends
'                                If Ord_Cnt = Del_Cnt - (Ret_Cnt + Exp_Cnt) Then
'                                    .value = STS_NM_END '"종결"
'                                    .ForeColor = DCM_Title_Blue
'                                End If
                                'Assign후 반환이나 폐기가된 경우라면 상태를 이전으로 변경.
                            Case BBSOrdStatus.stsEnd
'                                If Ord_Cnt = Del_Cnt - (Ret_Cnt + Exp_Cnt) Then
                                If Ord_Cnt = Del_Cnt - Ret_Cnt Then
                                    .value = STS_NM_END '"종결"
                                ElseIf Ord_Cnt = Del_Cnt And Del_Cnt = Exp_Cnt Then
                                    .value = STS_NM_END '"종결"
                                Else
                                    .value = STS_NM_DONE '"완료"
                                End If
                                .ForeColor = BBSOrderStatusColor.cIrEND
                        End Select
                End If
                strTmp = RS.Fields("ptid").value & ""
Skip:
                objPro.value = ii
                
                RS.MoveNext
            Loop
            .ReDraw = True
            cmdPrint.Enabled = True
        End With
        Set objPro = Nothing
    Else
        MsgBox "조건에 맞는 자료가 없습니다.", vbCritical, Me.Caption
        cmdPrint.Enabled = False
    End If
SKIP2:
    Screen.MousePointer = vbDefault
    Set RS = Nothing
    Set objSql = Nothing
    Set objTransReason = Nothing
    Set objPro = Nothing
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    chkQue(0).Caption = STS_NM_ORDER
    chkQue(1).Caption = STS_NM_COLLECT
    chkQue(2).Caption = STS_NM_ACCESS
    chkQue(3).Caption = STS_NM_INPROGRESS
    chkQue(4).Caption = STS_NM_DONE
    
    Call cmdClear_Click
End Sub

Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    With tblList
        .Row = .ActiveRow
        .Col = TblColumn.TcMESG
        txtRemark = .value
    End With
End Sub

Private Sub txtDeptCd_Change()
    If lblDeptNm.Caption <> "" Then
        lblDeptNm.Caption = ""
        Call medClearTable(tblList)
        cmdPrint.Enabled = False
    End If
End Sub

Private Sub txtDeptCd_GotFocus()
'    txtDeptCd.tag = txtDeptCd.Text
    
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDeptcd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'Private Sub txtDeptcd_LostFocus()
'    If txtDeptCd.Text = "" Then
'        lblDeptNm.Caption = ""
'    Else
'        If txtDeptCd.tag = txtDeptCd.Text Then
'            Exit Sub
'        Else
'            Dim strDeptNm As String
'
'            strDeptNm = GetDeptNm(txtDeptCd.Text)
'            If strDeptNm = "" Then
'                txtDeptCd.Text = "": lblDeptNm.Caption = ""
'            Else
'                lblDeptNm.Caption = strDeptNm
'            End If
''            If ObjBBSComCode.DeptCd.Exists(txtDeptCd.Text) Then
''                ObjBBSComCode.DeptCd.KeyChange txtDeptCd.Text
''                lblDeptNm.Caption = ObjBBSComCode.DeptCd.Fields("deptnm")
''            Else
''                txtDeptCd.Text = "": lblDeptNm.Caption = ""
''            End If
'        End If
'    End If
''    cmdQuery.SetFocus
'End Sub

Private Sub txtDeptCd_Validate(Cancel As Boolean)
    Dim strDeptNm As String
    
    If txtDeptCd.Text = "" Then Exit Sub
    
    strDeptNm = GetDeptNm(txtDeptCd.Text)
    
    If strDeptNm = "" Then
        Cancel = True
        MsgBox "해당 부서코드가 존재하지 않습니다.", vbExclamation
    Else
        lblDeptNm.Caption = strDeptNm
    End If
        
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtDoct_Change()
    If lblDoctNm.Caption <> "" Then
        lblDoctNm.Caption = ""
        Call medClearTable(tblList)
        cmdPrint.Enabled = False
    End If
End Sub

Private Sub txtDoct_GotFocus()
'    txtDoct.tag = txtDoct
'    txtDoct.SelStart = 0
'    txtDoct.SelLength = Len(txtDoct)
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDoct_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'Private Sub txtDoct_LostFocus()
'    If txtDoct.Text = "" Then
'        lblDoctNm.Caption = ""
'    Else
'        If txtDoct.tag = txtDoct Then
'            Exit Sub
'        Else
'            lblDeptNm.Caption = GetEmpNm(txtDoct.Text)
'            If lblDeptNm.Caption = "" Then
'                txtDoct.Text = "": lblDoctNm.Caption = ""
'            End If
'        End If
'    End If
''    cmdQuery.SetFocus
'End Sub

Private Sub txtDoct_Validate(Cancel As Boolean)
    Dim strDoctNm As String
    
    If txtDoct.Text = "" Then Exit Sub
    
    strDoctNm = GetDoctNm(txtDoct.Text)
    
    If strDoctNm = "" Then
        Cancel = True
        MsgBox "해당 처방의가 존재하지 않습니다.", vbExclamation
    Else
        lblDoctNm.Caption = strDoctNm
    End If
        
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtPtid_Change()
    If lblPtNm.Caption <> "" Then
        lblPtNm.Caption = ""
        Call medClearTable(tblList)
        cmdPrint.Enabled = False
    End If
End Sub

Private Sub txtPtId_GotFocus()
'    txtPtID.tag = txtPtID
'    txtPtID.SelStart = 0
'    txtPtID.SelLength = Len(txtPtID.Text)
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPtid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'Private Sub txtPtId_LostFocus()
'    If txtPtID.Text = "" Then
'        lblPtNm.Caption = ""
'    Else
'        If txtPtID.tag = txtPtID.Text Then
'            Exit Sub
'        Else
'            Call GetOrderPatientName(txtPtID.Text)
'        End If
'    End If
''    cmdQuery.SetFocus
'End Sub

Private Function GetOrderPatientName(ByVal qPtid As String) As String
    Dim objSql      As clsGetSqlStatement
'    Dim strTmp      As String
    Dim strFromDt   As String
    Dim strToDt     As String
    
    Set objSql = New clsGetSqlStatement
    
    strToDt = Format(dtpTo.value, PRESENTDATE_FORMAT)
    strFromDt = Format(dtpFrom.value, PRESENTDATE_FORMAT)
    
    GetOrderPatientName = objSql.Get_OrderStatusPt(qPtid, strFromDt, strToDt)
    
'    If strTmp <> "" Then
'        txtPtID.Text = medGetP(strTmp, 1, COL_DIV)
'        lblPtNm.Caption = medGetP(strTmp, 2, COL_DIV)
'    Else
'        MsgBox "조건에 맞는 자료가 없습니다.", vbExclamation
'    End If
    
    Set objSql = Nothing
End Function

Private Sub txtPtID_Validate(Cancel As Boolean)
    Dim strTmp As String
    
    If txtPtid.Text = "" Then Exit Sub
    
    txtPtid.Text = Format(txtPtid.Text, String(BBS_PTID_LENGTH, "0"))
    
    strTmp = GetOrderPatientName(txtPtid.Text)
    If strTmp = "" Then
        Cancel = True
        MsgBox "해당 환자가 존재하지 않습니다.", vbExclamation
    Else
'        txtPtID.Text = medGetP(strTmp, 1, COL_DIV)
        lblPtNm.Caption = medGetP(strTmp, 2, COL_DIV)
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtWardID_Change()
    If lblWardNm.Caption <> "" Then
        lblWardNm.Caption = ""
        Call medClearTable(tblList)
        cmdPrint.Enabled = False
    End If
End Sub

Private Sub txtWardId_GotFocus()
'    txtWardID.tag = txtWardID
'    txtWardID.SelStart = 0
'    txtWardID.SelLength = Len(txtWardID)
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtWardID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

'Private Sub txtWardId_LostFocus()
'    If txtWardID = "" Then
'        lblWardNm.Caption = ""
'    Else
'        If txtWardID.tag = txtWardID.Text Then
'            Exit Sub
'        Else
'            Dim strWardNm As String
'            strWardNm = GetWardNm(txtWardID.Text)
'            If strWardNm = "" Then
'                txtWardID.Text = ""
'            Else
'                lblWardNm.Caption = strWardNm
'            End If
''            If ObjBBSComCode.wardid.Exists(txtWardID.Text) Then
''                ObjBBSComCode.wardid.KeyChange (txtWardID.Text)
''                txtWardID.Text = ObjBBSComCode.wardid.Field("wardid")
''                lblWardNm.Caption = ObjBBSComCode.wardid.Field("wardnm")
''            Else
''                txtWardID.Text = ""
''            End If
'        End If
'    End If
''    cmdQuery.SetFocus
'End Sub

Private Sub cmdPop_Click(Index As Integer)
    Dim objSql  As clsCrossMatching
    Dim ObjDic  As clsDictionary
    
    Dim strFromDt   As String
    Dim strToDt     As String
    
    Set objMyList = New clsPopUpList
    Set objSql = New clsCrossMatching
    
    strToDt = Format(dtpTo.value, PRESENTDATE_FORMAT)
    strFromDt = Format(dtpFrom.value, PRESENTDATE_FORMAT)
    
    
    With objMyList
        .FormCaption = "코드조회": .ColumnHeaderText = "코드;코드명"
        .Connection = DBConn
        
        Select Case Index
            Case 0
                Call .LoadPopUp(GetSQLWardList)
                
                txtWardID.Text = .SelectedItems(0)
                lblWardNm.Caption = .SelectedItems(1)
            Case 1
'                txtPtID.Text = "": lblPtNm.Caption = ""
'                Call .LoadPopUp(objSql.Get_OrdPt(strFROMDt, strToDt)) ', 2350, 7650, ObjDic)
'                If .SelectedString <> "" Then
'                    txtPtID.Text = .SelectedItems(0)
'                    lblPtNm.Caption = .SelectedItems(1)
'                End If
            Case 2
                Call .LoadPopUp(GetSQLDeptList)
                    
                txtDeptCd.Text = .SelectedItems(0)
                lblDeptNm.Caption = .SelectedItems(1)
            Case 3
                Call .LoadPopUp(GetSQLDoctList)
                
                txtDoct.Text = .SelectedItems(0)
                lblDoctNm.Caption = .SelectedItems(1)
        End Select
    End With
    
    Set ObjDic = Nothing
    Set objSql = Nothing
    Set objMyList = Nothing
    
End Sub

Private Sub Crystal_Print(ByVal CrystalNm As CrystalReport, ByVal strTmp As String, _
                            ByVal strFilePath As String, ByVal strRptPath As String)
    
    'CrystalNm:Crystal컨트롤 Name
    'strTmp: Record String(출력값)
    'strRptPath: Rpt파일 경로
    'strFilePath: text Fil 경로
    
    Dim intFNum As Integer
    
    intFNum = FreeFile
    Open strFilePath For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    With CrystalNm
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
        .Reset
    End With
End Sub

Private Sub txtWardID_Validate(Cancel As Boolean)
    Dim strWardNm As String
    
    If txtWardID.Text = "" Then Exit Sub
    
    strWardNm = GetWardNm(txtWardID.Text)
    
    If strWardNm = "" Then
        Cancel = True
        MsgBox "해당 병동코드가 존재하지 않습니다.", vbExclamation
    Else
        lblWardNm.Caption = strWardNm
    End If
        
    If Cancel Then SendKeys "{Home}+{End}"
End Sub
