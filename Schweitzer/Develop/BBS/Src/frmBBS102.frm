VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBBS102 
   BackColor       =   &H00DBE6E6&
   Caption         =   "수혈처방출력"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   Icon            =   "frmBBS102.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14580
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdOrderView 
      BackColor       =   &H00F4F0F2&
      Caption         =   "처방별조회(&C)"
      Height          =   510
      Left            =   7380
      Style           =   1  '그래픽
      TabIndex        =   48
      Top             =   8580
      Width           =   1500
   End
   Begin VB.CheckBox chkAutoPrint 
      BackColor       =   &H00800000&
      Caption         =   "수혈의뢰 자동출력"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   8280
      TabIndex        =   47
      Top             =   1530
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Enabled         =   0   'False
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   43
      Tag             =   "15101"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   42
      Tag             =   "124"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   41
      Tag             =   "128"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdCollect 
      BackColor       =   &H00F4F0F2&
      Caption         =   "접수(&O)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   40
      Tag             =   "15101"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.Frame fraStore 
      BorderStyle     =   0  '없음
      Height          =   2535
      Left            =   10260
      TabIndex        =   33
      Top             =   2580
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ListBox lstLeg 
         Height          =   1680
         Left            =   60
         TabIndex        =   36
         Top             =   420
         Width           =   1215
      End
      Begin VB.ListBox lstRow 
         Height          =   1680
         Left            =   1320
         TabIndex        =   35
         Top             =   420
         Width           =   1215
      End
      Begin VB.ListBox lstCol 
         Height          =   1680
         Left            =   2580
         TabIndex        =   34
         Top             =   420
         Width           =   1215
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   315
         Left            =   60
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   60
         Width           =   3735
         _ExtentX        =   6588
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
         Alignment       =   1
         Caption         =   "보관장소"
         Appearance      =   0
      End
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   1320
         TabIndex        =   39
         Top             =   2220
         Width           =   570
      End
      Begin VB.Label lblCancel 
         AutoSize        =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2100
         TabIndex        =   38
         Top             =   2220
         Width           =   705
      End
   End
   Begin VB.ComboBox cboLeg 
      Height          =   300
      ItemData        =   "frmBBS102.frx":000C
      Left            =   13125
      List            =   "frmBBS102.frx":000E
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   1455
      Width           =   990
   End
   Begin VB.CheckBox chkSPos 
      BackColor       =   &H00800000&
      Caption         =   "보관장소 자동부여"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   10350
      TabIndex        =   3
      Top             =   1545
      Width           =   2055
   End
   Begin VB.ComboBox cboBuilding 
      Height          =   300
      Left            =   11265
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   75
      Width           =   3210
   End
   Begin VB.CommandButton cmdRePrint 
      BackColor       =   &H00F4F0F2&
      Height          =   315
      Left            =   14145
      Picture         =   "frmBBS102.frx":0010
      Style           =   1  '그래픽
      TabIndex        =   0
      ToolTipText     =   "출고전표를 재발행합니다."
      Top             =   1455
      Width           =   315
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   255
      Index           =   1
      Left            =   12435
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1500
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   450
      BackColor       =   8388608
      ForeColor       =   16777215
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
      Caption         =   "Rack"
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   900
      Top             =   8430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   6690
      Left            =   75
      TabIndex        =   44
      Top             =   1785
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   11800
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
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
      MaxCols         =   43
      MaxRows         =   25
      OperationMode   =   1
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS102.frx":0542
      TextTip         =   4
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   315
      Left            =   75
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1455
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
      Caption         =   "  처방 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Width           =   14400
      _ExtentX        =   25400
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
      Caption         =   "  조회 조건"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1140
      Left            =   75
      TabIndex        =   6
      Top             =   315
      Width           =   14400
      Begin VB.CheckBox chkDc 
         BackColor       =   &H00DBE6E6&
         Caption         =   "DC제외"
         Height          =   240
         Left            =   11820
         TabIndex        =   27
         Top             =   720
         Value           =   1  '확인
         Width           =   930
      End
      Begin VB.ComboBox cboInOut 
         Height          =   300
         ItemData        =   "frmBBS102.frx":15F9
         Left            =   4545
         List            =   "frmBBS102.frx":1606
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   270
         Width           =   990
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   12945
         Style           =   1  '그래픽
         TabIndex        =   15
         Tag             =   "15101"
         Top             =   390
         Width           =   1320
      End
      Begin VB.TextBox txtWardId 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   5550
         TabIndex        =   14
         Text            =   "7123456"
         Top             =   270
         Width           =   1110
      End
      Begin VB.CommandButton cmdWardId 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   315
         Left            =   6675
         Style           =   1  '그래픽
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Width           =   360
      End
      Begin VB.TextBox txtPtId 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   9615
         TabIndex        =   12
         Text            =   "7123456"
         Top             =   270
         Width           =   1155
      End
      Begin VB.CommandButton cmdPtId 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   330
         Left            =   10800
         Style           =   1  '그래픽
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   360
      End
      Begin VB.CheckBox chkStat 
         BackColor       =   &H00DBE6E6&
         Caption         =   "응급처방만"
         Height          =   240
         Left            =   10560
         TabIndex        =   10
         Top             =   720
         Width           =   1230
      End
      Begin VB.ComboBox cboOrd 
         Height          =   300
         ItemData        =   "frmBBS102.frx":161C
         Left            =   1200
         List            =   "frmBBS102.frx":1626
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   660
         Width           =   3150
      End
      Begin VB.CheckBox chkTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         Height          =   240
         Left            =   4560
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   705
         Width           =   855
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   8535
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   1050
         _ExtentX        =   1852
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
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   330
         Left            =   1185
         TabIndex        =   17
         Top             =   285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   84017155
         CurrentDate     =   36838
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   330
         Left            =   2910
         TabIndex        =   18
         Top             =   285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   84017155
         CurrentDate     =   36838
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   7080
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
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
         Left            =   11190
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   270
         Width           =   1440
         _ExtentX        =   2540
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
         Index           =   0
         Left            =   105
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   285
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "예 정 일 자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   675
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "혈액제제별"
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   495
         Left            =   5580
         TabIndex        =   21
         Top             =   540
         Width           =   4935
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "완료"
            Height          =   255
            Index           =   4
            Left            =   3900
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   180
            Width           =   735
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "검사중"
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   180
            Width           =   855
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "접수"
            Height          =   255
            Index           =   2
            Left            =   1860
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   180
            Value           =   1  '확인
            Width           =   675
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   180
            Value           =   1  '확인
            Width           =   675
         End
         Begin VB.CheckBox chkQue 
            BackColor       =   &H00DBE6E6&
            Caption         =   "처방"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Width           =   675
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   2715
         TabIndex        =   32
         Top             =   345
         Width           =   135
      End
      Begin VB.Label lblAge 
         Height          =   195
         Left            =   11505
         TabIndex        =   31
         Top             =   180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblSex 
         Height          =   240
         Left            =   10725
         TabIndex        =   30
         Top             =   180
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "☞ 마우스오른쪽 버튼을 사용하시면 검체추가요청 및 검사장소변경 기능을 사용 가능."
      ForeColor       =   &H00854F3F&
      Height          =   180
      Left            =   75
      TabIndex        =   46
      Top             =   8775
      Width           =   6900
   End
End
Attribute VB_Name = "frmBBS102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    tcSEL = 1
    tcPTID
    tcPTNM
    tcABO
    tcORDNM
    
    tcORDDT
    tcUNITQTY
    TcMESG
    tcSTATnm
    tcDCNM
    
    tcSTSNM
    tcWARD
    tcROOM
    tcDEPT
    tcSPCNO
    
    tcSTORE
    tcACCNO
    tcCENTERNM
    tcBUSSDIV
    tcORDDTDB
    
    tcORDNO
    tcORDSEQ
    tcSTATFG
    tcDCFG
    tcBedInDT
    
    tCLegRowCol
    tcCENTERCD
    tcNOACCSSS
    tcPHERESIS
    tcSTSCD
    
    tcREASON
    tcDISEASE
    tcDISEASE2
    tcDISEASE3
    tcDISEASE4
    
    tcTime
    tcORDDIV
    tcDUPCHK
    tcREQDT
    tcDOCT
    
    tcTRANSDT
    tcACCDTTM
End Enum


Private WithEvents objListPop   As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1
Private WithEvents objPtInfo    As frmPtInfo
Attribute objPtInfo.VB_VarHelpID = -1
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1

Private Const MENU_ADD& = 1
Private Const MENU_SEP$ = 2
Private Const MENU_XM& = 3

Private Const RowHeight& = 12

Private aryLeg()
Private aryRow()
Private aryCol()
Private SortTF As Boolean

'Private Sub cboDateDiv_Click()
'    tblPtList.MaxRows = 0
'End Sub

Private Sub cboInOut_Click()
    If cboInOut.ListIndex = 0 Then
        txtWardId = ""
        lblWardNm.Caption = ""
        txtWardId.Enabled = False
        cmdWardId.Enabled = False
        
        txtWardId.BackColor = Me.BackColor
    Else
        txtWardId = ""
        lblWardNm.Caption = ""
        txtWardId.Enabled = True
        cmdWardId.Enabled = True
        
        txtWardId.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub cboInOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboOrd_Click()
    tblPtList.MaxRows = 0
End Sub

Private Sub chkTot_Click()
    chkQue(0).value = chkTot.value
    chkQue(1).value = chkTot.value
    chkQue(2).value = chkTot.value
    chkQue(3).value = chkTot.value
    chkQue(4).value = chkTot.value
End Sub

Private Sub cmdClear_Click()
    Call ClearAll
    dtpFrDt.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS102 = Nothing
End Sub

Private Sub cmdOrderView_Click()
    Dim i As Integer
    Dim pFrmName As String
    If Len(txtPtId.Text) < 2 Then GoTo End2Stop

    pFrmName = "frm401ResultView"
    
    medMain.lblSubMenu.Caption = "처방결과조회" 'medGetP(Button.Tag, 1, "(")
    
    frmLisReview.ButtonKey = "LIS155A" 'Button.Key
    frmLisReview.Ptid = txtPtId.Text
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

    Exit Sub

PermissionDenied:
   
'    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
End2Stop:
End Sub

Private Sub cmdPrint_Click()
    Me.MousePointer = 11
'    Call PrintTransReport
'    Call PrintIntionlize
'    Call PrintHeader_Trans("홍길동", "EM", "0010313", "M", "Dise", "A+", "Trans", "IM", "김철승", "임상")
    Call PrintOrderList
    Me.MousePointer = 0
End Sub

Private Sub cmdPtId_Click()
    objPtInfo.Show vbModal
End Sub

Private Sub cmdQuery_Click()
    cmdQuery.tag = "1"
    lblTitle.Caption = " 처방 리스트"

    If cboInOut.ListIndex = 1 Then
        If txtWardId = "" Then
            MsgBox "병동을 선택하십시요.", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    If cboInOut.ListIndex = 2 Then
        If txtWardId = "" Then
            MsgBox "진료과를 선택하십시요.", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Me.MousePointer = 11
    
    Call Query
    
'    Call SpreadCellBorder(tblPtList)

    Me.MousePointer = 0
    
    If tblPtList.MaxRows > 0 Then
        cmdPrint.Enabled = True
        tblPtList.SetFocus
    Else
        cmdPrint.Enabled = False
        MsgBox "해당자료가 없습니다", vbInformation, Me.Caption
        If cboInOut.ListIndex = 0 Then
        Else
            txtWardId.SetFocus
        End If
    End If
    '2001-11-30추가
    cmdCollect.Enabled = True

End Sub

Private Sub cmdRePrint_Click()
    Dim i As Long
    Dim strPtnm As String
    Dim StrWARD As String
    Dim strPtid As String
    Dim strDiease As String
    Dim strABO As String
    Dim strTrans As String
    Dim strDoct As String
    Dim strDept As String
    Dim strSexAge As String
    
    If tblPtList.MaxRows <= 0 Then
        MsgBox "먼저 처방내역 조회한 후 출력하세요.", vbExclamation
        Exit Sub
    End If
    
    '접수 이상의 status 인 경우에만 재출력 가능
    
    tblPtList.Col = TblColumn.tcSTSNM
    tblPtList.Row = tblPtList.ActiveRow
    
    If tblPtList.value = "" Then Exit Sub
    If tblPtList.Row < 1 Then Exit Sub
        
    If tblPtList.value = STS_NM_ORDER Or tblPtList.value = STS_NM_COLLECT Then '처방,채혈
        MsgBox "재발행 대상이 아닙니다. 접수이상의 상태인 경우에만 재발행할 수 있습니다.", vbExclamation
        Exit Sub
    End If
    
'    Call PrintDeliveryList(True)
    Call PrintTransList(CStr(tblPtList.ActiveRow))
End Sub

Private Sub cmdWardId_Click()
    
    Set objListPop = New clsPopUpList
    With objListPop
        txtWardId.Text = "": lblWardNm.Caption = ""
        .Connection = DBConn
        .Delimiter = ";"
        Select Case cboInOut.ListIndex
            Case 1
                .FormCaption = "병동 조회": .ColumnHeaderText = "코드;코드명"
                .LoadPopUp GetSQLWardList
            Case 2
                .FormCaption = "진료과조회": .ColumnHeaderText = "코드;코드명"
                .LoadPopUp GetSQLDeptList
        End Select
        
        If .SelectedString <> "" Then
            If txtWardId <> .SelectedItems(0) Then
                tblPtList.MaxRows = 0
            End If
            txtWardId.Text = .SelectedItems(0)
            lblWardNm.Caption = .SelectedItems(1)
            dtpFrDt.SetFocus
        Else
            txtWardId.SetFocus
        End If
    End With
    Set objListPop = Nothing
    
End Sub

Private Sub dtpFrDt_Change()
    tblPtList.MaxRows = 0
End Sub

Private Sub dtpFrDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpToDt_Change()
    tblPtList.MaxRows = 0
End Sub

Private Sub dtpToDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim objAccess   As clsBBSAccess
    Dim objBBSsql   As clsGetSqlStatement
    Dim RS          As Recordset
    Dim Rsord       As Recordset
    Dim ii          As Long
    
    Set objPtInfo = New frmPtInfo
    Set objAccess = New clsBBSAccess
    Set objBBSsql = New clsGetSqlStatement
    Set Rsord = objBBSsql.Get_CompoRecordSet
    
    chkQue(0).Caption = STS_NM_ORDER
    chkQue(1).Caption = STS_NM_COLLECT
    chkQue(2).Caption = STS_NM_ACCESS
    chkQue(3).Caption = STS_NM_INPROGRESS
    chkQue(4).Caption = STS_NM_DONE
    
    With objAccess
        Set RS = New Recordset
        
        RS.Open .Get_LegPos(ObjSysInfo.BuildingCd), DBConn
        
        If RS.EOF = False Then
            cboLeg.Clear
            cboLeg.AddItem ""
            Do Until RS.EOF = True
                cboLeg.AddItem RS.Fields("legcd").value & ""
                RS.MoveNext
            Loop
        End If
        If cboLeg.ListCount <> 0 Then cboLeg.ListIndex = 0
        
    End With
    
    '검사항목
    With Rsord
        cboOrd.Clear
        cboOrd.AddItem "전체혈액제제"
        For ii = 1 To .RecordCount
             cboOrd.AddItem .Fields("compocd").value & "" & Space(2) & .Fields("abbrnm").value & ""
            .MoveNext
        Next ii
    End With
    
    '건물정보를 사용할 경우 건물리스트 로드
    If ObjSysInfo.UseBuildingInfo Then
        cboBuilding.Visible = True
        Call LoadBuilding
    Else
        cboBuilding.Visible = False
    End If
    
    dtpFrDt = DateAdd("d", -3, GetSystemDate)
    dtpToDt = GetSystemDate
    
    cboInOut.ListIndex = 0
    chkStat.value = False
    Call ClearAll
    Me.Show
    
    Set RS = Nothing
    Set Rsord = Nothing
    Set objAccess = Nothing
    Set objBBSsql = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objPtInfo = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub lblApply_Click()
    Dim LegCd   As String
    Dim RowNo   As String
    Dim ColNo   As String
    Dim store   As String
    
    Dim i       As Long
    Dim Row     As Long
    Dim spcno   As String
    
    If lstLeg.ListIndex < 0 Then
        Exit Sub
    ElseIf lstLeg.ListIndex > 0 Then
        If lstRow.ListIndex < 0 Then Exit Sub
        If lstCol.ListIndex < 0 Then Exit Sub
    End If
    
    If lstLeg.ListIndex = 0 Then
        LegCd = ""
        RowNo = ""
        ColNo = ""
        store = ""
    Else
        LegCd = lstLeg.Text
        RowNo = lstRow.Text
        ColNo = lstCol.Text
        store = LegCd & "(" & RowNo & "," & ColNo & ")"
    End If
    
    '----------이 보관장소를 다른 검체번호에 지정해놨을까?
    If store <> "" Then
        With tblPtList
            .Row = Row
            .Col = TblColumn.tcSPCNO: spcno = .value
            
            For i = 1 To .MaxRows
                .Row = i
                .Col = TblColumn.tcSPCNO
                If spcno <> .value Then
                    .Col = TblColumn.tcSTORE
                    If store = .value Then
                        MsgBox "이미 보관중이거나 보관대기중인 장소입니다.", vbCritical, Me.Caption
                        Exit Sub
                    End If
                End If
            Next i
        End With
    End If
    
    '----------반영(같은 검체번호이면 보관장소도 같다)
    Row = Val(fraStore.tag)
    
    With tblPtList
        .Row = Row
        .Col = TblColumn.tcSTORE:     .value = store
                                      .ForeColor = vbBlue
        .Col = TblColumn.tCLegRowCol: .value = LegCd & ";" & RowNo & ";" & ColNo
        
        .Col = TblColumn.tcSPCNO:     spcno = .value
        
        For i = 1 To .MaxRows
            If i <> Row Then
                .Row = i
                .Col = TblColumn.tcSPCNO
                If .value = spcno Then
                    '같은 검체번호다. 쓰자......
                    .Col = TblColumn.tcSTORE:     .value = store
                                                  .ForeColor = vbBlue
                    .Col = TblColumn.tCLegRowCol: .value = LegCd & ";" & RowNo & ";" & ColNo
                End If
            End If
        Next i
    End With
    
    fraStore.Visible = False
End Sub

Private Sub lblCancel_Click()
    fraStore.Visible = False
End Sub

Private Sub lstLeg_Click()
    Dim i       As Long
    Dim LegCd   As String
    Dim objXM   As clsCrossMatching
    Dim DrRS    As Recordset
    
    lstRow.Clear
    lstCol.Clear
    
    If lstLeg.ListIndex = 0 Then Exit Sub
    
    LegCd = lstLeg.Text
    
    Set objXM = New clsCrossMatching
    
    Set DrRS = New Recordset
    DrRS.Open objXM.Get_Row(LegCd, ObjSysInfo.BuildingCd), DBConn
    
    With DrRS
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                lstRow.AddItem .Fields("rowno").value & ""
                .MoveNext
            Next i
        End If
    End With
    Set DrRS = Nothing
    
    Set DrRS = New Recordset
    DrRS.Open objXM.Get_Col(LegCd, ObjSysInfo.BuildingCd), DBConn
    With DrRS
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                lstCol.AddItem .Fields("colno").value & ""
                .MoveNext
            Next i
        End If
    End With
    Set DrRS = Nothing
    
    Set objXM = Nothing
End Sub

'Private Sub mnuAddSpc_Click()
'
'    With tblPtList
'        .Row = .ActiveRow
'        .Col = TblColumn.tcACCNO
'        frmBBS204.txtAccNo = .value
'        frmBBS204.Show
'    End With
'End Sub

'Private Sub mnuMoveLoc_Click()
''2001-11-29 추가
'    Dim objBg       As clsBeginTrans
'    Dim Resp        As VbMsgBoxResult
'    Dim strSpcNo    As String
'    Dim strSQL      As String
'
'
'    Resp = MsgBox("해당 환자의 검사를 " & ObjSysInfo.BuildingNm & " 검사실에서 수행하시겠습니까?", vbQuestion + vbYesNo, "검사장소변경")
'    If Resp = vbNo Then Exit Sub
'
'    tblPtList.Col = TblColumn.tcSPCNO
'    strSpcNo = tblPtList.value
'
'    Set objBg = New clsBeginTrans
'    strSQL = objBg.Change_Location(medGetP(strSpcNo, 1, "-"), medGetP(strSpcNo, 2, "-"), _
'                                         ObjSysInfo.BuildingCd)
'    Set objBg = Nothing
'
'On Error GoTo Err_Trap
'    DBConn.BeginTrans
'
'    DBConn.Execute strSQL
'
'    DBConn.CommitTrans
'
'    Call Query
'
'    Exit Sub
'
'Err_Trap:
'    DBConn.RollbackTrans
'    MsgBox Err.Description, vbCritical, "오류"
'End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_ADD
            With tblPtList
                .Row = .ActiveRow
                .Col = TblColumn.tcACCNO
                frmBBS204.txtAccNo = .value
                frmBBS204.Show
            End With
        Case MENU_XM
            With tblPtList
                .Row = .ActiveRow
                .Col = TblColumn.tcACCNO
                DoEvents
                frmBBS201.Show
                frmBBS201.txtSpcNO.Text = Mid(.value, 3)
                frmBBS201.CallByExtForm
            End With
    End Select
        
End Sub

Private Sub objPtInfo_Click(ByVal isSELECT As Boolean, ByVal ptInfo As S2BBS_Library.clsPtInformation)
    txtPtId.Text = "": lblPtNm.Caption = ""
    On Error Resume Next
    If txtPtId.Text <> ptInfo.Ptid Then tblPtList.MaxRows = 0
    txtPtId.Text = ptInfo.Ptid
    lblPtNm.Caption = ptInfo.ptnm

End Sub

Private Function CanSelect(ByVal Col As Long, ByVal Row As Long) As Boolean
    
    Dim objSql   As clsQueryOrder
    Dim CenterCd As String
    Dim noaccess As String
    Dim pheresis As String
    Dim sel      As String
    Dim spcno    As String
    Dim KeepOur  As Long
    Dim i        As Long
    
    '중간에 나가면 불가능한 것이다.....
    CanSelect = False
    
    With tblPtList
        '검체번호가 있는 것만 대상
        '접수번호가 없는 것(처방미접수)만 대상
        '보관장소가 없는 것(검체미접수)만 대상
        'D/C처방은 제외
        '검체보관시간 지나지 않은것만 대상
        'irradiation 처방이 아닌 처방만 대상
        
        .Row = Row
        
        '건물코드가 다르면 접수할수 없다.
        .Col = TblColumn.tcCENTERCD: CenterCd = .value
        If CenterCd <> ObjSysInfo.BuildingCd Then Exit Function
        
        'D/C발생한 처방에 대해서는 접수할수 없다.
        .Col = TblColumn.tcDCFG
        If .value = "1" Then Exit Function
        
        '검체번호가 없으면 접수할수 없다.
        .Col = TblColumn.tcSPCNO
        If .value = "" Then Exit Function
        
        '접수번호가 있으면 접수할수 없다.
        .Col = TblColumn.tcACCNO
        If .value <> "" Then Exit Function
        
        '상태가 처방인것은 접수할수 없다.
        .Col = TblColumn.tcSTSNM
        If .value = STS_NM_ORDER Then Exit Function '"처방"
        
        '72시간이 지난 검체는 접수할수 없다.
'        .Col = TblColumn.tcTime
'        If Val(.value) > KeepOur Then Exit Function
        
        'IRRAdiation 처방은 접수할수 없다.
        .Col = TblColumn.tcORDDIV
        If .value = "Z" Then Exit Function
    End With

    CanSelect = True
End Function
Private Sub SPreadSort(ByVal Col As Integer)
    With tblPtList
        .ReDraw = False
        .SortBy = SortByRow
        .SortKey(1) = Col
        If SortTF = True Then
            .SortKeyOrder(1) = SortKeyOrderAscending
            SortTF = False
        Else
            SortTF = True
            .SortKeyOrder(1) = SortKeyOrderDescending
        End If
        .Col = 1:  .COL2 = .MaxCols
        .Row = 1:  .Row2 = .MaxRows
        .BlockMode = True
        .Action = 25
        .BlockMode = False
        .ReDraw = True
    End With
End Sub
Private Sub tblPtList_Click(ByVal Col As Long, ByVal Row As Long)
    Static BfRow    As Long
    Dim clrBackOdd  As Long
    Dim clrForeOdd  As Long
    Dim clrBackEven As Long
    Dim clrForeEven As Long
    
    Dim CenterCd    As String
    Dim noaccess    As String
    Dim pheresis    As String
    Dim sel         As String
    Dim spcno       As String
    Dim i           As Long
    
    If Row < 1 Then
        Call SPreadSort(Col)
        Exit Sub
    End If
    If Row > tblPtList.MaxRows Then Exit Sub
    If fraStore.Visible = True Then Exit Sub
        
    With tblPtList
    
        Call .GetOddEvenRowColor(clrBackOdd, clrForeOdd, clrBackEven, clrForeEven)
        
        If BfRow <> Row Then
            .Row = BfRow: .Row2 = BfRow
            .Col = 1: .COL2 = .MaxCols
            .BlockMode = True
            If (BfRow Mod 2) = 0 Then
                .BackColor = clrBackEven
            Else
                .BackColor = clrBackOdd
            End If
            .BlockMode = False
        End If
        
        .Row = Row: .Row2 = Row
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .BackColor = .SelBackColor
        .BlockMode = False
        
        BfRow = Row
    End With
    
    
    With tblPtList
        Select Case Col
            Case TblColumn.tcSTORE
                If chkSPos.value = 1 Then Exit Sub
                .Row = Row
                .Col = TblColumn.tcNOACCSSS: noaccess = .value
                .Col = TblColumn.tcCENTERCD: CenterCd = .value
                
                '-------------------아직 검체접수가 안된 것만 처리.
                If noaccess = "0" Then Exit Sub
                '---------------------우리 센터에서 처리할 수 없다.
                If CenterCd <> ObjSysInfo.BuildingCd Then Exit Sub
                
                fraStore.tag = Row
                fraStore.Visible = True
            Case TblColumn.tcSEL
                .Col = Col
                .Row = Row
                If .CellType <> CellTypeCheckBox Then Exit Sub
                
                If CanSelect(Col, Row) = False Then
                    .Col = Col
                    .Row = Row
                    .value = 0
                    Exit Sub
                End If
                
                'pheresis 처방일경우는 처방한건당체크가 가능하다.....
                .Row = Row
                .Col = TblColumn.tcSPCNO: spcno = .value
                .Col = TblColumn.tcSEL:   sel = .value
                .value = IIf(sel = 1, 0, 1)
'                If pheresis <> "1" Then

                For i = 1 To .MaxRows
                    If i <> Row Then
                        .Row = i
                        .Col = TblColumn.tcORDDIV
                        'irradiation처방인것을 구분하기 위해서......
                        If .value = C_WORKAREA Then
                            .Col = TblColumn.tcSPCNO
                            '같은 채혈번호를 가질때...
                            If spcno = .value Then
                                '접수번호가 ""(접수않된거만)....
                                .Col = TblColumn.tcACCNO
                                If .value = "" Then
                                    '접수처리가 가능한 혈액에 대해서만....
                                    .Col = Col
                                    If .CellType = CellTypeCheckBox Then
                                        .Col = TblColumn.tcSEL
                                        .value = IIf(sel = 1, 0, 1)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
'                End If
        End Select
    End With
End Sub

Private Sub tblPtList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Then Exit Sub
    With tblPtList
        .Row = Row
        .Col = TblColumn.tcACCNO
        If .value = "" Then Exit Sub
        .Col = TblColumn.tcSTSNM
        If .value = STS_NM_DONE Or .value = STS_NM_END Then Exit Sub '"완료","종결"
        .Action = ActionActiveCell
        
        Set objPop = New clsPopupMenu
        With objPop
            .AddMenu MENU_ADD, "검체추가요청"
            .AddMenu MENU_SEP, "-"
            .AddMenu MENU_XM, "XM 결과등록"
            
            .PopupMenus Me.hwnd
        End With
        Set objPop = Nothing
'
'
'        Set mnuPopup = frmControl.mnuPopup
'        Set mnuAddSpc = frmControl.mnuSub
'        mnuAddSpc.Caption = "검체추가요청"
'        PopupMenu mnuPopup
'        Set mnuPopup = Nothing
'        Set mnuAddSpc = Nothing
    End With
End Sub
Private Function GetTestInformation(ByVal sPtid As String) As String
    Dim objSql As clsCrossMatching
    Dim RS     As Recordset
    Dim strTmp As String
    Dim SSQL   As String
    Dim ii     As Integer
    
    Set objSql = New clsCrossMatching
    SSQL = objSql.TestResultXM(sPtid)
    If SSQL <> "" Then
    Set RS = New Recordset
    RS.Open SSQL, DBConn
        If Not RS.EOF Then
             Do Until RS.EOF
                 strTmp = strTmp & RS.Fields("workarea").value & "" & "-" & _
                          RS.Fields("accdt").value & "" & "-" & _
                          RS.Fields("accseq").value & "" & _
                          "    " & RS.Fields("abbrnm10").value & "" & " : " & _
                          RS.Fields("rstcd").value & "" & vbNewLine & "       "
                RS.MoveNext
            Loop
        End If
        Set RS = Nothing
    End If
    
    If strTmp <> "" Then
        strTmp = "  ★ 관련검사 ★ " & vbNewLine & "       " & strTmp
        GetTestInformation = strTmp
    End If
    
    Set objSql = Nothing
End Function
Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim objQuery    As clsQueryOrder
    Dim objDisease  As clsDisease
    Dim RS          As Recordset
'    Dim blnComplete As Boolean
    Dim intord      As Integer
    
    Dim strAccNo     As String  '접수번호
    Dim strSpcNo     As String  '검체번호
    Dim strStore     As String  '보관장소
    Dim StrWARD      As String  '병동
    Dim strDept      As String  '진료과
    Dim strReason    As String  '수혈사유
    Dim strDisea1    As String  '진단명
    Dim strDisea2    As String  '진단명2
    Dim strDisea3    As String  '진단명3
    Dim strDisea4    As String  '진단명4
    Dim coldttm      As String  '경과시간을 가지고 오기위한 변수
    Dim strTime      As String
    Dim strDiseaDisp As String
    Dim strReqDt     As String
    Dim strAccdttm As String
    Dim strMesg      As String
    
    Dim strAccDt    As String
    Dim strAccSeq   As String
    
    'IRRADIATION처방인경우..
    Dim strPtid      As String
    Dim strOrdDt     As String
    Dim strOrdNo     As String
    Dim strROrd      As String
    
    Dim i            As Long
    Dim strtip       As String
    Dim sICSStr         As String
    Dim strTmp          As String
    
    
    Dim blnCompleted As Boolean '완료여부
    Dim blnAccomplished As Boolean '종결여부
    
    If Row < 1 Then Exit Sub
    
    
    Set objQuery = New clsQueryOrder
    Set objDisease = New clsDisease
    
    With tblPtList
        Call .SetTextTipAppearance("굴림체", 9, False, False, &HFFFFC0, vbBlack)
        .Row = Row
        .Col = TblColumn.tcPTID:        strPtid = .value
        .Col = TblColumn.tcACCNO:       strAccNo = .value
        .Col = TblColumn.tcSPCNO:       strSpcNo = .value
        .Col = TblColumn.tcSTORE:       strStore = .value
        .Col = TblColumn.tcWARD:        StrWARD = .value
        .Col = TblColumn.tcDEPT:        strDept = .value
        .Col = TblColumn.tcREQDT:       strReqDt = .Text
        .Col = TblColumn.tcACCDTTM: strAccdttm = .Text
        .Col = TblColumn.TcMESG:        strMesg = .value
        .Col = TblColumn.tcORDDT:       strOrdDt = Replace(.value, "-", "")
        .Col = TblColumn.tcORDNO:       strOrdNo = .value
        '진단명을 구한다.
        objDisease.Clear
        objDisease.Ptid = strPtid
        objDisease.OrdDt = strOrdDt
        objDisease.ordno = strOrdNo
        
        If objDisease.GetDisease Then
            i = 0
            Do
                If objDisease.EOF Then Exit Do

                If objDisease.DiseaseCd <> "" Then
                    i = i + 1
                    Select Case i
                        Case 1: strDisea1 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 2: strDisea2 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 3: strDisea3 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                        Case 4: strDisea4 = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                    End Select
                End If
                objDisease.MoveNext
            Loop
        End If
        
        strDiseaDisp = strDisea1
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea2
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea3
        If strDisea2 <> "" Then strDiseaDisp = strDiseaDisp & vbNewLine & _
                                               "             " & strDisea4
                                               
        '수혈사유
        strReason = objQuery.GetTransReason(strPtid, strOrdDt, strOrdNo): If strReason = "" Then strReason = "(없음)"
        
        '----------------------------
        '검체경과 시간을 구하기위해서
        '----------------------------
        If strSpcNo <> "-" Then
            Set RS = New Recordset
            RS.Open objQuery.Get_spcTime(medGetP(strSpcNo, 1, "-"), medGetP(strSpcNo, 2, "-")), DBConn
            If Not RS.EOF Then
                If Len(RS.Fields("coltm").value & "") = 4 Then
                    coldttm = RS.Fields("coltm").value & "" & "00"
                    coldttm = Format(RS.Fields("coldt").value & "", "0###-##-##") & " " & Format(coldttm, "0#:##:##")
                Else
                    coldttm = Format(RS.Fields("coldt").value & "", "0###-##-##") & " " & Format(RS.Fields("coltm").value & "", "0#:##:##")
                End If
                strTime = DateDiff("h", coldttm, GetSystemDate) & "시간"
            End If
            Set RS = Nothing
        End If
        
        .Col = TblColumn.tcORDDIV
        '-----------------------------------------------
        'irradiation 처방인경우 검사중인 처방도 보여준다
        '-----------------------------------------------
        If .value = "Z" Then
            
            Set RS = objQuery.GetRelationOrder(strPtid, strOrdDt)
            If Not RS.EOF Then
                With RS
                    Do Until RS.EOF
                        intord = intord + 1
                    '검사중인거....
                        If .Fields("stscd").value & "" = "3" Then
                            Call CheckCompleted(.Fields("accdt").value & "", .Fields("accseq").value & "", .Fields("unitqty").value & "", _
                                                blnCompleted, blnAccomplished)
'                            blnComplete = CompleteOrderChk(.Fields("accdt").value & "", _
'                                                           .Fields("accseq").value & "", _
'                                                           .Fields("unitqty").value & "")
                            If intord <= 1 Then
                                If blnCompleted = False Then
                                    strROrd = strROrd & "  관련처방 : " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_INPROGRESS & vbNewLine '검사중
                                Else
                                    If blnAccomplished Then
                                        strROrd = strROrd & "  관련처방 : " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_END & vbNewLine '종결"
                                    Else
                                        strROrd = strROrd & "  관련처방 : " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_DONE & vbNewLine '완료"
                                    End If
                                End If
                            Else
                                If blnCompleted = False Then
                                    strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_INPROGRESS & vbNewLine '검사중"
                                Else
                                    If blnAccomplished Then
                                        strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_END & vbNewLine '종결"
                                    Else
                                        strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_DONE & vbNewLine '완료"
                                    End If
                                End If
                            End If
                            
                        Else
                            If intord <= 1 Then
                                Select Case .Fields("stscd").value & ""
                                    Case "0": strROrd = strROrd & "  관련처방 : " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_ORDER & vbNewLine '처방"
                                    Case "1": strROrd = strROrd & "  관련처방 : " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_COLLECT & vbNewLine '채혈"
                                    Case "2": strROrd = strROrd & "  관련처방 : " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_ACCESS & vbNewLine '접수"
                                End Select
                            Else
                                Select Case .Fields("stscd").value & ""
                                    Case "0": strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_ORDER & vbNewLine '처방"
                                    Case "1": strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_COLLECT & vbNewLine '채혈"
                                    Case "2": strROrd = strROrd & "             " & .Fields("testnm").value & "" & "(수량:" & .Fields("unitqty").value & "" & ") ▶ " & STS_NM_ACCESS & vbNewLine '접수"
                                End Select
                            End If
                        End If
                        .MoveNext
                    Loop
                End With
            End If
        End If
        
        sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
        
        strtip = "  접수번호 : [" & strAccNo & "], 검체번호 : [" & strSpcNo & "], 보관장소 : [" & strStore & "]" & vbNewLine & "  경과시간 : " & strTime & vbNewLine & _
                 "  병동/과  : " & StrWARD & "/" & strDept '& vbNewLine & _
                 "  수혈사유 : " & strREASON & vbNewLine & _
                 "  예정일시 : " & strReqDt & vbNewLine & _
                 "  처방비고 : " & strMesg & vbNewLine & _
                 strDiseaDisp
        
        If strReason <> "" Then strtip = strtip & vbNewLine & "  수혈사유 : " & strReason
        If strReqDt <> "" Then strtip = strtip & vbNewLine & "  예정일시 : " & strReqDt
        If strAccdttm <> "" Then strtip = strtip & vbNewLine & "  접수일시 : " & strAccdttm
        If strMesg <> "" Then strtip = strtip & vbNewLine & "  처방비고 : " & strMesg
        If sICSStr <> "" Then strtip = strtip & vbNewLine & " 감염여부 : " & sICSStr
        
        If strDiseaDisp <> "" Then strtip = strtip & vbNewLine & "  진 단 명 : " & strDiseaDisp
        
        If strROrd <> "" Then strtip = strtip & vbNewLine & Mid(strROrd, 1, Len(strROrd) - 1)
        strtip = strtip & vbNewLine & objQuery.GetAccWorkLoad(strAccNo)
        
        '** 추가 X-Match 상세결과 By M.G.Choi 2007.11.14
        strtip = strtip & vbNewLine & DetailRst(medGetP(strAccNo, 1, "-"), medGetP(strAccNo, 2, "-"))
        
        strTmp = GetTestInformation(strPtid)
        If strTmp <> "" Then
            strtip = strtip & vbNewLine & strTmp
        End If
        
        TipWidth = 6500 '6350
        MultiLine = 1
        TipText = vbNewLine & strtip & vbNewLine
        ShowTip = True
    End With
    
    Set RS = Nothing
    Set objQuery = Nothing
    Set objDisease = Nothing
    
End Sub

Private Function DetailRst(ByVal pAccDt As String, ByVal pAccSeq As String) As String
    Dim strSQL      As String
    Dim RS          As New ADODB.Recordset
    Dim strTmp      As String
    Dim strS1       As String
    Dim strS2       As String
    Dim strS3       As String
    Dim strS4       As String
    
    strSQL = " select step1, step2, step3, step4 from " & T_BBS302 & _
             "  where workarea = 'B' " & _
             "    and accdt = " & DBS(pAccDt) & _
             "    and accseq = " & DBN(pAccSeq)
             
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        strS1 = "saline" & IIf(RS.Fields("step1").value & "" = "1", "(O)", "(X)")
        strS2 = "bovine" & IIf(RS.Fields("step2").value & "" = "1", "(O)", "(X)")
        strS3 = "37'C" & IIf(RS.Fields("step3").value & "" = "1", "(O)", "(X)")
        strS4 = "coombs" & IIf(RS.Fields("step4").value & "" = "1", "(O)", "(X)")
        
        strTmp = "  X-match : " & strS1 & "," & strS2 & "," & strS3 & "," & strS4
    End If
    
    RS.Close
    Set RS = Nothing
    
    DetailRst = strTmp
    
End Function

Private Sub txtPtId_GotFocus()
    txtPtId.tag = txtPtId
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtId_LostFocus()
    If Screen.ActiveForm.ActiveControl.name = cmdClear.name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.name = cmdExit.name Then Exit Sub
    
    If txtPtId.tag = txtPtId Then Exit Sub
    If SearchPTINFO = False Then
        txtPtId.SetFocus
    Else
        txtPtId.tag = txtPtId.Text
    End If

End Sub

Private Function SearchPTINFO() As Boolean
    SearchPTINFO = Search_PtInfo
    tblPtList.MaxRows = 0
End Function

Private Sub txtWardId_GotFocus()
    txtWardId.tag = txtWardId
    txtWardId.SelStart = 0
    txtWardId.SelLength = Len(txtWardId)
End Sub

Private Sub txtWardId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If SearchWard = True Then
            txtWardId.tag = txtWardId
            SendKeys "{TAB}"
        Else
            txtWardId.SelStart = 0
            txtWardId.SelLength = Len(txtWardId)
        End If
    End If
End Sub

Private Sub txtWardId_LostFocus()
    If Screen.ActiveForm.ActiveControl.name = cmdClear.name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.name = cmdExit.name Then Exit Sub
    
    If txtWardId.tag = txtWardId Then Exit Sub
    If SearchWard = False Then txtWardId.SetFocus
End Sub

Private Function SearchWard() As Boolean

    SearchWard = Search_Ward
    
    tblPtList.MaxRows = 0
End Function

Private Sub ClearAll()
    Call ICSPatientMark
    txtWardId = ""
    lblWardNm.Caption = ""
    txtPtId = ""
    lblPtNm.Caption = ""
    tblPtList.MaxRows = 0
    chkSPos.value = 1
    cboOrd.ListIndex = 0
End Sub

Private Function Search_PtInfo() As Boolean
    Dim objPtInfo As clsPtInformation
    Dim DrRS      As Recordset
    Dim ii        As Long
    Dim strLng    As String
    
    If txtPtId = "" Then
        lblPtNm.Caption = ""
        Search_PtInfo = True
    Else
        For ii = 1 To Val(BBS_PTID_LENGTH) - 1
            strLng = strLng & "0"
        Next ii
        

        If Len(Trim(txtPtId.Text)) <> BBS_PTID_LENGTH Then
            txtPtId.Text = Format(txtPtId.Text, strLng & "#")
        End If
        
        '감염관리
        Call ICSPatientMark(txtPtId.Text, enICSNum.BBS_ALL)
        
        Set objPtInfo = New clsPtInformation
        Set DrRS = New Recordset
        DrRS.Open objPtInfo.Get_Ptid(txtPtId), DBConn
        
        If DrRS.EOF = False Then
            With objPtInfo
                .BedPt_Chk txtPtId.Text, Format(GetSystemDate, PRESENTDATE_FORMAT)
                If .PtDiv = "BED" Then
                    'txtPtId = .ptid
                    lblPtNm.Caption = .ptnm
                    lblSex = .Sex
                    lblAge = .Age
                Else
                    'txtPtId = .ptid
                    lblPtNm.Caption = .ptnm
                    lblSex = .Sex
                    lblAge = .Age
                End If
            End With
            Search_PtInfo = True
        Else
            MsgBox "해당되는 환자가 없습니다. 확인후 조회하세요.", vbInformation + vbOKOnly, Me.Caption
            txtPtId = ""
            lblPtNm.Caption = ""
            Search_PtInfo = False
        End If
        Set DrRS = Nothing
        Set objPtInfo = Nothing
    End If
End Function

Private Function Search_Ward() As Boolean
    If txtWardId = "" Then
        lblWardNm.Caption = ""
        Search_Ward = True
    Else
        txtWardId.Text = UCase(txtWardId.Text)
        lblWardNm.Caption = GetWardNm(txtWardId.Text)
        If lblWardNm.Caption = "" Then
            MsgBox "해당되는 자료가 없습니다. 확인후 입력하세요.", vbInformation + vbOKOnly, "병동입력"
            lblWardNm.Caption = ""
            Search_Ward = False
        End If
    End If
End Function

Private Sub CheckCompleted(ByVal vAccdt As String, ByVal vAccseq As String, ByVal vUnitqty As Long, _
                           ByRef pCompleted As Boolean, ByRef pAccomplished As Boolean)
'2005/05/31 modify by legends
'완료여부와 종결여부를 구하기 위한 루틴
'완료 : 처방 수량 만큼 준비되어 있는 경우
'종결 : 처방 수량 만큰 출고된 경우(반환하면 출고아님으로 간주)

    Dim objXM As clsCrossMatching
    Dim A_Cnt As Long   'Assign수량
    Dim C_Cnt As Long   'Assign Cancel 수량
    Dim O_Cnt As Long   '출고수량
    Dim R_Cnt As Long   '반환수량
    Dim X_Cnt As Long   '폐기수량
    Dim T_Cnt As Long   '총Assign 수량
    Dim M_Cnt As Long   '총 출고된 수량

    'pCompleted : Assign이 완료되었는지 여부
    'pAccomplished : 출고가 완료되었는지 여부

    'CompleteOrderChk=True이면 완결처방
    'CompleteOrderChk=미완결처방
    Set objXM = New clsCrossMatching
    
    pCompleted = False
    pAccomplished = False
    
    If vAccdt <> "" Then
        With objXM
            .Assign_Cnt vAccdt, Val(vAccseq)
            A_Cnt = .AssignCnt
            C_Cnt = .CancelCnt
            O_Cnt = .OutCnt
            R_Cnt = .RetCnt
            X_Cnt = .ExpCnt
        End With
        Set objXM = Nothing
        
        '출고갯수와 상관없이 처방수량과, Assign 수량을 비교한다.
        '총Assign 수량=Assign수량-Assign취소 수량
        
        T_Cnt = A_Cnt - C_Cnt '실제 Assign된 량 모두 Assign되었으면 완료
        M_Cnt = O_Cnt - (R_Cnt + X_Cnt) '출고된 수량-(반환된 수량+폐기된 수량)'실제 출고량
        
        '출고는 하나도 안하고 어싸인만 했다가 모두 어싸인 취소하면 접수상태로 롤백...
        
        '모두 출고했다가 폐기되었을 경우 종결로 표시(반환된 경우 제외)
        '처방=출고=폐기 인 경우 종결로 표시
        
        'vUnitqty : 처방수량
        '처방수량만큼 Assign이 되었으면 완료, 아니면 검사중
        If vUnitqty <= T_Cnt Then 'vUnitqty = T_Cnt
            If O_Cnt >= 1 Then '출고 액션이 한번이라도 된 경우
                If M_Cnt >= 1 Then '실제 출고가 한건 이상인 경우
                    pCompleted = True
                End If
            Else '출고가 하나도 안된 경우
                pCompleted = True
            End If
        Else
            pCompleted = False
        End If
        
'        If vUnitqty <= T_Cnt Then
'            pCompleted = True
'        End If
        
        If vUnitqty = M_Cnt Then
            pAccomplished = True
        End If
        
        '아래 조건이 추가되었음.2005/10/24
        If vUnitqty = O_Cnt And O_Cnt = X_Cnt Then
            pCompleted = True
            pAccomplished = True
        End If
    End If
    Set objXM = Nothing
End Sub

'Private Function CompleteOrderChk(ByVal accdt As String, ByVal accseq As String, ByVal unitqty As Long) As Boolean
'    Dim objXM As clsCrossMatching
'    Dim A_Cnt As Long   'Assign수량
'    Dim C_Cnt As Long   'Assign Cancel 수량
'    Dim O_Cnt As Long   '출고수량
'    Dim R_Cnt As Long   '반환수량
'    Dim X_Cnt As Long   '폐기수량
'    Dim T_Cnt As Long   '총Assign 수량
'
'
'    'CompleteOrderChk=True이면 완결처방
'    'CompleteOrderChk=미완결처방
'    Set objXM = New clsCrossMatching
'    CompleteOrderChk = False
'    If accdt <> "" Then
'
'        With objXM
'            .Assign_Cnt accdt, Val(accseq)
'            A_Cnt = .AssignCnt
'            C_Cnt = .CancelCnt
'            O_Cnt = .OutCnt
'            R_Cnt = .RetCnt
'            X_Cnt = .ExpCnt
'        End With
'        Set objXM = Nothing
'
'        '출고갯수와 상관없이 처방수량과, Assign 수량을 비교한다.
'        '총Assign 수량=Assign수량-Assign취소 수량
'
'        T_Cnt = A_Cnt - C_Cnt
'       ' T_Cnt = A_Cnt - C_Cnt - R_Cnt - X_Cnt
'
'        If unitqty <= T_Cnt Then
'            CompleteOrderChk = True
'        End If
'    End If
'    Set objXM = Nothing
'End Function

'Private Function CheckAccomplished(ByVal vAccdt As String, ByVal vAccseq As String, ByVal vUnitqty As Long) As Boolean
''2005/05/31 Append by legends
''완결 여부 체크
''처방수량과 실제 출고수량이 같은 경우 완결 처리
''출고 된 후 반환되지 않은 경우 수량.
'
'    Dim strSql As String
'    Dim Rs As Recordset
'
'    strSql = " select count(*) as cnt from " & T_BBS402
'    strSql = strSql & " where " & DBW("workarea=", "B")
'    strSql = strSql & " and " & DBW("accdt=", vAccdt)
'    strSql = strSql & " and " & DBW("accseq=", vAccseq)
'    strSql = strSql & " and (retfg<>'1' or retfg is not null)"
'
'    Set Rs = New Recordset
'    Rs.Open strSql, DBConn, , , adCmdText
'
'    If Rs.EOF Or Rs.BOF Then
'        CheckAccomplished = False
'    Else
'        CheckAccomplished = True
'    End If
'
'    Set Rs = Nothing
'End Function

Private Function IRR_DUPchk(ByVal Ptid As String, ByVal OrdDt As String) As Boolean
    Dim ii      As Integer
    Dim strTmp  As String
    
    strTmp = Ptid & COL_DIV & OrdDt
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcDUPCHK
            If .value = strTmp Then
                IRR_DUPchk = True
                Exit Function
            End If
        Next
    End With
End Function

Private Function GetABO(ByVal Ptid As String) As String
'혈액형,부작용,감염정보,상병코드,상병을 조회한다.
    Dim ObjABO As clsABO
    
    Set ObjABO = New clsABO
    With ObjABO
        .Ptid = Ptid
        If .GetABO = True Then
            GetABO = .ABO & .Rh
        Else
            GetABO = ""
        End If
    End With
    Set ObjABO = Nothing
    
End Function

Private Sub Query()
    Dim i           As Long
    Dim j           As Long
    
    Dim RS          As Recordset
    Dim QueryOrder  As clsQueryOrder
    
    Dim accno       As String
    Dim reason      As String
    Dim status      As String
    Dim spcno       As String
    Dim storeleg    As String
    Dim storerow    As String
    Dim storecol    As String
    Dim center      As String
    
    Dim inout       As String
    Dim MaxRowCnt   As Long
    Dim TestDiv     As String
'    Dim blnComplete As Boolean
    
    Dim objPrgBar   As clsProgress
    Dim objDisease  As clsDisease

    
    '윗줄과 같은내용이면 글자를 감추기 위한변수들
    Dim bkPtId      As String
    Dim bkReason    As String
    Dim bkReqDt     As String
    Dim bkOrdDt     As String
    Dim bkRoomid    As String
    Dim bkWard      As String
    Dim bkDept      As String
    
    Dim strDc       As String
    
    Dim blnCompleted As Boolean
    Dim blnAccomplished As Boolean
    
    tblPtList.MaxRows = 0
    
    Call Save_LegRowCol
    
    Set QueryOrder = New clsQueryOrder
    
    
    If cboOrd.ListIndex <> 0 Then TestDiv = medGetP(cboOrd.Text, 1, " ")
    '-----------
    '상태별 조회
    '-----------
    If chkTot.value Then
        '미완결만
        QueryOrder.stscd = "'0','1','2','3'"
        If TRANS_REQUIRE_USED = True Then QueryOrder.stscd = "'0','1','2','3','4'"
    Else
        'If chkAccess.value Then
            '처방
            If chkQue(0).value Then QueryOrder.stscd = "'0'"
            '채혈
            If chkQue(1).value Then
                If QueryOrder.stscd <> "" Then
                    QueryOrder.stscd = QueryOrder.stscd & ",'1'"
                Else
                    QueryOrder.stscd = "'1'"
                End If
            End If
            '접수
            If chkQue(2).value Then
                If TRANS_REQUIRE_USED Then
                    If QueryOrder.stscd <> "" Then
                        QueryOrder.stscd = QueryOrder.stscd & ",'2','3'"
                    Else
                        QueryOrder.stscd = "'2','3'"
                    End If
                Else
                    If QueryOrder.stscd <> "" Then
                        QueryOrder.stscd = QueryOrder.stscd & ",'2'"
                    Else
                        QueryOrder.stscd = "'2'"
                    End If
                End If
            End If
            '검사중
            If chkQue(3).value Then
                If QueryOrder.stscd <> "" Then
                    If TRANS_REQUIRE_USED Then
                        QueryOrder.stscd = QueryOrder.stscd & ",'3','4'"
                    Else
                        QueryOrder.stscd = QueryOrder.stscd & ",'3'"
                    End If
                Else
                    If TRANS_REQUIRE_USED Then
                        QueryOrder.stscd = "'3','4'"
                    Else
                        QueryOrder.stscd = "'3'"
                    End If
                End If
            End If
            '완결
            If chkQue(4).value Then
                If chkQue(3).value = False Then
                    If QueryOrder.stscd <> "" Then
                        If TRANS_REQUIRE_USED Then
                            QueryOrder.stscd = QueryOrder.stscd & ",'3','4'"
                        Else
                            QueryOrder.stscd = QueryOrder.stscd & ",'3'"
                        End If
                    Else
                        If TRANS_REQUIRE_USED Then
                            QueryOrder.stscd = "'3','4'"
                        Else
                            QueryOrder.stscd = "'3'"
                        End If
                    End If
                End If
            End If
    End If
    
    Select Case cboInOut.ListIndex
        Case 0: inout = ""
        Case 1: inout = "2"
        Case 2: inout = "1"
    End Select
    If chkDc.value = "1" Then strDc = "1"
    
    Set RS = QueryOrder.QueryOrder(Format(dtpFrDt, PRESENTDATE_FORMAT), Format(dtpToDt, PRESENTDATE_FORMAT), chkStat.value, txtPtId.Text, inout, strDc, txtWardId, TestDiv)
    
    If RS Is Nothing Then
        Set RS = Nothing
        Set QueryOrder = Nothing
        Exit Sub
    End If
    
    
    Set objPrgBar = New clsProgress
    objPrgBar.Container = medMain.stsBar
    
    objPrgBar.Min = 1
    objPrgBar.Max = RS.RecordCount
    
    
    With tblPtList
        bkPtId = ""
        .ReDraw = False
        For i = 1 To RS.RecordCount
        
            objPrgBar.value = i
            
            '건물정보를 가지고 온다.(검체정보도 같이)
            Call QueryOrder.GetSpcNoAndStore(RS.Fields("ptid").value & "", spcno, storeleg, storerow, storecol, center)
            
            '2001-11-23 추가 :
            '건물정보를 사용할 경우, 그리고 (전체)가 아닐경우 해당 건물의 데이타만 보여준다.
            '건물코드가 틀리면 건너뛴다.
            If center = "" Then center = ObjSysInfo.BuildingCd & vbTab & ObjSysInfo.BuildingNm
            

            If ObjSysInfo.UseBuildingInfo = 1 And cboBuilding.ListIndex <> 0 Then
                If medGetP(center, 1, vbTab) <> medGetP(cboBuilding.Text, 1, " ") Then: GoTo Skip
            End If
'            'X-Matching가 임상병리검사항목마스터에서도 존재하기에 인위적으로 빼줘야한다.
'            If (RS.Fields("workarea").value & "") <> "B" And (RS.Fields("workarea").value & "") <> "" Then GoTo SKIP
            
'            blnComplete = CompleteOrderChk(Rs.Fields("accdt").value & "", Rs.Fields("accseq").value & "", Rs.Fields("unitqty").value & "")
            Call CheckCompleted(RS.Fields("accdt").value & "", RS.Fields("accseq").value & "", RS.Fields("unitqty").value & "", _
                                blnCompleted, blnAccomplished)
            '검사중 or 완료 버튼 선택되어있을시....
            If chkQue(3).value Or chkQue(4).value Then
                '완료버튼만 선택되어있을시.....
                If chkQue(4).value And chkQue(3).value = 0 Then
                    If RS.Fields("orddiv").value & "" = "Z" Then GoTo Skip1
                    '처방,채혈,접수 조회시....
                    If RS.Fields("stscd").value & "" = "0" Or RS.Fields("stscd").value & "" = "1" Or RS.Fields("stscd").value & "" = "2" Then GoTo Skip1
                    '검사중인 처방은 skip
                    If blnCompleted = False Then GoTo Skip
                    '검사중버튼만 선택되어있을시...
                ElseIf chkQue(3).value And chkQue(4).value = 0 Then
                    '처방이 완료 되어있을시 skip......
                    If blnCompleted = True Then GoTo Skip
                    If .MaxRows >= 0 And RS.Fields("orddiv").value & "" = "Z" Then
                        If IRR_DUPchk(RS.Fields("ptid").value & "", RS.Fields("orddt").value & "") = False Then GoTo Skip
                    End If
                    '조건에 처방/채혈/접수가 선택되어있을시...
                    If RS.Fields("stscd").value & "" = "0" Or RS.Fields("stscd").value & "" = "1" Or RS.Fields("stscd").value & "" = "2" Then GoTo Skip1
                End If
            End If
Skip1:
            MaxRowCnt = MaxRowCnt + 1
            .MaxRows = MaxRowCnt: .RowHeight(-1) = RowHeight
            .Row = MaxRowCnt
            accno = Trim(RS.Fields("accdt").value & "") & "-" & Val(Trim(RS.Fields("accseq").value & ""))
            If accno = "-0" Then accno = "" 'accno = "미접수"
            
            .Col = TblColumn.tcACCNO:      .value = accno
            .Col = TblColumn.tcPTID:       .value = RS.Fields("ptid").value & ""
            .Col = TblColumn.tcPTNM:       .value = GetPtNm(RS.Fields("ptid").value & "")
            .Col = TblColumn.tcORDNM:      .value = RS.Fields("testnm").value & ""
            .Col = TblColumn.tcORDDT:      .value = Format(RS.Fields("orddt").value & "", "####-##-##")
            .Col = TblColumn.tcUNITQTY:    .value = RS.Fields("unitqty").value & ""
            .Col = TblColumn.tcREASON:     .value = Trim(Trim0(reason))
            .Col = TblColumn.tcREQDT:      .value = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
          
            '2001-11-30추가
            '출고전표에 담당의사/최근수혈일 출력하기위함
            .Col = TblColumn.tcDOCT:       .value = RS.Fields("orddoct").value & ""
            .Col = TblColumn.tcWARD:       .value = RS.Fields("wardid").value & ""
            .Col = TblColumn.tcROOM:       .value = RS.Fields("hosilid").value & ""
            .Col = TblColumn.tcDEPT:       .value = RS.Fields("deptcd").value & ""
            .Col = TblColumn.tcBUSSDIV:    .value = RS.Fields("bussdiv").value & ""
            .Col = TblColumn.tcORDDTDB:    .value = RS.Fields("orddt").value & ""
            .Col = TblColumn.tcORDNO:      .value = Val(RS.Fields("ordno").value & "")
            .Col = TblColumn.tcORDSEQ:     .value = Val(RS.Fields("ordseq").value & "")
            .Col = TblColumn.tcSTATFG:     .value = RS.Fields("statfg").value & ""
            .Col = TblColumn.tcSTATnm:     .value = IIf(RS.Fields("statfg").value & "" = "1", "Y", ""): .ForeColor = vbRed: .FontBold = True
            .Col = TblColumn.tcBedInDT:    .value = RS.Fields("bedindt").value & ""
            .Col = TblColumn.tcDCFG:       .value = RS.Fields("dcfg").value & ""
            .Col = TblColumn.tcDCNM:       .value = IIf(RS.Fields("dcfg").value & "" = "1", "Y", ""): .ForeColor = vbBlue: .FontBold = True
            .Col = TblColumn.tcPHERESIS:   .value = RS.Fields("testdiv").value & ""
            .Col = TblColumn.tcSTSCD:      .value = RS.Fields("stscd").value & ""
            .Col = TblColumn.tcSTSNM
                                            If TRANS_REQUIRE_USED Then
                                                    Select Case RS.Fields("stscd").value & ""
                                                         Case "0": .value = STS_NM_ORDER: .ForeColor = DCM_Gray '"처방"
                                                         Case "1": .value = STS_NM_COLLECT '"채혈"
                                                         Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"접수"
                                                         Case "3": .value = STS_NM_REQUEST: .ForeColor = DCM_Red '"요청"
                                                                   '출고했다 모두 반환하거나 어싸인했다 모두 어싸인 취소하면 검사중으로 표시...
                                                                   
                                                                   .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_REQUEST) '"종결","완료","요청"
                                                                   
                                                                   If .value = STS_NM_DONE Then .ForeColor = IIf(blnCompleted, &H8000&, DCM_Red) '"완료"
'                                                                   If .value = STS_NM_DONE Then .ForeColor = DCM_Red '"완료"
                                                                   If .value = STS_NM_END Then .ForeColor = DCM_Blue '"종결"
                                                         Case "4": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS) '"종결","완료","검사중"
                                                                   .ForeColor = IIf(blnCompleted, &H8000&, DCM_Brown)
                                                         Case Else: .value = ""
                                                    End Select
                                            Else
                                                    Select Case RS.Fields("stscd").value & ""
                                                         Case "0": .value = STS_NM_ORDER '"처방"
                                                         Case "1": .value = STS_NM_COLLECT: .ForeColor = DCM_LightRed '"채혈"
                                                         Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"접수"
                                                         Case "3": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS): .ForeColor = DCM_Brown '"종결","완료","검사중"
                                                                   If .value = STS_NM_DONE Then .ForeColor = DCM_Red '"완료"
                                                                   If .value = STS_NM_END Then .ForeColor = DCM_Blue '"종결"
                                                         Case Else: .value = ""
                                                    End Select
                                            End If
                                            
            .Col = TblColumn.TcMESG: .value = RS.Fields("mesg").value & ""
            

            '--------------------------------------------------------------------------------------
            .Col = TblColumn.tcCENTERNM:    .value = medGetP(center, 2, vbTab)
            .Col = TblColumn.tcCENTERCD:    .value = medGetP(center, 1, vbTab)
            
            '다른센터에있는 검체표기.
            If medGetP(center, 2, vbTab) <> ObjSysInfo.BuildingNm Then .Col = TblColumn.tcSTORE:   .value = medGetP(center, 1, vbTab)
            'Workarea표기
            .Col = TblColumn.tcORDDIV:      .value = RS.Fields("orddiv").value & ""
            '보관장소표기
            If .value = C_WORKAREA Then
                If storerow = "0" Then storerow = ""
                If storecol = "0" Then storecol = ""
                
                .Col = TblColumn.tCLegRowCol:   .value = storeleg & ";" & storerow & ";" & storecol
                .Col = TblColumn.tcSPCNO:       .value = spcno
                
                If spcno = "" Then
                    .Col = TblColumn.tcSTORE:   .value = "" '.value = "미채혈"
                Else
                    If storeleg = "" Then
                        .Col = TblColumn.tcSTORE:    .value = ""
                        .Col = TblColumn.tcNOACCSSS: .value = "1"
                    Else
                        .Col = TblColumn.tcSTORE:    .value = storeleg & "(" & storerow & "," & storecol & ")"
                        .Col = TblColumn.tcNOACCSSS: .value = "0"
                    End If
                End If
            End If
            
            
            .Col = TblColumn.tcDUPCHK: .value = RS.Fields("ptid").value & "" & COL_DIV & RS.Fields("orddt").value & ""
            .Col = TblColumn.tcTRANSDT: '.value = QueryOrder.GetLatestTrandDt(RS.Fields("ptid").value & "")
            .Col = TblColumn.tcACCDTTM: .value = IIf(RS.Fields("rcvdt").value & "" = "", "", Format(RS.Fields("rcvdt").value & "", "0###-##-##") & " " & Format(RS.Fields("rcvtm").value & "", "0#:##:##"))
            
            
            '진단명을 구한다.
            Set objDisease = Nothing
            Set objDisease = New clsDisease
            With objDisease
                .Clear
                .Ptid = RS.Fields("ptid").value & ""
                .OrdDt = RS.Fields("orddt").value & ""
                .ordno = RS.Fields("ordno").value & ""
            End With
            
            If objDisease.GetDisease = False Then
                .Col = TblColumn.tcDISEASE: .value = ""
                .Col = TblColumn.tcDISEASE2: .value = ""
                .Col = TblColumn.tcDISEASE3: .value = ""
                .Col = TblColumn.tcDISEASE4: .value = ""
            Else
                j = 0
                Do
                    If objDisease.EOF Then Exit Do
                    
                    If objDisease.DiseaseCd <> "" Then
                        j = j + 1
                        Select Case j
                            Case 1: .Col = TblColumn.tcDISEASE
                            Case 2: .Col = TblColumn.tcDISEASE2
                            Case 3: .Col = TblColumn.tcDISEASE3
                            Case 4: .Col = TblColumn.tcDISEASE4
                        End Select
                        .value = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                    End If
                    objDisease.MoveNext
                Loop
            End If
            Set objDisease = Nothing
            
            '-------------------------
            '중복되는 값은 안보이게...
            '-------------------------
            
            If bkPtId <> RS.Fields("ptid").value & "" Then
                bkPtId = RS.Fields("ptid").value & ""
                bkReason = reason
                bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
                bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                bkRoomid = RS.Fields("hosilid").value & ""
                bkWard = RS.Fields("wardid").value & ""
                bkDept = RS.Fields("deptcd").value & ""
                
            Else
                .Row = i - 1
                .Col = TblColumn.tcWARD: bkWard = .value
                .Col = TblColumn.tcDEPT: bkDept = .value
                
                .Row = i
                .Col = TblColumn.tcPTID: .ForeColor = .BackColor
                .Col = TblColumn.tcPTNM: .ForeColor = .BackColor
                If bkReason = reason Then
                    If reason <> "(없음)" Then .Col = TblColumn.tcREASON: .ForeColor = .BackColor
                Else
                    bkReason = reason
                End If
                If bkWard = RS.Fields("wardid").value & "" Then
                    .Col = TblColumn.tcWARD: .ForeColor = .BackColor
                End If
                If bkDept = RS.Fields("deptcd").value & "" Then
                    .Col = TblColumn.tcDEPT: .ForeColor = .BackColor
                End If
                
                If bkRoomid = RS.Fields("hosilid").value & "" Then
                    .Col = TblColumn.tcROOM: .ForeColor = .BackColor
                Else
                    bkRoomid = RS.Fields("hosilid").value & ""
                End If
'                If bkReqDt = Format(RS.Fields("reqdt").value, "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value, 1, 4), "00:00") Then
'                    .Col = TblColumn.tcREQDT: .ForeColor = .BackColor
'                Else
'                    bkReqDt = Format(RS.Fields("reqdt").value, "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value, 1, 4), "00:00")
'                End If
                If bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##") Then
                    .Col = TblColumn.tcORDDT: .ForeColor = .BackColor
                Else
                    bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                End If
            End If
            
            'Irradiation 처방인 경우 글꼴을 도드라지게 표시해준다.
            .Row = .Row: .Row2 = .Row
            .Col = 1: .COL2 = .MaxCols
            .BlockMode = True
            If RS.Fields("irradfg").value & "" = "1" Then
                .FontBold = True
            Else
                If RS.Fields("statfg").value & "" = "1" Then
                    .Col = TblColumn.tcSTATFG
                    .FontBold = True
                Else
                    .FontBold = False
                End If
            End If
            .BlockMode = False
            
            '2007-06-29 추가 (퇴원일자)
            .Col = 43
            .value = GetOUTDT(RS.Fields("ptid").value & "", RS.Fields("orddt").value & "")
            
Skip:

            RS.MoveNext
        Next i
'        .ReDraw = True
        '혈액형을 일괄적으로 가지고온다.
        Set objPrgBar = Nothing
        If .DataRowCnt > 0 Then Call GetBatchABO
    End With
    

    Set QueryOrder = Nothing
End Sub

Private Function GetOUTDT(ByVal pPtId As String, ByVal pOrdDt As String) As String
    Dim RS      As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error Resume Next
    
    strSQL = " select nvl(dschdate,to_char(sysdate,'yyyymmdd')) dschdate " & _
             "   from " & T_HIS002 & _
             "  where patno = " & DBS(pPtId) & _
             "    and nvl(dschdate,to_char(sysdate,'yyyymmdd')) >= " & DBS(pOrdDt)
             
    RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        GetOUTDT = RS.Fields("dschdate").value & ""
    End If
    
    RS.Close
    Set RS = Nothing
    
End Function

Private Sub GetBatchABO()
    Dim ObjABO      As clsABO
    Dim objPrgBar   As clsProgress
    Dim QueryOrder  As clsQueryOrder
    Dim ii          As Integer
    Dim tmpptid     As String
    Dim sPtid       As String
    Dim sORDDT      As String
    Dim sLastDt     As String
    
    Set ObjABO = New clsABO
    Set objPrgBar = New clsProgress
    Set QueryOrder = New clsQueryOrder
    
    objPrgBar.Container = medMain.stsBar

    With tblPtList
        objPrgBar.Max = .DataRowCnt
        .ReDraw = False
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblColumn.tcPTID
            If tmpptid <> Trim(.value) Then
                sLastDt = ""
                sPtid = .value
                '혈액형구하기
                ObjABO.Ptid = sPtid
                If ObjABO.GetABO = False Then
                    .Col = TblColumn.tcABO:  .value = ""
                Else
                    .Col = TblColumn.tcABO:  .value = ObjABO.ABO & ObjABO.Rh
                End If
                sLastDt = QueryOrder.GetLatestTrandDt(sPtid)
                .Col = TblColumn.tcTRANSDT:  .value = sLastDt
            Else
                .Col = TblColumn.tcABO:      .value = ObjABO.ABO & ObjABO.Rh
                .Col = TblColumn.tcTRANSDT:  .value = sLastDt
            End If
            .Col = TblColumn.tcPTID: tmpptid = Trim(.value)
            If CanSelect(1, ii) Then
                .Row = ii
                .Col = TblColumn.tcSEL
                .CellType = CellTypeCheckBox
                .TypeCheckCenter = True
            Else
                .Row = ii
                .Col = TblColumn.tcSEL
                .CellType = CellTypeStaticText
                .Col = TblColumn.tcSTSNM
                If .value = STS_NM_DONE Or .value = STS_NM_END Then '"완료","종결"
                    .Col = TblColumn.tcSEL
                    .Text = "√"
                    .ForeColor = vbRed
                End If
            End If
            
            objPrgBar.value = ii: objPrgBar.Message = tmpptid & " 의 혈액형을 검색중입니다."
        Next
        .ReDraw = True
    End With
    
    Set ObjABO = Nothing
    Set QueryOrder = Nothing
    Set objPrgBar = Nothing
End Sub
Private Sub Save_LegRowCol()
'보관장소 지정이 자동이 아닐경우 보관장소를 입력받아야 하므로
'접수 버튼 클릭이전에 실행하여
'배열에 담아논다.
    Dim objXM   As New clsCrossMatching
    Dim DrRS    As New Recordset
    Dim strTmp  As String
    Dim ii      As Integer
    
    lstLeg.Clear
    lstLeg.AddItem "(없음)"
    
    DrRS.Open objXM.Get_Leg(ObjSysInfo.BuildingCd), DBConn
    With DrRS
        For ii = 1 To .RecordCount
            lstLeg.AddItem .Fields("legcd").value & ""
            .MoveNext
        Next ii
    End With
    Set DrRS = Nothing
    Set objXM = Nothing
End Sub
Private Function SaveCheckNotAuto() As Boolean
'보관장소의 입력이 되었는지 체크한다.
    Dim SavePos    As String
    Dim SaveTF     As String
    Dim DcFg       As String
    Dim strRowCol  As String
    Dim strCol     As String
    Dim ii As Integer
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If Val(.value) = 1 Then
                .Col = TblColumn.tcSTORE: SavePos = .value
                If SavePos <> "" Then
                    SaveCheckNotAuto = True
                Else
                    SaveCheckNotAuto = False
                    Exit Function
                End If
            End If
        Next
    End With
End Function
Private Function Save_Check() As Boolean
    Dim lngColCnt   As Long
    Dim ii          As Long
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If Val(.value) = 1 Then
                lngColCnt = lngColCnt + 1
                Exit For
            End If
        Next
    End With
    
    '보관장소의 입력여부를 확인한다.
    If chkSPos.value = 0 Then
        If SaveCheckNotAuto = False Then
            MsgBox "보관장소가 누락되었습니다." & vbNewLine & "확인하신후 접수하세요.", vbInformation + vbOKOnly, Me.Caption
            Exit Function
        End If
    Else
        If cboLeg.ListIndex < 1 Then
            MsgBox "보관장소 자동 부여인 경우 Rack은 반드시 선택하셔야 합니다.", vbInformation + vbOKOnly, "보관장소 Rack선택"
            Exit Function
        End If
    End If
    
    If lngColCnt = 0 Then
        '접수하고자 하는 건수를 구한다
        MsgBox "접수대상항목이 없습니다.", vbCritical + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    If Collect_Cnt = False Then Exit Function
    
    Save_Check = True

End Function

Private Sub cmdCollect_Click()
    Dim objNumbers     As clsBBSNumbers
    Dim objBg          As clsBeginTrans
    Dim RS             As Recordset
    Dim strColDt       As String
    Dim strColTm       As String
    Dim strAccDt       As String
    Dim lngAccNo       As Long
    Dim ii             As Integer
    
'    접수를 위한 변수들
    Dim strCenterCd As String
    Dim strPtid     As String
    Dim strOrdDt    As String
    Dim strPtnm As String
    Dim strSexAge As String
    Dim StrWARD As String
    Dim strDiease As String
    Dim strABO As String
    Dim strTrans As String
    Dim strDoct As String
    Dim strDept As String
    Dim strTmp      As String
    Dim strSpcYYR   As String
    Dim strFullSpc  As String
    Dim strLeg      As String
    Dim pheresis    As String
    Dim store_cnt   As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim lngSpcNoR   As Long
    Dim lngOrdseq   As Long
    Dim lngOrdNo    As Long
    Dim blnSave     As Boolean
    
    Dim SSQL        As String
    Dim strRow As String
    
    If Save_Check = False Then Exit Sub
    
    Set objBg = New clsBeginTrans
    
    Me.MousePointer = 11
    strCenterCd = ObjSysInfo.BuildingCd         '센터코드
    strColDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strColTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    
    Set objNumbers = New clsBBSNumbers
    With objNumbers
        strAccDt = .Get_AccdtFormat
        lngAccNo = Val(.Get_AccDT_Seq(strAccDt))
    End With
    
On Error GoTo Save_Spc_Error

    DBConn.BeginTrans
    
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            
            If Val(.value) = 1 Then
                strRow = strRow & ii & COL_DIV
                
                .Col = TblColumn.tcDCFG
                
                .Col = TblColumn.tcPTID:     strPtid = .value
                .Col = TblColumn.tcSPCNO:    strSpcYYR = Mid(.value, 1, 2)
                                             lngSpcNoR = Val(Mid(.value, 4))
                                             strFullSpc = strSpcYYR & CStr(lngSpcNoR)
                .Col = TblColumn.tcPHERESIS: pheresis = IIf(.value = "1", "1", "0")
                
                .Col = TblColumn.tcORDDT:    strOrdDt = Mid(.value, 1, 4) & Mid(.value, 6, 2) & Mid(.value, 9, 2)
                .Col = TblColumn.tcORDNO:    lngOrdNo = Val(.value)
                .Col = TblColumn.tcORDSEQ:   lngOrdseq = Val(.value)
                
                SSQL = objBg.Set_UpdateL101(strPtid, strOrdDt, CStr(lngOrdNo))
                DBConn.Execute SSQL
                
                SSQL = objBg.Set_UpdateL102(strPtid, strOrdDt, CStr(lngOrdNo), CStr(lngOrdseq), strAccDt, CStr(lngAccNo))
                DBConn.Execute SSQL
                
                
                SSQL = objBg.Set_BBS202_Insert(strAccDt, lngAccNo, strPtid, strOrdDt, CStr(lngOrdNo), CStr(lngOrdseq), ObjMyUser.EmpId, pheresis)
                DBConn.Execute SSQL
                
                'OCS 관련 Acting Check
'                If OCSActingCheck(strPtid, strOrdDt, CStr(lngOrdNo), CStr(lngOrdseq)) = False Then GoTo Save_Spc_Error
                
               '검체번호는 있는 데 처방 접수가 없는 경우
               '성분헌혈이 아닌경우는 검체 해당자료는 저장하지 않는다.
               '이미 검체가 지정되어있는 경우는 검체보관장소를 update 해주지 않는다.
               
                .Col = TblColumn.tcACCNO
                If .value = "" And strFullSpc <> "" Then
                    If strTmp <> strPtid Then
                        If chkSPos.value = 0 Then    '보관장소임의 지정
                            Set RS = objBg.SavePositionRs(strCenterCd, strSpcYYR, CStr(lngSpcNoR))
                            If Not RS.EOF Then
                                strLeg = RS.Fields("legcd").value & ""
                                lngRow = Val(RS.Fields("rowno").value & "")
                                lngCol = Val(RS.Fields("colno").value & "")
                            Else
                                .Col = TblColumn.tcSTORE
                                strLeg = Mid(.value, 1, 1)
                                lngRow = Val(medGetP(medGetP(.value, 1, ","), 2, "("))
                                lngCol = Val(medGetP(medGetP(.value, 2, ","), 1, ")"))
                                SSQL = objBg.Set_UpdateB206(strCenterCd, strLeg, lngRow, lngCol, strSpcYYR, CStr(lngSpcNoR))
                                DBConn.Execute SSQL
                            End If
                            SSQL = objBg.Set_UpdateB201(strFullSpc, ObjMyUser.EmpId, strLeg, lngRow, lngCol)
                            DBConn.Execute SSQL
                            
                        Else                        '보관장소 자동지정
                            Set RS = objBg.SavePositionRs(strCenterCd, strSpcYYR, CStr(lngSpcNoR))
                            If Not RS.EOF Then
                                strLeg = RS.Fields("legcd").value & ""
                                lngRow = Val(RS.Fields("rowno").value & "")
                                lngCol = Val(RS.Fields("colno").value & "")
                            Else
                                store_cnt = store_cnt + 1
                                strLeg = aryLeg(store_cnt - 1)
                                lngRow = aryRow(store_cnt - 1)
                                lngCol = aryCol(store_cnt - 1)
    
                                SSQL = objBg.Set_UpdateB201(strFullSpc, ObjMyUser.EmpId, strLeg, lngRow, lngCol)
                                DBConn.Execute SSQL
                            End If
                            SSQL = objBg.Set_UpdateB206(strCenterCd, strLeg, lngRow, lngCol, strSpcYYR, CStr(lngSpcNoR))
                            DBConn.Execute SSQL
                            
                        End If
                        Set RS = Nothing
                    End If
                    
                    strTmp = strPtid
                End If
                
                '조회시 속도개선을 위해서 접수시 필요데이터를 생성한다.
                Dim objCollect  As clsBBSCollection
                Dim SQLTmp      As String
                
                Set objCollect = New clsBBSCollection
                SQLTmp = objCollect.Set_AccUnitSQL_203(strPtid, strAccDt, CStr(lngAccNo))
                SSQL = medGetP(SQLTmp, 1, COL_DIV)
                DBConn.Execute SSQL
                SSQL = medGetP(SQLTmp, 2, COL_DIV)
                DBConn.Execute SSQL
                Set objCollect = Nothing
                lngAccNo = lngAccNo + 1
                blnSave = True
            End If
            
        Next ii
    End With
    
    If blnSave = True Then
        SSQL = objNumbers.Set_NumbersCom099(BN_ACC_NO, strAccDt, lngAccNo - 1)
        DBConn.Execute SSQL
    End If
    
    DBConn.CommitTrans
    Call Query
    
    Me.MousePointer = 0
    MsgBox "접수되었습니다.", vbInformation, "접수"
    
    If blnSave And (chkAutoPrint.value = 1) Then '정상적으로 처리된 경우에 출고전표를 출력해준다.
        DoEvents
        Call PrintTransList(strRow)
    End If
    
    Set objBg = Nothing
    Set objNumbers = Nothing
    Exit Sub
    
Save_Spc_Error:
    
    DBConn.RollbackTrans
    Me.MousePointer = 0
    MsgBox "정상적으로 처리되지 않았습니다.", vbInformation, "접수오류"
    Set objBg = Nothing
    Set objNumbers = Nothing
End Sub


'Private Function OCSActingCheck(ByVal strPtid As String, ByVal strOrdDt As String, _
'                                ByVal strOrdNo As String, ByVal strOrdSeq As String) As Boolean
'    Dim Rs          As Recordset
'    Dim SqlStmt     As String
'    Dim strOcsOrdNo As String
'    Dim strBussdiv  As String
'
'On Error GoTo Errors
'
'    '접수시 OCS 관련 Table 에 Acting_Check를 해준다.
'
'    SqlStmt = " SELECT a.ocsordno,b.bussdiv " & _
'              " FROM " & T_LAB101 & " b," & T_LAB102 & " a" & _
'              " WHERE " & DBW("a.ptid =", strPtid) & _
'              " AND " & DBW("a.orddt=", strOrdDt) & _
'              " AND " & DBW("a.ordno=", strOrdNo) & _
'              " AND " & DBW("a.ordseq=", strOrdSeq) & _
'              " AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno"
'    Set Rs = New Recordset
'    Rs.Open SqlStmt, DBConn
'
'    If Not Rs.EOF Then
'        strOcsOrdNo = Val(Trim(Rs.Fields("ocsordno").value & ""))
'        strBussdiv = Trim(Rs.Fields("bussdiv").value & "")
'        '병동은 ipd_order_dmc,ipd_order_update_dmc 업데이트
'        '외래는 opd_order_dmc 업데이트
'        If strBussdiv = enBussDiv.BussDiv_InPatient Then
'            SqlStmt = " UPDATE med_ocs.ipd_order_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'            DBConn.Execute SqlStmt
'            SqlStmt = " UPDATE med_ocs.ipd_order_update_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'            DBConn.Execute SqlStmt
'        Else
'            SqlStmt = " UPDATE med_ocs.opd_order_dmc SET acting_check='1' where order_key=" & strOcsOrdNo
'            DBConn.Execute SqlStmt
'        End If
'    End If
'
'    Set Rs = Nothing
'    OCSActingCheck = True
'    Exit Function
'
'Errors:
'    Set Rs = Nothing
'    OCSActingCheck = False
'End Function


Private Function Collect_Cnt() As Boolean
    Dim objSpec     As clsSpecManagement
    Dim strTmp      As String
    Dim strCollect  As String        '접수여부...
    Dim strGather   As String         '채혈여부...
    Dim store_cnt   As Integer
    Dim lngColCnt   As Integer
    Dim ii          As Integer
    
    Set objSpec = New clsSpecManagement

    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcSEL
            If Val(.value) = 1 Then
                lngColCnt = lngColCnt + 1
                .Col = TblColumn.tcPTID
                If .value <> strTmp Then
                    store_cnt = store_cnt + 1
                End If
                strTmp = .value
            End If
        Next
    End With
    If chkSPos.value = 1 Then
        If lngColCnt <> 0 Then
            With objSpec
                If .Save_Spc_Search(store_cnt, ObjSysInfo.BuildingCd, cboLeg.Text) Then
                    ReDim aryLeg(store_cnt)
                    ReDim aryRow(store_cnt)
                    ReDim aryCol(store_cnt)
                    For ii = 1 To store_cnt
                        aryLeg(ii - 1) = .Leg(ii)
                        aryRow(ii - 1) = .Row(ii)
                        aryCol(ii - 1) = .Col(ii)
                    Next
                    Collect_Cnt = True
                Else
                    Collect_Cnt = False
                End If
            End With
        End If
    End If
    Set objSpec = Nothing

End Function

Private Sub PrintOrderList()
'출력하자.....크리스탈
    Dim strPtid As String, strPtnm As String, strABO As String, strOrdDt As String, STRUNIT As String, strReqDt As String
    Dim strStat As String, STRDCFG As String, STRSTS As String, strSpcNo As String, strSave As String, STRBUILD As String
    Dim StrWARD As String, strDept As String, StrACC As String, strOrdNm As String, STRREAN As String, STRDISEA As String
    Dim strTmp  As String
    
    Dim strRfile   As String
    Dim strRptPath As String
    Dim intFNum    As Integer
    Dim ii         As Integer
    
    Dim sDupChk    As String
    Dim sICSStr    As String

    If tblPtList.MaxRows = 0 Then Exit Sub
    Me.MousePointer = 11
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            STRDISEA = ""
            
            .Col = TblColumn.tcPTID:    strPtid = .value
            
            If sDupChk <> strPtid Then
                sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
                .Col = TblColumn.tcPTNM:    strPtnm = Trim(.value) & sICSStr
            Else
                .Col = TblColumn.tcPTNM:    strPtnm = Trim(.value) & sICSStr
            End If
            
            sICSStr = ""
            sDupChk = strPtid
            .Col = TblColumn.tcABO:     strABO = Trim(.value)
            .Col = TblColumn.tcORDNM:   strOrdNm = Trim(.value)
            .Col = TblColumn.tcORDDT:   strOrdDt = Trim(.value)
            .Col = TblColumn.tcUNITQTY: STRUNIT = Trim(.value)
            .Col = TblColumn.tcREASON:  STRREAN = Trim(.value)
            .Col = TblColumn.tcREQDT:    strReqDt = Trim(.value)
            .Col = TblColumn.tcSTATnm:   strStat = Trim(.value)
            .Col = TblColumn.tcDCNM:    STRDCFG = Trim(.value)
            
            .Col = TblColumn.tcSTSNM:   STRSTS = Trim(.value)
            
            .Col = TblColumn.tcDISEASE: STRDISEA = Trim(.value)
            
            If STRDISEA <> "" Then
                .Col = TblColumn.tcDISEASE2
                If .value <> "" Then
                    STRDISEA = STRDISEA & "," & Trim(.value)
                Else
                    STRDISEA = STRDISEA
                End If
                .Col = TblColumn.tcDISEASE3
                If .value <> "" Then
                    STRDISEA = STRDISEA & "," & Trim(.value)
                Else
                    STRDISEA = STRDISEA
                End If
                .Col = TblColumn.tcDISEASE4
                If .value <> "" Then
                    STRDISEA = STRDISEA & "," & Trim(.value)
                Else
                    STRDISEA = STRDISEA
                End If
            End If
                        
            .Col = TblColumn.tcSPCNO:    strSpcNo = Trim(.value)
            .Col = TblColumn.tcSTORE:    strSave = Trim(.value)
            .Col = TblColumn.tcACCNO:    StrACC = Trim(.value)
            
            .Col = TblColumn.tcCENTERNM: STRBUILD = Trim(.value)
            .Col = TblColumn.tcWARD:     StrWARD = Trim(.value)
            .Col = TblColumn.tcDEPT:     strDept = Trim(.value)
            strTmp = strTmp & strPtid & vbTab & strPtnm & vbTab & strABO & vbTab & strOrdDt & vbTab & STRUNIT & vbTab & strReqDt & vbTab & _
                     strStat & vbTab & STRDCFG & vbTab & STRSTS & vbTab & strSpcNo & vbTab & strSave & vbTab & STRBUILD & vbTab & _
                     StrWARD & vbTab & strDept & vbTab & StrACC & vbTab & strOrdNm & vbTab & STRREAN & vbTab & STRDISEA & vbCr
        Next ii
    End With
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 1)

    strRfile = InstallDir & "BBS\Rpt" & "\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\Rpt" & "\frmBBS102.rpt"
    
    Crystal_Print CReport, strTmp, strRfile, strRptPath
    Me.MousePointer = 0
End Sub

'Private Sub PrintTransReport()
''출력하자.....크리스탈
'    Dim strPtID As String, strPtNm As String, strABO As String, strOrdDt As String, STRUNIT As String, strReqDt As String
'    Dim strStat As String, STRDCFG As String, STRSTS As String, strSpcNo As String, strSave As String, STRBUILD As String
'    Dim StrWARD As String, STRDEPT As String, StrACC As String, strOrdNm As String, STRREAN As String, STRDISEA As String
'
'    Dim ii         As Integer
'
'    Dim sDupChk    As String
'    Dim sICSStr    As String
'
'
'    Dim objPrint   As clsBBSPrint
'
'    Dim strHeader1 As String
'    Dim strHeader2 As String
'    Dim strHeader3 As String
'    Dim strBody    As String
'
'    If tblPtList.MaxRows = 0 Then Exit Sub
'    Me.MousePointer = 11
'    Set objPrint = New clsBBSPrint
'
'    strHeader1 = "수혈처방출력"
'    strHeader2 = "♣ 출력자 : " & ObjSysInfo.EmpNm & Space(5) & "♣ 출력일 : " & Format(Now, "YYYY-MM-DD HH:MM") & COL_DIV & "5" & COL_DIV & "1"
'    strHeader3 = "번호" & COL_DIV & "5" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "환자ID" & COL_DIV & "15" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "환자명" & COL_DIV & "35" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "혈액형" & COL_DIV & "75" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "처방일자" & COL_DIV & "90" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "수량" & COL_DIV & "120" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "예정일시" & COL_DIV & "130" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "응급" & COL_DIV & "170" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "D/C" & COL_DIV & "180" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "상태" & COL_DIV & "190" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "검체번호" & COL_DIV & "205" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "보관장소" & COL_DIV & "230" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "검사장소" & COL_DIV & "250" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "Location" & COL_DIV & "270" & COL_DIV & "1"
'    strHeader3 = strHeader3 & vbTab & "접수번호" & COL_DIV & "15" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "처방명" & COL_DIV & "35" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "사유" & COL_DIV & "120" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "진단명" & COL_DIV & "205" & COL_DIV & "1"
'
'    With tblPtList
'        For ii = 1 To .MaxRows
'            .Row = ii
'            STRDISEA = ""
'            .Col = TblColumn.tcPTID:    strPtID = .value
'            If sDupChk <> strPtID Then
'                sICSStr = ICSPatientString(strPtID, enICSNum.BBS_ALL)
'                .Col = TblColumn.tcPTNM:    strPtNm = Trim(.value) & sICSStr
'            Else
'                .Col = TblColumn.tcPTNM:    strPtNm = Trim(.value) & sICSStr
'            End If
'            sICSStr = ""
'            sDupChk = strPtID
'            .Col = TblColumn.TcABO:     strABO = Trim(.value)
'            .Col = TblColumn.tcORDNM:   strOrdNm = Trim(.value)
'            .Col = TblColumn.tcORDDT:   strOrdDt = Trim(.value)
'            .Col = TblColumn.tcUNITQTY: STRUNIT = Trim(.value)
'            .Col = TblColumn.tcREASON:  STRREAN = Trim(.value)
'            .Col = TblColumn.tcREQDT:    strReqDt = Trim(.value)
'            .Col = TblColumn.tcSTATnm:   strStat = Trim(.value)
'            .Col = TblColumn.tcDCNM:    STRDCFG = Trim(.value)
'
'            .Col = TblColumn.tcSTSNM:   STRSTS = Trim(.value)
'
'            .Col = TblColumn.tcDISEASE: STRDISEA = Trim(.value)
'
'            If STRDISEA <> "" Then
'                .Col = TblColumn.tcDISEASE2
'                If .value <> "" Then
'                    STRDISEA = STRDISEA & vbTab & Trim(.value)
'                Else
'                    STRDISEA = STRDISEA
'                End If
'                .Col = TblColumn.tcDISEASE3
'                If .value <> "" Then
'                    STRDISEA = STRDISEA & vbTab & Trim(.value)
'                Else
'                    STRDISEA = STRDISEA
'                End If
'                .Col = TblColumn.tcDISEASE4
'                If .value <> "" Then
'                    STRDISEA = STRDISEA & vbTab & Trim(.value)
'                Else
'                    STRDISEA = STRDISEA
'                End If
'            End If
'
'
'            .Col = TblColumn.tcSPCNO:    strSpcNo = Trim(.value)
'            .Col = TblColumn.tcSTORE:    strSave = Trim(.value)
'            .Col = TblColumn.tcACCNO:    StrACC = Trim(.value)
'
'            .Col = TblColumn.tcCENTERNM: STRBUILD = Trim(.value)
'            .Col = TblColumn.tcWARD:     StrWARD = Trim(.value)
'            .Col = TblColumn.tcDEPT:     STRDEPT = Trim(.value)
'            If StrWARD <> "" Then
'                StrWARD = StrWARD & "-" & STRDEPT
'            Else
'                StrWARD = STRDEPT
'            End If
'
'            strBody = strBody & ii & COL_DIV & "5" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strPtID & COL_DIV & "15" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strPtNm & COL_DIV & "35" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strABO & COL_DIV & "75" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strOrdDt & COL_DIV & "90" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRUNIT & COL_DIV & "120" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strReqDt & COL_DIV & "130" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strStat & COL_DIV & "170" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRDCFG & COL_DIV & "180" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRSTS & COL_DIV & "190" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strSpcNo & COL_DIV & "205" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strSave & COL_DIV & "230" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRBUILD & COL_DIV & "250" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & StrWARD & COL_DIV & "270" & COL_DIV & "1" & COL_DIV & "0"
'            strBody = strBody & vbTab & StrACC & COL_DIV & "15" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & strOrdNm & COL_DIV & "35" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRREAN & COL_DIV & "120" & COL_DIV & "0" & COL_DIV & "0"
'            strBody = strBody & vbTab & STRDISEA & COL_DIV & "205" & COL_DIV & "1" & COL_DIV & "1" & vbTab
'
'        Next ii
'    End With
'    strBody = Mid(strBody, 1, Len(strBody) - 1)
'
'    With objPrint
'        .Header1 = strHeader1
'        .Header2 = strHeader2
'        .Header3 = strHeader3
'        .Body = strBody
'        Call .CallPrint("가로")
'    End With
'
'    Set objPrint = Nothing
'
'    Me.MousePointer = 0
'End Sub

'2001-11-30추가
Private Sub PrintDeliveryList(Optional ByVal blnReprint As Boolean = False)

'출력하자.....크리스탈
    Dim strPtid As String, strPtnm As String, strABO As String, STRUNIT As String, strReqDt As String
    Dim StrWARD As String, strDept As String, strOrdNm As String, STRDISEA As String
    Dim strTmp  As String, strDoct As String, strTransDt As String
    
    Dim strRfile   As String
    Dim strRptPath As String
    Dim intFNum    As Integer
    Dim ii         As Integer
    Dim jj         As Integer
    Dim lngCnt     As Long
    
    Dim sDupChk     As String
    Dim sICSStr     As String
    

    If tblPtList.MaxRows = 0 Then Exit Sub
    Me.MousePointer = 11
    lngCnt = 0
    STRDISEA = ""
    With tblPtList
        For ii = 1 To .MaxRows
            .Row = ii
            If ii = 1 Then
                .Col = TblColumn.tcPTID:    strPtid = .value
                
                If sDupChk <> strPtid Then
                    sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
                    .Col = TblColumn.tcPTNM:    strPtnm = Trim(.value) & sICSStr
                Else
                    .Col = TblColumn.tcPTNM:    strPtnm = Trim(.value) & sICSStr
                End If
                sICSStr = ""
                sDupChk = strPtid
                
                .Col = TblColumn.tcABO:     strABO = Trim(.value)
                .Col = TblColumn.tcREQDT:   strReqDt = Trim(.value)
                .Col = TblColumn.tcDISEASE: STRDISEA = Trim(.value)
                .Col = TblColumn.tcWARD:    StrWARD = Trim(.value)
                .Col = TblColumn.tcDEPT:    strDept = Trim(.value)
                .Col = TblColumn.tcDOCT:    strDoct = Trim(.value)
                .Col = TblColumn.tcTRANSDT: strTransDt = Trim(.value)
                
                strDoct = GetDoctNm(strDoct)
                strDept = GetDeptNm(strDept)
                
                If STRDISEA <> "" Then
                    .Col = TblColumn.tcDISEASE2
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                    .Col = TblColumn.tcDISEASE3
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                    .Col = TblColumn.tcDISEASE4
                    If .value <> "" Then
                        STRDISEA = STRDISEA & "," & Trim(.value)
                    Else
                        STRDISEA = STRDISEA
                    End If
                End If
            End If
            .Col = TblColumn.tcORDNM:   strOrdNm = Trim(.value)
            .Col = TblColumn.tcUNITQTY: STRUNIT = Trim(.value)
            
'            If Not blnReprint Then
                For jj = 1 To Val(STRUNIT)
                    strTmp = strTmp & "" & vbTab & strOrdNm & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
                             "" & vbTab & "" & vbTab & "" & vbTab & "" & vbCr
                    lngCnt = lngCnt + 1
                Next
'            End If
        Next ii
    End With

'    If blnReprint Then
'        strTmp = String(23, vbCr)
'    Else
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1) & String(24 - lngCnt, vbCr)
'    End If

    strRfile = InstallDir & "BBS\RPT\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\RPT\frmBBS102_1.rpt"

    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    With CReport
        .ReportFileName = strRptPath
        .ParameterFields(0) = "ptid;" & strPtid & ";TRUE"
        .ParameterFields(1) = "ptnm;" & strPtnm & ";TRUE"
        .ParameterFields(2) = "ward;" & StrWARD & ";TRUE"
        .ParameterFields(3) = "abo;" & strABO & ";TRUE"
        .ParameterFields(4) = "sicknm;" & STRDISEA & ";TRUE"
        .ParameterFields(5) = "doct;" & strDoct & ";TRUE"
        .ParameterFields(6) = "dept;" & strDept & ";TRUE"
        .ParameterFields(7) = "hostnm;" & HOSPITAL_NAME & ";TRUE"
        .ParameterFields(8) = "transdt;" & Format(strTransDt, CS_DateLongMask) & ";TRUE"
        .ParameterFields(9) = "sexage;" & lblSex.Caption & " / " & lblAge.Caption & ";TRUE"
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
        .Reset
    End With
    Me.MousePointer = 0
End Sub

Private Sub PrintIntialize()
    PrtLeft = 5
    LineSpace = 6
    lngCurYPos = 10
    
    
    Printer.Font = "굴림체"
    Printer.FontSize = 9
    Printer.Orientation = vbPRORPortrait '/* 좁게
    Printer.ScaleMode = vbMillimeters
    

    Twidth = Printer.ScaleWidth

    LastLineYpos = Printer.ScaleHeight             '마지막라인Y위치

End Sub

Private Sub PrintTrans(ByVal vRow As Long)
'프린트 오브젝트를 사용할 경우에만 씀
    Dim lngX1 As Long
    Dim lngX2 As Long
    Dim lngX3 As Long
    
    Dim i As Long
    Dim strPtnm As String
    Dim StrWARD As String
    Dim strPtid As String
    Dim strDiease As String
    Dim strABO As String
    Dim strTrans As String
    Dim strDoct As String
    Dim strDept As String
    Dim strSexAge As String
    
    
'처방이 다른 경우에 출력
'접수번호, 검체번호까지 출력해줘야..


    With tblPtList
        For i = 1 To .DataRowCnt
            .Col = TblColumn.tcPTNM: strPtnm = .value
            .Col = TblColumn.tcWARD: StrWARD = .value
            .Col = TblColumn.tcPTID: strPtid = .value
'            .Col = "" 'Sex
            .Col = TblColumn.tcDISEASE: strDiease = .value
            .Col = TblColumn.tcABO: strABO = .value
            .Col = TblColumn.tcTRANSDT: strTrans = .value
'            .Col = "" 'IM
            .Col = TblColumn.tcDOCT: strDoct = .value
            .Col = TblColumn.tcDEPT: strDept = .value
            
            strDoct = GetDoctNm(strDoct)
            strDept = GetDeptNm(strDept)
            
            If strDiease <> "" Then
                .Col = TblColumn.tcDISEASE2
                If .value <> "" Then
                    strDiease = strDiease & "," & Trim(.value)
                Else
                    strDiease = strDiease
                End If
                .Col = TblColumn.tcDISEASE3
                If .value <> "" Then
                    strDiease = strDiease & "," & Trim(.value)
                Else
                    strDiease = strDiease
                End If
                .Col = TblColumn.tcDISEASE4
                If .value <> "" Then
                    strDiease = strDiease & "," & Trim(.value)
                Else
                    strDiease = strDiease
                End If
            End If
            
            Call PrintIntialize
        Next
    End With
    
    
    lngX1 = 10
    lngX2 = lngX1 + Printer.TextWidth("성    명 : ")
    lngX3 = lngX1 + 70
    
    Printer.FontSize = 16: Printer.FontBold = True
    Call Print_Setting("수혈 요청 및 출고 전표", PrtLeft, lngCurYPos, Twidth, "C", "C", False)
    Printer.FontSize = 13: Printer.FontBold = False
    
    lngCurYPos = lngCurYPos + 20
    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos + 70), , B 'Box 그리기
    
    '성명, 병동, 혈액형 같은 Top에 그리기
    lngCurYPos = lngCurYPos + LineSpace
    Call Print_Setting("성    명 : " & strPtnm, lngX1, LineSpace, , , "C", False)
    Call Print_Setting("병    동 : " & StrWARD, lngX3, LineSpace, , , "C", False)
    Call Print_Setting("   혈액형 ", 130, LineSpace, , "L", "C", False)
    
    '등록번호, 성별/나이, 혈액형값 같은 Top에 그리기
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("등록번호 : " & strPtid, lngX1, LineSpace, , , "C", False)
    Call Print_Setting("성별/나이 : " & strSexAge, lngX3, LineSpace, , , "C", False)
    Printer.FontBold = True: Printer.FontSize = 40
    Call Print_Setting(strABO, 135, LineSpace, , , "C", False)
    Printer.FontBold = False: Printer.FontSize = 13
    
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("진 단 명 : " & strDiease, lngX1, 10, , , "C", False)
    
    lngCurYPos = lngCurYPos + 10
'    Call Print_Setting("수 혈 력 :     □ 무      □ 유 " & pTrans, lngX1, 10, , , "C", False)
'    lngCurYPos = lngCurYPos + 10
'    Call Print_Setting("임 신 력 :     □ 무      □ 유  (     주)" & pIM, lngX1, 10, , , "C", False)
'    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("담당의사 : " & strDoct, lngX1, 10, , , "C", False)
    Call Print_Setting("진 료 과 : " & strDept, lngX3, 10, , , "C", False)
    
'    lngCurYPos = lngCurYPos + 10
'
'    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos)
'    Dim ii As Integer
'
'    lngCurYPos = lngCurYPos + 2
'
'    For ii = 1 To 12
'        Printer.Line (PrtLeft, lngCurYPos + 8 * ii)-(Twidth - PrtLeft, lngCurYPos + 8 * ii)
'    Next
'
''혈액불출
'    Printer.Line (PrtLeft, lngCurYPos - 2)-(PrtLeft, lngCurYPos + 8 * 12)
'
'    '혈액번호
'    Printer.Line (lngX2, lngCurYPos + 8)-(lngX2, lngCurYPos + 8 * 12)
'    '혈액종류
'    Printer.Line (lngX2 + 30, lngCurYPos + 8)-(lngX2 + 30, lngCurYPos + 8 * 12)
'    '혈액형
'    Printer.Line (lngX2 + 45, lngCurYPos + 8)-(lngX2 + 45, lngCurYPos + 8 * 12)
'    '채혈자
'    Printer.Line (lngX2 + 60, lngCurYPos + 8)-(lngX2 + 60, lngCurYPos + 8 * 12)
'
'    '수혈시작시간
'
'    Printer.Line (lngX2 + 75, lngCurYPos + 8)-(lngX2 + 75, lngCurYPos + 8 * 12)
'
'    Printer.Line (lngX2 + 90, lngCurYPos - 2)-(lngX2 + 90, lngCurYPos + 8 * 12)
'
'
'
'
'    Printer.Line (lngX2 + 105, lngCurYPos + 8)-(lngX2 + 105, lngCurYPos + 8 * 12)
'
'
'
'    '수혈끝시간
'    Printer.Line (lngX2 + 120, lngCurYPos + 8)-(lngX2 + 120, lngCurYPos + 8 * 12)
'    'Dr
'    Printer.Line (lngX2 + 130, lngCurYPos + 8)-(lngX2 + 130, lngCurYPos + 8 * 12)
'    'Nr
'    Printer.Line (lngX2 + 140, lngCurYPos + 8)-(lngX2 + 140, lngCurYPos + 8 * 12)
'    '수혈부작용
'    'Printer.Line (lngX2 + 165, lngCurYPos + 8)-(lngX2 + 142, lngCurYPos + 8 * 12)
'
'    '마지막
'    Printer.Line (Twidth - PrtLeft, lngCurYPos - 2)-(Twidth - PrtLeft, lngCurYPos + 8 * 12)
'
'    Printer.FontSize = 10
'
'    Call Print_Setting("혈액불출기록", PrtLeft, 8, , , "C", False)
'    Call Print_Setting("수혈기록", lngX2 + 90, 8, , , "C", False)
'
'    lngCurYPos = lngCurYPos + LineSpace
'
'    Call Print_Setting("혈액불출시간", PrtLeft, 12, lngX2 - PrtLeft, "C", "C", False)
'    Call Print_Setting("혈액번호", lngX2, 12, 30, "C", "C", False)
'    Call Print_Setting("혈액종류", lngX2 + 30, 12, 15, "C", "C", False)
'    Call Print_Setting("혈액형", lngX2 + 45, 12, 15, "C", "C", False)
'    Call Print_Setting("채혈일", lngX2 + 60, 12, 15, "L", "C", False)
'    Call Print_Setting("출고자", lngX2 + 75, 12, 27, "L", "C", False)
'    Call Print_Setting("수혈시간", lngX2 + 90, 12, 20, "L", "C", False)
'    Call Print_Setting("수혈끝", lngX2 + 105, 12, 20, "L", "C", False)
'    Call Print_Setting("Dr.", lngX2 + 120, 12, 10, "C", "C", False)
'
'    Call Print_Setting("Nr.", lngX2 + 130, 12, 10, "C", "C", False)
'    Call Print_Setting("수혈부작용", lngX2 + 140, 12, 20, "C", "C", False)
'
'    lngCurYPos = lngCurYPos + 8 * 12
'    Printer.FontBold = True
'    Call Print_Setting("Memo (Special v/s 및 환자상태기록)", PrtLeft, LineSpace, , , "C")
'
'    Printer.Line (PrtLeft, lngCurYPos)-(Twidth - PrtLeft, lngCurYPos + 50), , B
'
'
'    Printer.Line (PrtLeft, lngCurYPos + 55)-(Twidth - PrtLeft, lngCurYPos + 55)
'
'    lngCurYPos = lngCurYPos + 60
'
'    Call Print_Setting(HOSPITAL_NAME, PrtLeft, LineSpace, Twidth, "C", "C", False)
'    Printer.FontBold = False
    
    Printer.EndDoc
End Sub

Private Sub PrintTransList(ByVal vRow As String)
'크로스 매칭용 전표 작성..
'환자 정보및 처방 정보는 파라미터로 넘기고
'관련검사는 필드로 넘겨준다.
'기본값은 접수가 완료된 후 자동 발행 (선택한 로우에 대한 재발행 기능)
    
    '파라미터용 변수 선언
    Dim strPtnm As String
    Dim strWardNm As String
    Dim strABO As String
    Dim strStat As String
    Dim strPtid As String
    Dim strSexAge As String
    Dim strOrdDoct As String
    Dim strDept As String
    Dim strDisease As String
    Dim strOrdDt As String
    Dim strOrdNo As String
    Dim strOrdNm As String
    Dim strColdttm As String
    Dim strColNm As String
    Dim strUnitQty As String
    Dim strSpcNo As String
    Dim strStore As String
    Dim strAccNo As String
    Dim strAccdttm As String
    Dim strAccNm As String
    Dim strRelTest As String
    Dim aryRelTest() As String
    Dim strTemp As String
    Dim aryRow() As String
    Dim strabScreen As String
    Dim strdCoombs As String
    
    Dim strRfile   As String
    Dim strRptPath As String
    Dim lngFileNo As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
'    Dim objDisease As clsDisease
    
    If tblPtList.MaxRows = 0 Then Exit Sub
        
    aryRow = Split(vRow, COL_DIV)
    
    Me.MousePointer = vbHourglass
    
    With tblPtList
        For i = LBound(aryRow) To UBound(aryRow)
            If aryRow(i) <> "" Then
                .Row = Val(aryRow(i))
                
                .Col = TblColumn.tcPTNM: strPtnm = .value
                .Col = TblColumn.tcWARD:
                If .value <> "" Then
                    strWardNm = GetWardNm(.value)
                Else
                    strWardNm = "외래"
                End If
                .Col = TblColumn.tcABO: strABO = .value
                .Col = TblColumn.tcSTATFG: strStat = IIf(.value = "1", "응급", "")
                .Col = TblColumn.tcPTID: strPtid = .value
                .Col = 0 'SexAge
                .Col = TblColumn.tcDOCT: strOrdDoct = GetDoctNm(.value)
                .Col = TblColumn.tcDEPT: strDept = GetDeptNm(.value)
                If strDept = "응급의학과" Then
                    strWardNm = "EM"
                End If
                .Col = TblColumn.tcDISEASE: strDisease = Trim(.value)
                If strDisease <> "" Then
                    .Col = TblColumn.tcDISEASE2
                    If .value <> "" Then
                        strDisease = strDisease & "," & Trim(.value)
                    Else
                        strDisease = strDisease
                    End If
                    .Col = TblColumn.tcDISEASE3
                    If .value <> "" Then
                        strDisease = strDisease & "," & Trim(.value)
                    Else
                        strDisease = strDisease
                    End If
                    .Col = TblColumn.tcDISEASE4
                    If .value <> "" Then
                        strDisease = strDisease & "," & Trim(.value)
                    Else
                        strDisease = strDisease
                    End If
                End If
                .Col = TblColumn.tcORDDT: strOrdDt = .value
                .Col = TblColumn.tcORDNO: strOrdNo = .value
                .Col = TblColumn.tcORDNM: strOrdNm = .value
                .Col = 0 'Coldttm
                .Col = 0 'Colnm
                .Col = TblColumn.tcUNITQTY: strUnitQty = .value
                .Col = TblColumn.tcSPCNO: strSpcNo = .value
                .Col = TblColumn.tcSTORE: strStore = .value
                .Col = TblColumn.tcACCNO: strAccNo = .value
                .Col = TblColumn.tcACCDTTM: strAccdttm = .value
                .Col = 0 'Accnm
                                
'                '상병불러오기 최초 상병만 불러온다.
'                Set objDisease = Nothing
'                Set objDisease = New clsDisease
'
'                objDisease.Clear
'                objDisease.PtId = strPtid
'                objDisease.orddt = Format(strOrdDt, "yyyyMMdd")
'                objDisease.ordno = strOrdNo
'
'                If objDisease.GetDisease Then
'                    strDisease = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
'                End If
'
'                Set objDisease = Nothing
                
                '환자마스터에서 SexAge를 구한다.
                strSexAge = GetSexAge(strPtid)
                
                '채혈, 접수정보를 읽는다.
                Call GetColAccInfo(strSpcNo, strColdttm, strColNm, strAccdttm, strAccNm)
                
                '관련검사가 있는경우 조회
                strRelTest = GetRelTest(strPtid)
                If strRelTest <> "" Then aryRelTest = Split(strRelTest, vbTab)
                
                strRfile = InstallDir & "BBS\RPT\CrystalReport.txt"
                strRptPath = InstallDir & "BBS\RPT\frmBBS102_2.rpt"
            
                lngFileNo = FreeFile
                Open strRfile For Output As #lngFileNo
                Print #lngFileNo, strRelTest
                Close #lngFileNo
                With CReport
                    .ReportFileName = strRptPath
                    
                    .ParameterFields(0) = "ptnm;" & strPtnm & ";TRUE"
                    .ParameterFields(1) = "wardnm;" & strWardNm & ";TRUE"
                    .ParameterFields(2) = "abo;" & strABO & ";TRUE"
                    .ParameterFields(20) = "stat;" & strStat & ";TRUE"
                    .ParameterFields(3) = "ptid;" & strPtid & ";TRUE"
                    .ParameterFields(4) = "sexage;" & strSexAge & ";TRUE"
                    .ParameterFields(5) = "orddoct;" & strOrdDoct & ";TRUE"
                    .ParameterFields(6) = "dept;" & strDept & ";TRUE"
                    .ParameterFields(7) = "disease;" & strDisease & ";TRUE"
                    .ParameterFields(8) = "orddt;" & strOrdDt & ";TRUE"
                    .ParameterFields(9) = "ordnm;" & strOrdNm & ";TRUE"
                    .ParameterFields(10) = "coldttm;" & strColdttm & ";TRUE"
                    .ParameterFields(11) = "colnm;" & strColNm & ";TRUE"
                    .ParameterFields(12) = "unitqty;" & strUnitQty & ";TRUE"
                    .ParameterFields(13) = "spcno;" & strSpcNo & ";TRUE"
                    .ParameterFields(14) = "store;" & strStore & ";TRUE"
                    .ParameterFields(15) = "accno;" & strAccNo & ";TRUE"
                    .ParameterFields(16) = "accdttm;" & strAccdttm & ";TRUE"
                    .ParameterFields(17) = "accnm;" & strAccNm & ";TRUE"
                    .ParameterFields(18) = "hostnm;" & HOSPITAL_NAME & ";TRUE"
                    .ParameterFields(19) = "reltest;" & IIf(strRelTest = "", "(없음)", "") & ";TRUE"
                    .ParameterFields(21) = "prtnm;" & GetEmpNm(ObjSysInfo.EmpId) & ";TRUE"
                    
                    If strRelTest <> "" Then
                        strabScreen = ""
                        strdCoombs = ""
                        For j = LBound(aryRelTest) To UBound(aryRelTest)
                            If aryRelTest(j) <> "" Then
                                k = k + 1

' 2009.06.16. 양성현 And strdCoombs = ""  추가
' 일자별로 조회되기때문에 제일 처음에 선택되는 것이 가장 최근값이다.
' 이전검사조회 기간은 현재 설정된 값은 3600일전임
'                                If k < 13 Then Exit For
'                                .ParameterFields(22 + j) = "reltest" & (j + 1) & ";" & aryRelTest(j) & ";TRUE"
                                If k < 13 Then .ParameterFields(22 + j) = "reltest" & (j + 1) & ";" & aryRelTest(j) & ";TRUE"
                                strTemp = Trim(medGetP(Mid(aryRelTest(j), 13), 1, ":"))

'2015.09.15 온승호 Ab Screening 최근 결과 조회
'Ab-id 검사항목명이 있음
'                                If Mid(strTemp, 1, 3) = "Ab " And strabScreen = "" Then
                                If InStr(strTemp, "Ab ") > 0 And strabScreen = "" Then
                                    If Val(Trim(medGetP(strabScreen, 2, "-"))) < Val(Trim(medGetP(aryRelTest(j), 2, "-"))) Then
                                        strabScreen = Mid(aryRelTest(j), 1, 13) & " : " & Trim(medGetP(aryRelTest(j), 2, ":"))
                                    End If

                                End If
                                If Mid(strTemp, 1, 4) = "Coom" And strdCoombs = "" Then
                                    If Val(Trim(medGetP(strdCoombs, 2, "-"))) < Val(Trim(medGetP(aryRelTest(j), 2, "-"))) Then
                                        strdCoombs = Mid(aryRelTest(j), 1, 13) & " : " & Trim(medGetP(aryRelTest(j), 2, ":"))
                                    End If

'                                    strdCoombs = Mid(aryRelTest(j), 1, 13) & " : " & Trim(medGetP(aryRelTest(j), 2, ":"))
'                                    .ParameterFields(23) = "dCooms;" & Mid(aryRelTest(j), 1, 13) & " : " & Trim(medGetP(aryRelTest(j), 2, ":")) & ";TRUE"
'                                    Debug.Print strTemp
                                End If
                            End If
                        Next j
                        .ParameterFields(22 + j) = "abScreen;" & strabScreen & ";TRUE"
                        .ParameterFields(22 + j + 1) = "dCooms;" & strdCoombs & ";TRUE"
                    End If
                    
                    .RetrieveDataFiles
    '                .WindowState = crptMaximized
                    .Destination = crptToPrinter
                    .Action = 1
                    .Reset
                End With
            End If
        Next i
    End With
    
    Me.MousePointer = 0
End Sub

Private Sub GetColAccInfo(ByVal vSpcNo As String, _
                          ByRef pColDtTm As String, ByRef pColId As String, _
                          ByRef pAccDtTm As String, ByRef pAccId As String)

    'S2bbs201에서 채혈, 접수정보를 읽는다
    
    Dim RS As Recordset
    Dim strSQL As String
    
    Set RS = New Recordset
    strSQL = " select * from " & T_BBS201 & _
             " where " & DBW("spcyy=", medGetP(vSpcNo, 1, "-")) & _
             " and " & DBW("spcno=", medGetP(vSpcNo, 2, "-"))
    RS.Open strSQL, DBConn
    
    If RS.EOF = False Then
        pColDtTm = Format(RS.Fields("coldt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("coltm").value & "", 1, 4), "00:00")
        If RS.Fields("colid").value & "" <> "" Then
            pColId = GetEmpNm(RS.Fields("colid").value & "")
        End If
        pAccDtTm = Format(RS.Fields("rcvdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("rcvtm").value & "", 1, 4), "00:00")
        If RS.Fields("rcvid").value & "" <> "" Then
            pAccId = GetEmpNm(RS.Fields("rcvid").value & "")
        End If
    End If
    
    Set RS = Nothing
End Sub

Private Function GetSexAge(ByVal vPtID As String)
    Dim objPt As clsPatient
    
    Set objPt = New clsPatient
    
    Call objPt.GETPatient(vPtID)
    
    GetSexAge = objPt.sexage
    
    Set objPt = Nothing
End Function

Private Function GetRelTest(ByVal vPtID As String) As String
'크리스탈 리포트 출력용 스트링 만들기..

    Dim RS As Recordset
    Dim objSql As clsCrossMatching
    Dim strTmp As String
    Dim lngCnt As Long
    Dim strRstCd As String
    
    Set objSql = New clsCrossMatching
    Set RS = New Recordset
    
    RS.Open objSql.TestResultXM(vPtID), DBConn
        
    If RS.EOF Then
        GetRelTest = ""
    Else
        Do Until RS.EOF
            lngCnt = lngCnt + 1

' 2009.06.16. 양성현 오래전에 검사한  Ab Screen 등을 가져오지 못해서 막아버림.
'
'            If lngCnt > 12 Then Exit Do
'
            If RS.Fields("rstcdnm").value & "" = "" Then
                strRstCd = RS.Fields("rstcd").value & ""
            Else
                strRstCd = RS.Fields("rstcdnm").value & ""
            End If
            
            strTmp = strTmp & Format(RS.Fields("workarea").value & "" & "-" & _
                                     Mid(RS.Fields("accdt").value & "", 3) & "-" & _
                                     RS.Fields("accseq").value & "", "!" & String(17, "@")) & _
                     Format(RS.Fields("abbrnm10").value & "", "!" & String(11, "@")) & " : " & _
                     strRstCd & vbTab
            RS.MoveNext
        Loop
        
        strTmp = strTmp & vbNewLine
        
        GetRelTest = strTmp
    End If
    
    Set RS = Nothing
    Set objSql = Nothing
End Function

Private Sub LoadBuilding()
    
    Dim objcom003   As clsCom003
    Dim RS          As Recordset
    Dim i           As Long
    Dim itmX        As ListItem
    
    Set objcom003 = New clsCom003
    Set RS = objcom003.OpenRecordSet(BC2_CENTER)
    Set objcom003 = Nothing
    
    cboBuilding.Clear
    cboBuilding.AddItem "(전체)"
    If Not RS.EOF Then
        With RS
            For i = 1 To .RecordCount
                cboBuilding.AddItem .Fields("cdval1").value & " " & .Fields("field1").value & ""
                .MoveNext
            Next i
        End With
    End If
    Set RS = Nothing
    If cboBuilding.ListCount > 1 Then
        cboBuilding.ListIndex = medComboFind(cboBuilding, ObjSysInfo.BuildingCd)
    Else
        cboBuilding.ListIndex = 0
    End If
    
End Sub

'2001-11-30 추가
'출고전표 출력을 위한 Query
Private Sub QueryForReport()
    Dim i           As Long
    Dim j           As Long
    
    Dim RS        As Recordset
    Dim RsTime      As Recordset
    Dim QueryOrder  As clsQueryOrder
    Dim objDisease  As clsDisease
    Dim ObjABO      As clsABO
    
    Dim accno       As String
    Dim reason      As String
    Dim status      As String
    Dim spcno       As String
    Dim storeleg    As String
    Dim storerow    As String
    Dim storecol    As String
    Dim center      As String
    
    Dim strLeg      As String
    Dim strRow      As String
    Dim strCol      As String
    Dim inout       As String
    Dim MaxRowCnt   As Long
    Dim TestDiv     As String
'    Dim blnComplete As Boolean
    
    Dim objPrgBar   As clsProgress
    
    Dim otherCenter As Boolean
    
    '윗줄과 같은내용이면 글자를 감추기 위한변수들
    Dim bkPtId      As String
    Dim bkReason    As String
    Dim bkReqDt     As String
    Dim bkOrdDt     As String
    Dim bkRoomid    As String
    Dim bkWard      As String
    Dim bkDept      As String
    
    Dim strDc       As String
    
    Dim blnCompleted As Boolean
    Dim blnAccomplished As Boolean
    
    tblPtList.MaxRows = 0
    
    Call Save_LegRowCol
    
    Set QueryOrder = New clsQueryOrder
    
    If cboOrd.ListIndex <> 0 Then TestDiv = medGetP(cboOrd.Text, 1, " ")
    '-----------
    '상태별 조회
    '-----------

    QueryOrder.stscd = "'3','4'"

    '------------------------------------
    '출고전표출력에 맞는 조건으로 초기화
    cboInOut.ListIndex = 0
    chkDc.value = 0
    chkStat.value = 0
    txtWardId = ""
    cboOrd.ListIndex = -1
    
    inout = ""
    strDc = ""
    TestDiv = ""
    '------------------------------------
    
        
    
    
    Set RS = QueryOrder.QueryRequest(Format(dtpFrDt, PRESENTDATE_FORMAT), Format(dtpToDt, PRESENTDATE_FORMAT), _
                                      chkStat.value, txtPtId, inout, strDc, txtWardId, TestDiv)
    
    If RS Is Nothing Then
        Set RS = Nothing
        Set QueryOrder = Nothing
        Exit Sub
    End If
    
    Set objDisease = New clsDisease
    Set ObjABO = New clsABO
    
    Set objPrgBar = New clsProgress
    objPrgBar.Container = medMain.stsBar
    
    objPrgBar.Min = 1
    objPrgBar.Max = RS.RecordCount
    
    
    With tblPtList
        bkPtId = ""
        .ReDraw = False
        For i = 1 To RS.RecordCount
        
            objPrgBar.value = i
'            blnComplete = CompleteOrderChk(Rs.Fields("accdt").value & "", Rs.Fields("accseq").value & "", Rs.Fields("unitqty").value & "")
            Call CheckCompleted(RS.Fields("accdt").value & "", RS.Fields("accseq").value & "", RS.Fields("unitqty").value & "", _
                                blnCompleted, blnAccomplished)
            If blnCompleted = True Then GoTo Skip

Skip1:
            
            MaxRowCnt = MaxRowCnt + 1
            .MaxRows = MaxRowCnt
            .Row = MaxRowCnt
            accno = Trim(RS.Fields("accdt").value & "") & "-" & Val(Trim(RS.Fields("accseq").value & ""))
            If accno = "-0" Then accno = "" 'accno = "미접수"
            
            '수혈사유 구하기...
            reason = QueryOrder.GetTransReason(RS.Fields("ptid").value & "", RS.Fields("orddt").value & "", RS.Fields("ordno").value & "")
            
            
            If reason = "" Then reason = "(없음)"
            

            
            .Col = TblColumn.tcACCNO:      .value = accno
            .Col = TblColumn.tcPTID:       .value = RS.Fields("ptid").value & ""
            
            .Col = TblColumn.tcPTNM:       .value = RS.Fields("ptnm").value & ""
            .Col = TblColumn.tcORDNM:      .value = RS.Fields("testnm").value & ""
            .Col = TblColumn.tcORDDT:      .value = Format(RS.Fields("orddt").value & "", "####-##-##")
            '.Col = TblColumn.tcUNITQTY:    .value = RS.Fields("unitqty").value & ""
            .Col = TblColumn.tcUNITQTY:    .value = RS.Fields("reqcnt").value & ""
            .Col = TblColumn.tcREASON:     .value = Trim(Trim0(reason))
            .Col = TblColumn.tcREQDT:      .value = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
            '2001-11-30추가
            '출고전표에 담당의사/최근수혈일 출력하기위함
            .Col = TblColumn.tcDOCT:       .value = RS.Fields("orddoct").value & ""
            .Col = TblColumn.tcTRANSDT:    .value = QueryOrder.GetLatestTrandDt(RS.Fields("ptid").value & "")
'
            .Col = TblColumn.tcWARD:       .value = RS.Fields("wardid").value & ""
            .Col = TblColumn.tcROOM:       .value = RS.Fields("hosilid").value & ""
            
            .Col = TblColumn.tcDEPT:       .value = RS.Fields("deptcd").value & ""
            .Col = TblColumn.tcBUSSDIV:    .value = RS.Fields("bussdiv").value & ""
            .Col = TblColumn.tcORDDTDB:    .value = RS.Fields("orddt").value & ""
            .Col = TblColumn.tcORDNO:      .value = Val(RS.Fields("ordno").value & "")
            .Col = TblColumn.tcORDSEQ:     .value = Val(RS.Fields("ordseq").value & "")
            .Col = TblColumn.tcSTATFG:     .value = RS.Fields("statfg").value & ""
            .Col = TblColumn.tcSTATnm:     .value = IIf(RS.Fields("statfg").value & "" = "1", "Y", "")
                                           .ForeColor = vbRed
                                           .FontBold = True
            .Col = TblColumn.tcBedInDT:    .value = RS.Fields("bedindt").value & ""
            .Col = TblColumn.tcDCFG:       .value = RS.Fields("dcfg").value & ""
            .Col = TblColumn.tcDCNM:       .value = IIf(RS.Fields("dcfg").value & "" = "1", "Y", "")
                                           .ForeColor = vbBlue
                                           .FontBold = True
            '.Col = TblColumn.tcCENTERCD:   .value = center
            .Col = TblColumn.tcPHERESIS:   .value = RS.Fields("testdiv").value & ""
            .Col = TblColumn.tcSTSCD:      .value = RS.Fields("stscd").value & ""
            .Col = TblColumn.tcSTSNM
                                            If TRANS_REQUIRE_USED Then
                                                    Select Case RS.Fields("stscd").value & ""
                                                         Case "0": .value = STS_NM_ORDER: .ForeColor = DCM_Gray '"처방"
                                                         Case "1": .value = STS_NM_COLLECT '"채혈"
                                                         Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"접수"
                                                         '2001-11-15 수정 : '요청' Status 추가
                                                         Case "3": .value = STS_NM_REQUEST: .ForeColor = DCM_Red '"요청"
                                                         Case "4": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS) '"종결","완료","검사중"
                                                                   .ForeColor = IIf(blnCompleted, &H8000&, DCM_Brown)
                                                         'Case "3": .value = "검사중"
                                                         Case Else: .value = ""
                                                    End Select
                                            Else
                                                    Select Case RS.Fields("stscd").value & ""
                                                         Case "0": .value = STS_NM_ORDER '"처방"
                                                         Case "1": .value = STS_NM_COLLECT: .ForeColor = DCM_LightRed '"채혈"
                                                         Case "2": .value = STS_NM_ACCESS: .ForeColor = DCM_LightBlue '"접수"
                                                         Case "3": .value = IIf(blnCompleted = True, IIf(blnAccomplished, STS_NM_END, STS_NM_DONE), STS_NM_INPROGRESS): .ForeColor = DCM_Brown ''"종결","완료","검사중"
                                                         'Case "3": .value = "검사중"
                                                         Case Else: .value = ""
                                                    End Select
                                            End If
            .Col = TblColumn.TcMESG: .value = RS.Fields("mesg").value & ""
            '혈액형을 구한다.
            ObjABO.Ptid = RS.Fields("ptid").value & ""
            
            If ObjABO.GetABO = False Then
                .Col = TblColumn.tcABO:    .value = ""
            Else
                .Col = TblColumn.tcABO:    .value = ObjABO.ABO & ObjABO.Rh
            End If
            
            '진단명을 구한다.
            With objDisease
                .Clear
                .Ptid = RS.Fields("ptid").value & ""
                .OrdDt = RS.Fields("orddt").value & ""
                .ordno = RS.Fields("ordno").value & ""
            End With
            
            
            If objDisease.GetDisease = False Then
                .Col = TblColumn.tcDISEASE: .value = ""
                .Col = TblColumn.tcDISEASE2: .value = ""
                .Col = TblColumn.tcDISEASE3: .value = ""
                .Col = TblColumn.tcDISEASE4: .value = ""
            Else
                j = 0
                Do
                    If objDisease.EOF Then Exit Do
                    
                    If objDisease.DiseaseCd <> "" Then
                        j = j + 1
                        Select Case j
                            Case 1: .Col = TblColumn.tcDISEASE
                            Case 2: .Col = TblColumn.tcDISEASE2
                            Case 3: .Col = TblColumn.tcDISEASE3
                            Case 4: .Col = TblColumn.tcDISEASE4
                        End Select
                        .value = objDisease.DiseaseCd & " " & objDisease.DiseaseNm
                    End If
                    objDisease.MoveNext
                Loop
            End If
            
            otherCenter = False
            
            
            '-------------------------------------------
            '처방이 irradiation 처방이 아닌 처방일경우만
            '-------------------------------------------
            Call QueryOrder.GetSpcNoAndStore(RS.Fields("ptid").value & "", spcno, storeleg, storerow, storecol, center)
            
            '--------------------------------------------------------------------------------------
            '2001-11-23 추가 :
            '건물정보를 사용할 경우, 그리고 (전체)가 아닐경우 해당 건물의 데이타가 디스플레이
            If ObjSysInfo.UseBuildingInfo = 1 And cboBuilding.ListIndex <> 0 Then
                If medGetP(center, 1, vbTab) <> ObjSysInfo.BuildingCd Then
                    MaxRowCnt = MaxRowCnt - 1
                    .MaxRows = MaxRowCnt
                    GoTo Skip
                End If
            End If
            '--------------------------------------------------------------------------------------
            
            If center = "" Then center = ObjSysInfo.BuildingCd & vbTab & ObjSysInfo.BuildingNm
            .Col = TblColumn.tcCENTERNM:    .value = medGetP(center, 2, vbTab) 'GetCenterNm(medGetP(center, 1, vbTab))
            .Col = TblColumn.tcCENTERCD:    .value = medGetP(center, 1, vbTab)
            
            If medGetP(center, 1, vbTab) <> ObjSysInfo.BuildingCd Then
                '검체가 다른 센터에 있다.
                .Col = TblColumn.tcSTORE:   .value = medGetP(center, 2, vbTab) & "(" & medGetP(center, 1, vbTab) & ")"
                otherCenter = True
            End If
            
            .Col = TblColumn.tcORDDIV:      .value = RS.Fields("orddiv").value & ""
            If .value = C_WORKAREA Then
                '--------------------------
                '검체번호와 보관장소 구하기
                '--------------------------
                If storerow = "0" Then storerow = ""
                If storecol = "0" Then storecol = ""
                
                .Col = TblColumn.tCLegRowCol:   .value = storeleg & ";" & storerow & ";" & storecol
                
                .Col = TblColumn.tcSPCNO:       .value = spcno
                
                If spcno = "" Then
                    .Col = TblColumn.tcSTORE:   .value = "" '.value = "미채혈"
                Else
                    If storeleg = "" Then
                        .Col = TblColumn.tcSTORE:    .value = ""
                        .Col = TblColumn.tcNOACCSSS: .value = "1"
                    Else
                        .Col = TblColumn.tcSTORE:    .value = storeleg & "(" & storerow & "," & storecol & ")"
                        .Col = TblColumn.tcNOACCSSS: .value = "0"
                    End If
                End If
                '----------------------------
                '검체경과 시간을 구하기위해서
                '----------------------------
                Dim today   As Date
                Dim coldttm As String
                today = GetSystemDate
                
                If spcno <> "" Then
                    If Val(RS.Fields("stscd").value & "") > 2 Then
                        If QueryOrder.Get_ExistSPC(medGetP(spcno, 1, "-"), medGetP(spcno, 2, "-")) <> "1" Then
                            .Col = TblColumn.tcSPCNO: .ForeColor = DCM_LightGray
                            .Col = TblColumn.tcSTORE: .ForeColor = DCM_LightGray
                        End If
                    End If
                    Set RsTime = Nothing
                    Set RsTime = New Recordset
                    RsTime.Open QueryOrder.Get_spcTime(medGetP(spcno, 1, "-"), medGetP(spcno, 2, "-")), DBConn
                    
                    If Not RsTime.EOF Then
                        If Len(RsTime.Fields("coltm").value & "") = 4 Then
                            coldttm = RsTime.Fields("coltm").value & "" & "00"
                            coldttm = Format(RsTime.Fields("coldt").value & "", "0###-##-##") & " " & Format(coldttm, "0#:##:##")
                        Else
                            coldttm = Format(RsTime.Fields("coldt").value & "", "0###-##-##") & " " & Format(RsTime.Fields("coltm").value & "", "0#:##:##")
                        End If
                        
                       ' coldttm = Format(RsTime.Fields("coldt").value, "0###-##-##") & " " & Format(coldttm, "0#:##:##")
                        .Col = TblColumn.tcTime: .value = DateDiff("h", coldttm, today)
                    End If
                End If
            End If
            
            
            .Col = TblColumn.tcDUPCHK: .value = RS.Fields("ptid").value & "" & COL_DIV & RS.Fields("orddt").value & ""
            
            '-------------------------
            '중복되는 값은 안보이게...
            '-------------------------
            
            If bkPtId <> RS.Fields("ptid").value & "" Then
                bkPtId = RS.Fields("ptid").value & ""
                bkReason = reason
                bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
                bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                bkRoomid = RS.Fields("hosilid").value & ""
                bkWard = RS.Fields("wardid").value & ""
                bkDept = RS.Fields("deptcd").value & ""
                
            Else
                .Row = i - 1
                .Col = TblColumn.tcWARD: bkWard = .value
                .Col = TblColumn.tcDEPT: bkDept = .value
                
                .Row = i
                .Col = TblColumn.tcPTID: .ForeColor = .BackColor
                .Col = TblColumn.tcPTNM: .ForeColor = .BackColor
                If bkReason = reason Then
                    If reason <> "(없음)" Then .Col = TblColumn.tcREASON: .ForeColor = .BackColor
                Else
                    bkReason = reason
                End If
                If bkWard = RS.Fields("wardid").value & "" Then
                    .Col = TblColumn.tcWARD: .ForeColor = .BackColor
                End If
                If bkDept = RS.Fields("deptcd").value & "" Then
                    .Col = TblColumn.tcDEPT: .ForeColor = .BackColor
                End If
                
                If bkRoomid = RS.Fields("hosilid").value & "" Then
                    .Col = TblColumn.tcROOM: .ForeColor = .BackColor
                Else
                    bkRoomid = RS.Fields("hosilid").value & ""
                End If
                If bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00") Then
                    .Col = TblColumn.tcREQDT: .ForeColor = .BackColor
                Else
                    bkReqDt = Format(RS.Fields("reqdt").value & "", "####-##-##") & " " & Format(Mid(RS.Fields("reqtm").value & "", 1, 4), "00:00")
                End If
                If bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##") Then
                    .Col = TblColumn.tcORDDT: .ForeColor = .BackColor
                Else
                    bkOrdDt = Format(RS.Fields("orddt").value & "", "####-##-##")
                End If
            End If
Skip:
            '---------------------
            '접수할 수 있는 건인지
            '---------------------
            If MaxRowCnt > 0 Then
                If CanSelect(1, MaxRowCnt) Then
                    .Row = MaxRowCnt
                    .Col = TblColumn.tcSEL
                    .CellType = CellTypeCheckBox
                    .TypeCheckCenter = True
                Else
                    .Row = MaxRowCnt
                    
                    .Col = TblColumn.tcSEL
                    .CellType = CellTypeStaticText
                    .Col = TblColumn.tcSTSNM
                    If .value = STS_NM_DONE Or .value = STS_NM_END Then
                        .Col = TblColumn.tcSEL
                        .Text = "√"
                        .ForeColor = vbRed
                    End If
                End If
            End If
            RS.MoveNext
        Next i
        .ReDraw = True
    End With
     
    Set RS = Nothing
    Set ObjABO = Nothing
    Set objPrgBar = Nothing
    Set objDisease = Nothing
    Set QueryOrder = Nothing
End Sub


