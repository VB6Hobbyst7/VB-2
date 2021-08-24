VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm467TestTAT 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  '없음
   Caption         =   "TAT 목표달성율"
   ClientHeight    =   9165
   ClientLeft      =   585
   ClientTop       =   915
   ClientWidth     =   15345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Frame Frame3 
      Caption         =   "TAT 구간"
      Height          =   555
      Left            =   4020
      TabIndex        =   42
      Top             =   630
      Width           =   2085
      Begin VB.OptionButton Option6 
         Caption         =   "접수"
         Height          =   255
         Left            =   1200
         TabIndex        =   44
         Top             =   240
         Width           =   705
      End
      Begin VB.OptionButton Option5 
         Caption         =   "채혈"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   90
      Style           =   1  '그래픽
      TabIndex        =   38
      Tag             =   "127"
      Top             =   8580
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&S)"
      Height          =   510
      Left            =   11835
      Style           =   1  '그래픽
      TabIndex        =   37
      Tag             =   "158"
      Top             =   675
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Caption         =   "조회유형"
      Height          =   555
      Left            =   90
      TabIndex        =   30
      Top             =   630
      Width           =   3825
      Begin VB.OptionButton Option4 
         Caption         =   "외래"
         Height          =   180
         Left            =   2160
         TabIndex        =   36
         Top             =   270
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "ER"
         Height          =   180
         Left            =   3075
         TabIndex        =   33
         Top             =   270
         Width           =   645
      End
      Begin VB.OptionButton Option2 
         Caption         =   "응급"
         Height          =   180
         Left            =   1170
         TabIndex        =   32
         Top             =   270
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "병동"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   270
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   7260
      Left            =   90
      ScaleHeight     =   7200
      ScaleWidth      =   14265
      TabIndex        =   29
      Top             =   1215
      Width           =   14325
      Begin VB.Frame Frame2 
         Height          =   7125
         Left            =   1830
         TabIndex        =   39
         Top             =   30
         Width           =   11415
         Begin VB.CommandButton cmdExcel1 
            BackColor       =   &H00F4F0F2&
            Caption         =   "Excel"
            Height          =   450
            Left            =   10020
            Style           =   1  '그래픽
            TabIndex        =   41
            Tag             =   "127"
            Top             =   6600
            Width           =   1320
         End
         Begin FPSpread.vaSpread spdRstList 
            Height          =   6330
            Left            =   60
            TabIndex        =   40
            Tag             =   "45506"
            Top             =   180
            Width           =   11295
            _Version        =   196608
            _ExtentX        =   19923
            _ExtentY        =   11165
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            BackColorStyle  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            MaxCols         =   9
            OperationMode   =   1
            Protect         =   0   'False
            ScrollBars      =   2
            ShadowColor     =   14737632
            SpreadDesigner  =   "Lis467.frx":0000
            VisibleCols     =   5
            VisibleRows     =   500
         End
      End
      Begin FPSpread.vaSpread spdTemp 
         Height          =   3855
         Left            =   60
         TabIndex        =   35
         Top             =   2700
         Width           =   7815
         _Version        =   196608
         _ExtentX        =   13785
         _ExtentY        =   6800
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         SpreadDesigner  =   "Lis467.frx":1B9B
      End
      Begin FPSpread.vaSpread spdStat 
         Height          =   7215
         Left            =   -60
         TabIndex        =   34
         Top             =   0
         Width           =   14340
         _Version        =   196608
         _ExtentX        =   25294
         _ExtentY        =   12726
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "Lis467.frx":1D7C
      End
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Refresh"
      Height          =   510
      Left            =   13155
      Style           =   1  '그래픽
      TabIndex        =   26
      Tag             =   "158"
      Top             =   675
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   9165
      Style           =   1  '그래픽
      TabIndex        =   16
      Tag             =   "127"
      Top             =   8580
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   10485
      Style           =   1  '그래픽
      TabIndex        =   15
      Tag             =   "132"
      Top             =   8580
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13125
      Style           =   1  '그래픽
      TabIndex        =   14
      Tag             =   "128"
      Top             =   8580
      Width           =   1320
   End
   Begin VB.Frame frmPrgBar 
      BackColor       =   &H00AFBCC5&
      BorderStyle     =   0  '없음
      Caption         =   "                                                                                    "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000F5386&
      Height          =   1035
      Left            =   4545
      TabIndex        =   11
      Top             =   4305
      Visible         =   0   'False
      Width           =   6525
      Begin MSComctlLib.ProgressBar Prgbar 
         Height          =   225
         Left            =   60
         TabIndex        =   12
         Top             =   720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00DBE6E6&
         Height          =   1035
         Left            =   0
         Top             =   0
         Width           =   6525
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00A9B4BA&
         BackStyle       =   0  '투명
         Caption         =   "데이터를 로드중 입니다."
         Height          =   180
         Left            =   2355
         TabIndex        =   13
         Top             =   300
         Width           =   1980
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   1170
      TabIndex        =   6
      Top             =   0
      Width           =   13305
      Begin VB.TextBox txtCnt 
         Height          =   330
         Left            =   11700
         TabIndex        =   24
         Top             =   195
         Width           =   735
      End
      Begin VB.ComboBox cboWorkArea 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Left            =   7875
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   195
         Width           =   2565
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   765
         TabIndex        =   7
         Top             =   195
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   86245376
         CurrentDate     =   36238
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   3780
         TabIndex        =   8
         Top             =   195
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   86245376
         CurrentDate     =   36391
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   510
         Left            =   6705
         TabIndex        =   22
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "WorkArea"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   510
         Left            =   10530
         TabIndex        =   23
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "목표달성율"
         Appearance      =   0
      End
      Begin VB.Label Label5 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12510
         TabIndex        =   25
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3405
         TabIndex        =   9
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11805
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "128"
      Top             =   8580
      Width           =   1320
   End
   Begin VB.Frame fraCond 
      Appearance      =   0  '평면
      BackColor       =   &H00DBE6E6&
      Caption         =   "검사실"
      ForeColor       =   &H00864B24&
      Height          =   1110
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   8805
      Visible         =   0   'False
      Width           =   555
      Begin VB.ComboBox cboBuilding 
         BackColor       =   &H00F1F5F4&
         Height          =   300
         Left            =   270
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   660
         Width           =   2895
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전  체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   345
         Width           =   1035
      End
      Begin VB.ComboBox cboSort 
         BackColor       =   &H00F7FFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "Lis467.frx":22C1
         Left            =   2385
         List            =   "Lis467.frx":22D7
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Tag             =   "검사실"
         Top             =   315
         Width           =   780
      End
      Begin VB.CheckBox chkSubTot 
         BackColor       =   &H00DBE6E6&
         Caption         =   "소 계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   1590
         TabIndex        =   1
         Top             =   345
         Width           =   795
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   1605
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   3990
      TabIndex        =   17
      Top             =   5310
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   810
      Left            =   3420
      TabIndex        =   18
      Top             =   2385
      Visible         =   0   'False
      Width           =   1485
      _Version        =   196608
      _ExtentX        =   2619
      _ExtentY        =   1429
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "Lis467.frx":22F0
   End
   Begin MedControls1.LisLabel lblCondition 
      Height          =   510
      Left            =   60
      TabIndex        =   19
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   900
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "조회기간"
      Appearance      =   0
   End
   Begin VB.Label Label4 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ToTal Count : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7290
      TabIndex        =   28
      Top             =   810
      Width           =   1725
   End
   Begin VB.Label lblTotalCnt 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00DBE6E6&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   9060
      TabIndex        =   27
      Top             =   810
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00A9B4BA&
      BackStyle       =   0  '투명
      Caption         =   "데이터를 로드중 입니다."
      Height          =   180
      Left            =   6285
      TabIndex        =   20
      Top             =   4845
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   1  '투명하지 않음
      Height          =   1035
      Left            =   3915
      Top             =   4500
      Visible         =   0   'False
      Width           =   6525
   End
End
Attribute VB_Name = "frm467TestTat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_BUILDING = 5
Private Const COL_WORKAREA = 1
Private Const COL_EQUIPMENT = 2
Private Const COL_DEPTNM = 3
Private Const COL_TESTNM = 4
Private Const COL_COUNT = 6
Private Const COL_SERIES = 7
Private Const COL_POINTS = 8

'Private WithEvents objMyList    As clsS2DLP
Private objSQL          As New clsLISSqlStatistic
Dim rsDeptStat          As Recordset
Dim rsTestStat          As Recordset
Dim rsDeptTestStat      As Recordset
Dim QueryFlag           As Boolean
Dim MsgFg               As Boolean

Dim ColWid(6)           As Double
Dim SortKeys(6)         As Integer
Dim totCnt              As Long
Dim GrpColor(100)       As Long
Dim iPrgbarCount        As Long
Dim SubTot(6)           As Long

'Workarea별 검사코드 담아주는 Dictionary
Private objDic As clsDictionary
Public Event LastFormUnload()

Private Sub cboBuilding_Click()
    Call LoadEqpList
End Sub

Private Sub cboWorkArea_Click()
    Dim RS        As Recordset
    Dim sWorkarea As String
    Dim SSQL      As String
    
    Set objDic = New clsDictionary
    objDic.Clear
    objDic.FieldInialize "testcd", "testnm"
    
    If cboWorkArea.ListIndex < 0 Then Exit Sub
    
    sWorkarea = medGetP(cboWorkArea.Text, 1, " ")
    
    Set RS = New Recordset
    SSQL = objSQL.GetWorkareaTestItem(sWorkarea)
    
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Do Until RS.EOF
            If objDic.Exists(RS.Fields("testcd").Value & "") = False Then
                objDic.AddNew RS.Fields("testcd").Value & "", RS.Fields("testnm").Value & ""
            End If
            RS.MoveNext
        Loop
    End If
    RS.Close
    Set RS = Nothing
End Sub

Private Sub cfxStat_LButtonUp(ByVal X As Integer, ByVal Y As Integer, nRes As Integer)
'MsgBox nRes
End Sub

Private Sub chkAll_Click(Index As Integer)
    Dim ChkValue As Boolean

    ChkValue = IIf(chkAll(Index).Value = 0, True, False)
    Select Case Index
    Case 0:
        cboBuilding.Enabled = ChkValue
    Case 1:
        cboWorkArea.Enabled = ChkValue
    Case 2:
'        cboEqpCd.Enabled = ChkValue
    Case 3:
'        txtDeptCd.Text = ""
'        txtDeptCd.Enabled = ChkValue
'        cmdHelpList(0).Enabled = ChkValue
'        lblDeptNm.Caption = ""
    Case 4:
'        txtTestCd.Text = ""
'        txtTestCd.Enabled = ChkValue
'        cmdHelpList(1).Enabled = ChkValue
'        lblTestNm.Caption = ""
    End Select
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn
'    dtpStart.Value = Now
'    dtpEnd.Value = Now
    dtpStart.SetFocus
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp      As String
    
    With spdStat
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblexcel.MaxRows = .MaxRows + 1
        tblexcel.MaxCols = .MaxCols
        tblexcel.Row = 1: tblexcel.Row2 = tblexcel.MaxRows
        tblexcel.Col = 1: tblexcel.COL2 = tblexcel.MaxCols
        tblexcel.BlockMode = True
        tblexcel.Clip = strTmp
        tblexcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "TestTATCount"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)

End Sub

Private Sub cmdExcel1_Click()
    Dim strTmp      As String
    
    With spdRstList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblexcel.MaxRows = .MaxRows + 1
        tblexcel.MaxCols = .MaxCols
        tblexcel.Row = 1: tblexcel.Row2 = tblexcel.MaxRows
        tblexcel.Col = 1: tblexcel.COL2 = tblexcel.MaxCols
        tblexcel.BlockMode = True
        tblexcel.Clip = strTmp
        tblexcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "TestTATLIST"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdRefresh_Click()
    Call ShowData
End Sub

Private Sub cmdSave_Click()
    Dim iCnt As Long
    Dim sMonth, sWorkarea, sDiv, sTestCd, sTestNm, sWordTm, sEmTm, sOutTm, sTarget, sHTm As String
    Dim sTCnt, sOutCnt, sMark, sEntId, sEntDt, sEmtTm As String
    Dim varTmp
    Dim strSQL As String
    Dim tmpRs   As Recordset
    Dim sDivNm As String
    
'ENT_MONTH , WorkArea, TEST_DIV, TEST_CD, TEST_NM, WARD_TIME, EM_TIME,
'OUT_TIME , Target, TOTAL_CNT, OUT_CNT, MARK, ENT_ID, ENT_DT, EMT_TM
    
    With spdStat
        sMonth = Format(dtpStart.Value, "YYYYMM")
        sEntId = ObjSysInfo.EmpId
        sEntDt = Format(GetSystemDate, "YYYYMMDD")
        sEmtTm = Format(GetSystemDate, "hhmmss")
        If Option1.Value = True Then
            sDiv = "1": sDivNm = Option1.Caption
        ElseIf Option2.Value = True Then
            sDiv = "2": sDivNm = Option1.Caption
        ElseIf Option4.Value = True Then
            sDiv = "4": sDivNm = Option1.Caption
        Else
            sDiv = "3": sDivNm = Option1.Caption
        End If
                                
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM S2LAB910 WHERE ENT_MONTH = '" & sMonth & "' AND TEST_DIV = '" & sDiv & "' AND WORKAREA = '" & Mid(cboWorkArea.Text, 1, 2) & "'"
        
        Set tmpRs = New Recordset
        
        tmpRs.Open strSQL, DBConn
        
        If tmpRs.RecordCount > 0 Then
        
            Select Case MsgBox(sMonth & "월의 " & Trim(Mid(cboWorkArea.Text, 3)) & " " & sDivNm & "데이타는 이미 저장되어있습니다. 삭제하고 저장 하시겠습니까?", vbYesNo Or vbInformation Or vbDefaultButton1, App.Title)
                Case vbYes
                    strSQL = ""
                    strSQL = strSQL & " DELETE S2LAB910 WHERE ENT_MONTH = '" & sMonth & "' AND TEST_DIV = '" & sDiv & "' AND WORKAREA = '" & Mid(cboWorkArea.Text, 1, 2) & "'"
                    DBConn.Execute strSQL
                    
                    For iCnt = 1 To .MaxRows - 1
                        .GetText 1, iCnt, varTmp: sTestCd = varTmp
                        .GetText 2, iCnt, varTmp: sTestNm = varTmp
                        .GetText 3, iCnt, varTmp: sWordTm = varTmp
                        .GetText 4, iCnt, varTmp: sEmTm = varTmp
                        .GetText 5, iCnt, varTmp: sOutTm = varTmp
                        .GetText 6, iCnt, varTmp: sHTm = varTmp
                        .GetText 7, iCnt, varTmp: sTarget = varTmp
                        .GetText 8, iCnt, varTmp: sTCnt = varTmp
                        .GetText 9, iCnt, varTmp: sOutCnt = varTmp
                        .GetText 10, iCnt, varTmp: sMark = varTmp
                        
                        If InStr(sTestNm, ",") > 0 Then
                            sTestNm = Replace(sTestNm, "'", "")
                        End If
                        
                        strSQL = ""
                        strSQL = strSQL & "INSERT INTO S2LAB910 (ENT_MONTH , WorkArea, TEST_DIV, TEST_CD, TEST_NM, WARD_TIME, EM_TIME, " & vbCr
                        strSQL = strSQL & " OUT_TIME , Target, TOTAL_CNT, OUT_CNT, MARK, ENT_ID, ENT_DT, EMT_TM, H_TIME)" & vbCr
                        strSQL = strSQL & " VALUES ( " & vbCr
                        strSQL = strSQL & " '" & sMonth & "' , " & vbCr
                        strSQL = strSQL & " '" & Mid(cboWorkArea.Text, 1, 2) & "' , " & vbCr
                        strSQL = strSQL & " '" & sDiv & "' , " & vbCr
                        strSQL = strSQL & " '" & sTestCd & "' , " & vbCr
                        strSQL = strSQL & " '" & sTestNm & "' , " & vbCr
                        strSQL = strSQL & " '" & sWordTm & "' , " & vbCr
                        strSQL = strSQL & " '" & sEmTm & "' , " & vbCr
                        strSQL = strSQL & " '" & sOutTm & "' , " & vbCr
                        strSQL = strSQL & " '" & sTarget & "' , " & vbCr
                        strSQL = strSQL & " '" & sTCnt & "' , " & vbCr
                        strSQL = strSQL & " '" & sOutCnt & "' , " & vbCr
                        strSQL = strSQL & " '" & sMark & "' , " & vbCr
                        strSQL = strSQL & " '" & sEntId & "' , " & vbCr
                        strSQL = strSQL & " '" & sEntDt & "' , " & vbCr
                        strSQL = strSQL & " '" & sHTm & "' , " & vbCr
                        strSQL = strSQL & " '" & sEmtTm & "' ) " & vbCr
                        
                        DBConn.Execute strSQL
                    Next
                Case vbNo
                    Exit Sub
            End Select
                    
'            Call MsgBox(sMonth & "월의 " & Trim(Mid(cboWorkArea.Text, 3)) & " " & sDivNm & "데이타는 이미 저장되어있습니다.", vbExclamation, App.Title)
'            Exit Sub
        End If
        
        For iCnt = 1 To .MaxRows - 1
            .GetText 1, iCnt, varTmp: sTestCd = varTmp
            .GetText 2, iCnt, varTmp: sTestNm = varTmp
            .GetText 3, iCnt, varTmp: sWordTm = varTmp
            .GetText 4, iCnt, varTmp: sEmTm = varTmp
            .GetText 5, iCnt, varTmp: sOutTm = varTmp
            .GetText 6, iCnt, varTmp: sHTm = varTmp
            .GetText 7, iCnt, varTmp: sTarget = varTmp
            .GetText 8, iCnt, varTmp: sTCnt = varTmp
            .GetText 9, iCnt, varTmp: sOutCnt = varTmp
            .GetText 10, iCnt, varTmp: sMark = varTmp
            
            If InStr(sTestNm, ",") > 0 Then
                sTestNm = Replace(sTestNm, "'", "")
            End If
            
            strSQL = ""
            strSQL = strSQL & "INSERT INTO S2LAB910 (ENT_MONTH , WorkArea, TEST_DIV, TEST_CD, TEST_NM, WARD_TIME, EM_TIME, " & vbCr
            strSQL = strSQL & " OUT_TIME , Target, TOTAL_CNT, OUT_CNT, MARK, ENT_ID, ENT_DT, EMT_TM, H_TIME)" & vbCr
            strSQL = strSQL & " VALUES ( " & vbCr
            strSQL = strSQL & " '" & sMonth & "' , " & vbCr
            strSQL = strSQL & " '" & Mid(cboWorkArea.Text, 1, 2) & "' , " & vbCr
            strSQL = strSQL & " '" & sDiv & "' , " & vbCr
            strSQL = strSQL & " '" & sTestCd & "' , " & vbCr
            strSQL = strSQL & " '" & sTestNm & "' , " & vbCr
            strSQL = strSQL & " '" & sWordTm & "' , " & vbCr
            strSQL = strSQL & " '" & sEmTm & "' , " & vbCr
            strSQL = strSQL & " '" & sOutTm & "' , " & vbCr
            strSQL = strSQL & " '" & sTarget & "' , " & vbCr
            strSQL = strSQL & " '" & sTCnt & "' , " & vbCr
            strSQL = strSQL & " '" & sOutCnt & "' , " & vbCr
            strSQL = strSQL & " '" & sMark & "' , " & vbCr
            strSQL = strSQL & " '" & sEntId & "' , " & vbCr
            strSQL = strSQL & " '" & sEntDt & "' , " & vbCr
            strSQL = strSQL & " '" & sHTm & "' , " & vbCr
            strSQL = strSQL & " '" & sEmtTm & "' ) " & vbCr
            
            DBConn.Execute strSQL
        Next
        Call MsgBox("월별 TAT통계가 저장되었습니다..", vbExclamation, App.Title)
    End With
    Call cmdClear_Click
End Sub

Private Sub cmdStart_Click()

    Dim sStartDate As String, sEndDate As String

    If dtpStart.Value > dtpEnd.Value Then
        MsgBox "Duration input Error"
        Exit Sub
    End If

    sStartDate = Format(dtpStart.Value, CS_DateDbFormat)
    sEndDate = Format(dtpEnd.Value, CS_DateDbFormat)
    
    If cboWorkArea.Text = "" Then
        Call MsgBox("WorkArea를 선택하세요.", vbExclamation, App.Title)
        cboWorkArea.SetFocus
        Exit Sub
    End If
    
    If txtCnt.Text = "" Then
        Call MsgBox("목표달성율을 입력하세요.", vbExclamation, App.Title)
        txtCnt.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    QueryFlag = ReadData  ' True 이면 조회가 이루어 졌음을 의미

    If QueryFlag Then
        dtpStart.Enabled = False
        dtpEnd.Enabled = False
        cmdStart.Enabled = False

        cmdRefresh.Enabled = True

        cmdPrint.Enabled = True
        cmdExcel.Enabled = True
        Call cmdRefresh_Click
    Else
        MsgBox "해당 자료가 없습니다...", vbInformation
    End If
    Me.MousePointer = 0

End Sub



'Private Sub dtpEnd_Validate(Cancel As Boolean)
'    Call clearspdStat
'End Sub
'
'Private Sub dtpStart_Validate(Cancel As Boolean)
'    Call clearspdStat
'End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Dim i    As Integer
    Dim iCnt As Integer
'    Me.Show
    Call ClearRtn

    DoEvents

    Call LoadBuildingList
    Call LoadWorkAreaList
    Call LoadEqpList

    dtpStart.Value = Format(Now, "yyyy-mm-dd")
    dtpEnd.Value = Format(Now, "yyyy-mm-dd")

    ColWid(1) = 16
    ColWid(2) = 18
    ColWid(3) = 18
    ColWid(4) = 38.5
    ColWid(5) = 0

    'fpSpread1.AddCellSpan 병합기준 Col, 병합기준 Row, 병합할 Col수, 병합할 Row수
    
'    For iCnt = 5 To spdStat.MaxCols
'        spdStat.AddCellSpan iCnt, 0, 1, 2
'    Next
'
'    spdStat.AddCellSpan 1, 0, 4, 1
    
    spdStat.SetText 7, 0, "목표달성율% (채혈-결과보고)"
    spdRstList.SetText 5, 0, "채혈일"
    spdRstList.SetText 6, 0, "채혈시간"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QueryFlag = False
    
    Set objSQL = Nothing
    Set objDic = Nothing
    Set rsDeptStat = Nothing
    Set rsTestStat = Nothing
    Set rsDeptTestStat = Nothing

End Sub

Private Sub clearspdStat()
    With spdStat
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .MaxRows = 1
    End With
End Sub

Private Sub clearcfx(Ccfx As ChartFX)
    With Ccfx
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
    End With
End Sub

Public Sub LoadBuildingList()
    Dim tmpRs   As Recordset
    Dim i       As Integer
    Dim SqlStmt As String

    Set tmpRs = New Recordset
    SqlStmt = objSQL.GetBuildCd
    
    tmpRs.Open SqlStmt, DBConn
    
    cboBuilding.Clear
    For i = 1 To tmpRs.RecordCount
        cboBuilding.AddItem Trim("" & tmpRs.Fields("BuildCd").Value) & "   " & _
                            Trim("" & tmpRs.Fields("BuildNm").Value)
        tmpRs.MoveNext
    Next

    tmpRs.Close
    Set tmpRs = Nothing
    Set objSQL = Nothing
    
    If cboBuilding.ListCount > 0 Then cboBuilding.ListIndex = 0 'medComboFind(cboBuilding, objSysInfo.BuildingCd)
    
End Sub

Private Sub LoadWorkAreaList()
    Dim rsGetWA     As Recordset
    Dim sSqlGetWA   As String
    Dim i           As Integer
    
    Set rsGetWA = New Recordset
    rsGetWA.Open objSQL.GetWACd, DBConn

    cboWorkArea.Clear
    For i = 1 To rsGetWA.RecordCount
        cboWorkArea.AddItem "" & rsGetWA.Fields("WACd").Value & " " & _
                            "" & rsGetWA.Fields("WANm").Value
        rsGetWA.MoveNext
    Next i

    rsGetWA.Close
    Set rsGetWA = Nothing

End Sub

Public Sub LoadEqpList()
    Dim rsEQCode        As Recordset
    Dim sSqlGetEQCode   As String
    Dim strBldCd        As String
    Dim i               As Integer
    
    sSqlGetEQCode = objSQL.GetEqpCd
    If cboBuilding.ListIndex > 0 Then
        strBldCd = medGetP(cboBuilding.Text, 1, " ")
        sSqlGetEQCode = objSQL.GetEqpCd(False, strBldCd)
    End If
 
    Set rsEQCode = New Recordset
    rsEQCode.Open sSqlGetEQCode, DBConn

    rsEQCode.Close
    Set rsEQCode = Nothing
End Sub

Private Function ReadData() As Boolean
'    Dim objProBar   As clsProgressBar
    Dim RS          As Recordset
    Dim RS1         As Recordset
    Dim sInOut      As String               '입원 / 외래 구분조회변수
    Dim sOut        As String
    Dim SqlStmt, strSQL    As String
    Dim strTmp      As String               '진료과 초기 변수
    Dim sWorkarea   As String               'workarea
    Dim sWorkareaNm As String               'workarea 명
    Dim BlnFG       As Boolean              '처음의 진료과를 건너뛰기 위한 선언
    Dim StatFg      As String
    Dim i           As Integer
    Dim j           As Long
    Dim kk          As Long                 '마지막 진료과를 담기위한 변수
    Dim iCnt        As Long
    Dim strChkResult As Long
    Dim strWeek1, strWeek2, strWeek3, strWeek4, strWeek5 As String
    Dim strTime     As Long
    Dim tmpResult  As Double
    Dim strResult  As String
    Dim strTmp1 As Long
    Dim ResultCnt As Double
    Dim varTAT    As Variant
    
    Dim lngRow      As Long
    Dim lngTot      As Long
    Dim lngOUT      As Long
    
    Dim strBarno    As String
    Dim intSpd      As Long
    Dim sEM         As String
    
    Dim strEM       As String
    Dim intTAT      As Integer
    Dim booTAT      As Boolean
    
    strChkResult = 0
    strBarno = ""
    
    ReadData = False
    lblMsg = "입력된 기간동안의 검사건수를 집계하고 있습니다..."
   
'    Set objProBar = New clsProgressBar
'    With objProBar
'        .SetMyForm Me
'        .Choice = True
'        .XPos = spdStat.Left + 1700
'        .YPos = spdStat.Top + 30 '
'        .XWidth = (spdStat.Width - 1700)
'        .ForeColor = &H864B24
'        .Appearance = aPlate
'        .YHeight = 280
'        .Msg = "자료를 읽기 위해 준비중입니다..."
'        .Value = 1
'        DoEvents
'    End With
'
    sInOut = ""
    sOut = ""
    sEM = ""
    
    'workarea 별로 검색할시
    objSQL.WorkArea = ""
    objSQL.WorkArea = medGetP(cboWorkArea.Text, 1, " ")
    
    If Option1.Value = True Then
        sInOut = 1 '병동
    Else
        sInOut = 0
    End If
    
    If Option2.Value = True Then
        StatFg = 1
    Else
        StatFg = 2
    End If
    
    If Option4.Value = True Then
        sOut = 1 '외래
    Else
        sOut = 0
    End If
    
    If Option3.Value = True Then
        sEM = 1
    Else
        sEM = 0
    End If
    
    SqlStmt = objSQL.GetAccCnt_Bussdiv_New1(dtpStart.Value, dtpEnd.Value, sInOut, StatFg, sOut, sEM)
    
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If RS.EOF Then
        RS.Close: Set RS = Nothing
'        Set objProBar = Nothing
        Exit Function
    End If
    
'========================================================================================
'  WorkArea별로 선택해서 조회하였을경우의 Flow
'  Workarea 별로 모든 검사항목을 진료과별로 보여준다.
'========================================================================================
'
'    With objProBar
'        .Max = RS.RecordCount * 2
'        .Msg = ""
'    End With
        
    Dim objShow As clsDictionary
    
    Set objShow = New clsDictionary
    

    objShow.Clear
    objShow.FieldInialize "workarea,deptnm,testcd", "Cnt,eqpcd ,testnm,BuildCd"
    
    objShow.Sort = False

    
    sWorkarea = medGetP(cboWorkArea, 1, " ")
    sWorkareaNm = medGetP(cboWorkArea, 2, " ")
    
    strTmp = ""
    BlnFG = False
    kk = 0
    
    ResultCnt = 0
    intSpd = 0
    
    With spdStat
        .MaxRows = 0
'        objProBar.Value = 1
'        objProBar.Max = RS.RecordCount
        RS.MoveFirst
        Do Until RS.EOF
            j = j + 1
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            strBarno = ""
            .Col = 1: .Value = "" & RS.Fields("testcd")
            .Col = 2: .Value = "" & RS.Fields("testnm")
            varTAT = Split("" & RS.Fields("tats"), ";")
            intTAT = UBound(varTAT)
            
            .Col = 3: .Value = varTAT(0)
            strWeek1 = varTAT(0)
'            strWeek1 = "3시간"
            If InStr(strWeek1, "시") > 0 And InStr(strWeek1, "분") > 0 Then
                strWeek1 = Mid(strWeek1, 1, InStr(strWeek1, "시") - 1) & ":" & Mid(strWeek1, InStr(strWeek1, "분") - 1)
            ElseIf InStr(strWeek1, "시") > 0 And InStr(strWeek1, "분") = 0 Then
                strWeek1 = Mid(strWeek1, 1, InStr(strWeek1, "시") - 1) & ":" & "00"
            ElseIf InStr(strWeek1, "시") = 0 And InStr(strWeek1, "분") > 0 Then
                strWeek1 = "00" & ":" & Mid(strWeek1, 1, InStr(strWeek1, "분") - 1)
            Else
                strWeek1 = strWeek1
            End If
            
            .Col = 4: .Value = varTAT(1)
            strWeek2 = varTAT(1)
'            strWeek2 = "40분"
            If InStr(strWeek2, "시") > 0 And InStr(strWeek2, "분") > 0 Then
                strWeek2 = Mid(strWeek2, 1, InStr(strWeek2, "시") - 1) & ":" & Mid(strWeek2, InStr(strWeek2, "분") - 1)
            ElseIf InStr(strWeek2, "시") > 0 And InStr(strWeek2, "분") = 0 Then
                strWeek2 = Mid(strWeek2, 1, InStr(strWeek2, "시") - 1) & ":" & "00"
            ElseIf InStr(strWeek2, "시") = 0 And InStr(strWeek2, "분") > 0 Then
                strWeek2 = "00" & ":" & Mid(strWeek2, 1, InStr(strWeek2, "분") - 1)
            Else
                strWeek2 = strWeek2
            End If
            
            .Col = 5: .Value = varTAT(2)
            strWeek4 = varTAT(2)
'            strWeek2 = "40분"
            If InStr(strWeek4, "시") > 0 And InStr(strWeek4, "분") > 0 Then
                strWeek4 = Mid(strWeek4, 1, InStr(strWeek4, "시") - 1) & ":" & Mid(strWeek4, InStr(strWeek4, "분") - 1)
            ElseIf InStr(strWeek4, "시") > 0 And InStr(strWeek4, "분") = 0 Then
                strWeek4 = Mid(strWeek4, 1, InStr(strWeek4, "시") - 1) & ":" & "00"
            ElseIf InStr(strWeek4, "시") = 0 And InStr(strWeek4, "분") > 0 Then
                strWeek4 = "00" & ":" & Mid(strWeek4, 1, InStr(strWeek4, "분") - 1)
            Else
                strWeek4 = strWeek4
            End If
            
            If intTAT = 3 Then
                .Col = 6: .Value = varTAT(3)
                strWeek5 = varTAT(3)
    '            strWeek2 = "40분"
                If InStr(strWeek5, "시") > 0 And InStr(strWeek5, "분") > 0 Then
                    strWeek5 = Mid(strWeek5, 1, InStr(strWeek5, "시") - 1) & ":" & Mid(strWeek5, InStr(strWeek5, "분") - 1)
                ElseIf InStr(strWeek5, "시") > 0 And InStr(strWeek5, "분") = 0 Then
                    strWeek4 = Mid(strWeek5, 1, InStr(strWeek5, "시") - 1) & ":" & "00"
                ElseIf InStr(strWeek5, "시") = 0 And InStr(strWeek5, "분") > 0 Then
                    strWeek5 = "00" & ":" & Mid(strWeek5, 1, InStr(strWeek5, "분") - 1)
                Else
                    strWeek5 = strWeek5
                End If
            End If
            
            .Col = 7: .Value = txtCnt.Text & " %"
            .Col = 8: .Value = "" & RS.Fields("cnt")
            ResultCnt = ResultCnt + Val("" & RS.Fields("cnt"))
            Me.MousePointer = 11
            DoEvents
'            objProBar.Value = j
            
            strSQL = objSQL.GetAccCnt_ResultTime2(Format(dtpStart.Value, "yyyymmdd"), Format(dtpEnd.Value, "yyyymmdd"), "" & RS.Fields("testcd"), sInOut, StatFg, sOut, sEM)
            
            Set RS1 = New Recordset
            RS1.Open strSQL, DBConn
            
            If RS1.RecordCount > 0 Then
                RS1.MoveFirst
                For i = 1 To RS1.RecordCount
                    If Option5.Value = True Then
                        strWeek3 = Weekday(Format(RS1.Fields("coldt") & "", "####-##-##"))
                        strTime = DateDiff("n", Format(RS1.Fields("coldt") & "" & RS1.Fields("coltm") & "", "####-##-## ##:##:##"), Format(RS1.Fields("vfydt") & "" & RS1.Fields("vfytm") & "", "####-##-## ##:##:##"))
                        If Trim(Val(RS1.Fields("coltm") & "")) >= 90000 And Trim(Val(RS1.Fields("coltm") & "")) <= 160000 Then
                            booTAT = True
                        Else
                            If strWeek5 = "" Then
                                booTAT = True
                            Else
                                booTAT = False
                            End If
                        End If
                    Else
                        strWeek3 = Weekday(Format(RS1.Fields("rcvdt") & "", "####-##-##"))
                        strTime = DateDiff("n", Format(RS1.Fields("rcvdt") & "" & RS1.Fields("rcvtm") & "", "####-##-## ##:##:##"), Format(RS1.Fields("vfydt") & "" & RS1.Fields("vfytm") & "", "####-##-## ##:##:##"))
                        If Trim(Val(RS1.Fields("rcvtm") & "")) >= 90000 And Trim(Val(RS1.Fields("rcvtm") & "")) <= 160000 Then
                            booTAT = True
                        Else
                            If strWeek5 = "" Then
                                booTAT = True
                            Else
                                booTAT = False
                            End If
                        End If
                    End If
                    
'                    If Mid(cboWorkArea.Text, 1, 2) = "01" Then
'                        If Option2.Value = True Then
'                            Select Case strWeek3
'                                Case 1, 7 '일반

'                                If RS1.Fields("ptid") = "00394697" Then Stop
                                If RS1.Fields("statfg") = 1 Then
                                    'strTmp1 = Mid(strWeek2, 1, InStr(strWeek1, ":") - 1) * 60 + Mid(strWeek2, InStr(strWeek2, ":") + 1)
                                    If booTAT = True Then
                                        strTmp1 = Mid(strWeek2, 1, InStr(strWeek2, ":") - 1) * 60 + Mid(strWeek2, InStr(strWeek2, ":") + 1)
                                    Else
                                        strTmp1 = Mid(strWeek5, 1, InStr(strWeek5, ":") - 1) * 60 + Mid(strWeek5, InStr(strWeek5, ":") + 1)
                                    End If
                                    
                                    If strTime > strTmp1 Then
                                        strChkResult = strChkResult + 1
                                        If spdTemp.MaxRows < intSpd + 1 Then
                                            spdTemp.MaxRows = intSpd + 1
                                        End If
                                        With spdTemp
                                            .SetText 1, intSpd + 1, RS1.Fields("ptid")
                                            .SetText 2, intSpd + 1, RS1.Fields("workarea")
                                            .SetText 3, intSpd + 1, RS1.Fields("accdt")
                                            .SetText 4, intSpd + 1, RS1.Fields("accseq")
                                            .SetText 5, intSpd + 1, RS1.Fields("testcd")
                                            If Option5.Value = True Then
                                                .SetText 6, intSpd + 1, RS1.Fields("coldt")
                                                .SetText 7, intSpd + 1, RS1.Fields("coltm")
                                            Else
                                                .SetText 6, intSpd + 1, RS1.Fields("rcvdt")
                                                .SetText 7, intSpd + 1, RS1.Fields("rcvtm")
                                            End If
                                            .SetText 8, intSpd + 1, RS1.Fields("vfydt")
                                            .SetText 9, intSpd + 1, RS1.Fields("vfytm")
                                            .SetText 10, intSpd + 1, "응급"
                                            intSpd = intSpd + 1
                                        End With
                                    End If
                                ElseIf RS1.Fields("wardid") & "" = "" Then
                                    'strTmp1 = Mid(strWeek4, 1, InStr(strWeek4, ":") - 1) * 60 + Mid(strWeek4, InStr(strWeek4, ":") + 1)
                                    If booTAT = True Then
                                        strTmp1 = Mid(strWeek4, 1, InStr(strWeek4, ":") - 1) * 60 + Mid(strWeek4, InStr(strWeek4, ":") + 1)
                                    Else
                                        strTmp1 = Mid(strWeek5, 1, InStr(strWeek5, ":") - 1) * 60 + Mid(strWeek5, InStr(strWeek5, ":") + 1)
                                    End If
                                    
                                    If strTime > strTmp1 Then
                                        strChkResult = strChkResult + 1
                                        If spdTemp.MaxRows < intSpd + 1 Then
                                            spdTemp.MaxRows = intSpd + 1
                                        End If
                                        With spdTemp
                                            .SetText 1, intSpd + 1, RS1.Fields("ptid")
                                            .SetText 2, intSpd + 1, RS1.Fields("workarea")
                                            .SetText 3, intSpd + 1, RS1.Fields("accdt")
                                            .SetText 4, intSpd + 1, RS1.Fields("accseq")
                                            .SetText 5, intSpd + 1, RS1.Fields("testcd")
                                            If Option5.Value = True Then
                                                .SetText 6, intSpd + 1, RS1.Fields("coldt")
                                                .SetText 7, intSpd + 1, RS1.Fields("coltm")
                                            Else
                                                .SetText 6, intSpd + 1, RS1.Fields("rcvdt")
                                                .SetText 7, intSpd + 1, RS1.Fields("rcvtm")
                                            End If
                                            .SetText 8, intSpd + 1, RS1.Fields("vfydt")
                                            .SetText 9, intSpd + 1, RS1.Fields("vfytm")
                                            .SetText 10, intSpd + 1, RS1.Fields("deptcd")
                                            intSpd = intSpd + 1
                                        End With
                                    End If
                                Else
                                    'strTmp1 = Mid(strWeek1, 1, InStr(strWeek1, ":") - 1) * 60 + Mid(strWeek1, InStr(strWeek1, ":") + 1)
                                    If booTAT = True Then
                                        strTmp1 = Mid(strWeek1, 1, InStr(strWeek1, ":") - 1) * 60 + Mid(strWeek1, InStr(strWeek1, ":") + 1)
                                    Else
                                        strTmp1 = Mid(strWeek5, 1, InStr(strWeek5, ":") - 1) * 60 + Mid(strWeek5, InStr(strWeek5, ":") + 1)
                                    End If
                                    If strTime > strTmp1 Then
                                        strChkResult = strChkResult + 1
                                        If spdTemp.MaxRows < intSpd + 1 Then
                                            spdTemp.MaxRows = intSpd + 1
                                        End If
                                        With spdTemp
                                            .SetText 1, intSpd + 1, RS1.Fields("ptid")
                                            .SetText 2, intSpd + 1, RS1.Fields("workarea")
                                            .SetText 3, intSpd + 1, RS1.Fields("accdt")
                                            .SetText 4, intSpd + 1, RS1.Fields("accseq")
                                            .SetText 5, intSpd + 1, RS1.Fields("testcd")
                                            If Option5.Value = True Then
                                                .SetText 6, intSpd + 1, RS1.Fields("coldt")
                                                .SetText 7, intSpd + 1, RS1.Fields("coltm")
                                            Else
                                                .SetText 6, intSpd + 1, RS1.Fields("rcvdt")
                                                .SetText 7, intSpd + 1, RS1.Fields("rcvtm")
                                            End If
                                            .SetText 8, intSpd + 1, RS1.Fields("vfydt")
                                            .SetText 9, intSpd + 1, RS1.Fields("vfytm")
                                            .SetText 10, intSpd + 1, RS1.Fields("deptcd")
                                            intSpd = intSpd + 1
                                        End With
                                    End If
                                End If
                                
'                                Case Else ' 응급
'                                    strTmp1 = Mid(strWeek2, 1, InStr(strWeek2, ":") - 1) * 60 + Mid(strWeek2, InStr(strWeek2, ":") + 1)
'                                    If strTime > strTmp1 Then
'                                        strChkResult = strChkResult + 1
'                                        If spdTemp.MaxRows < intSpd + 1 Then
'                                            spdTemp.MaxRows = intSpd + 1
'                                        End If
'                                        With spdTemp
'                                            .SetText 1, intSpd + 1, RS1.Fields("ptid")
'                                            .SetText 2, intSpd + 1, RS1.Fields("workarea")
'                                            .SetText 3, intSpd + 1, RS1.Fields("accdt")
'                                            .SetText 4, intSpd + 1, RS1.Fields("accseq")
'                                            .SetText 5, intSpd + 1, RS1.Fields("testcd")
'                                            .SetText 6, intSpd + 1, RS1.Fields("rcvdt")
'                                            .SetText 7, intSpd + 1, RS1.Fields("rcvtm")
'                                            .SetText 8, intSpd + 1, RS1.Fields("vfydt")
'                                            .SetText 9, intSpd + 1, RS1.Fields("vfytm")
'                                            intSpd = intSpd + 1
'                                        End With
'                                    End If
'                            End Select
'                        Else
'                            If Val("" & RS1.Fields("rcvtm")) > 83000 And Val("" & RS1.Fields("rcvtm")) < 150000 Then
'                                Select Case strWeek3
'                                    Case 1, 7 '휴일
'                                        strTmp1 = Mid(strWeek2, 1, InStr(strWeek2, ":") - 1) * 60 + Mid(strWeek2, InStr(strWeek2, ":") + 1)
'                                        If strTime > strTmp1 Then
'                                            strChkResult = strChkResult + 1
'                                            If spdTemp.MaxRows < intSpd + 1 Then
'                                                spdTemp.MaxRows = intSpd + 1
'                                            End If
'                                            With spdTemp
'                                                .SetText 1, intSpd + 1, RS1.Fields("ptid")
'                                                .SetText 2, intSpd + 1, RS1.Fields("workarea")
'                                                .SetText 3, intSpd + 1, RS1.Fields("accdt")
'                                                .SetText 4, intSpd + 1, RS1.Fields("accseq")
'                                                .SetText 5, intSpd + 1, RS1.Fields("testcd")
'                                                .SetText 6, intSpd + 1, RS1.Fields("rcvdt")
'                                                .SetText 7, intSpd + 1, RS1.Fields("rcvtm")
'                                                .SetText 8, intSpd + 1, RS1.Fields("vfydt")
'                                                .SetText 9, intSpd + 1, RS1.Fields("vfytm")
'                                                intSpd = intSpd + 1
'                                            End With
'                                        End If
'                                    Case Else ' 평일
'                                        strTmp1 = Mid(strWeek1, 1, InStr(strWeek1, ":") - 1) * 60 + Mid(strWeek1, InStr(strWeek1, ":") + 1)
'                                        If strTime > strTmp1 Then
'                                            strChkResult = strChkResult + 1
'                                            If spdTemp.MaxRows < intSpd + 1 Then
'                                                spdTemp.MaxRows = intSpd + 1
'                                            End If
'                                            With spdTemp
'                                                .SetText 1, intSpd + 1, RS1.Fields("ptid")
'                                                .SetText 2, intSpd + 1, RS1.Fields("workarea")
'                                                .SetText 3, intSpd + 1, RS1.Fields("accdt")
'                                                .SetText 4, intSpd + 1, RS1.Fields("accseq")
'                                                .SetText 5, intSpd + 1, RS1.Fields("testcd")
'                                                .SetText 6, intSpd + 1, RS1.Fields("rcvdt")
'                                                .SetText 7, intSpd + 1, RS1.Fields("rcvtm")
'                                                .SetText 8, intSpd + 1, RS1.Fields("vfydt")
'                                                .SetText 9, intSpd + 1, RS1.Fields("vfytm")
'                                                intSpd = intSpd + 1
'                                            End With
'                                        End If
'                                End Select
'                            End If
'                        End If
'                    Else
'                        Select Case strWeek3
'                            Case 1, 7 '휴일
'                                strTmp1 = Mid(strWeek2, 1, InStr(strWeek2, ":") - 1) * 60 + Mid(strWeek2, InStr(strWeek2, ":") + 1)
'                                If strTime > strTmp1 Then
'                                    strChkResult = strChkResult + 1
'                                    If spdTemp.MaxRows < intSpd + 1 Then
'                                        spdTemp.MaxRows = intSpd + 1
'                                    End If
'                                    With spdTemp
'                                        .GetText 1, intSpd + 1, RS1.Fields("ptid")
'                                        .GetText 2, intSpd + 1, RS1.Fields("workarea")
'                                        .GetText 3, intSpd + 1, RS1.Fields("accdt")
'                                        .GetText 4, intSpd + 1, RS1.Fields("accseq")
'                                        .GetText 5, intSpd + 1, RS1.Fields("testcd")
'                                        .GetText 6, intSpd + 1, RS1.Fields("rcvdt")
'                                        .GetText 7, intSpd + 1, RS1.Fields("rcvtm")
'                                        .GetText 8, intSpd + 1, RS1.Fields("vfydt")
'                                        .GetText 9, intSpd + 1, RS1.Fields("vfytm")
'                                        intSpd = intSpd + 1
'                                    End With
'                                End If
'                            Case Else ' 평일
'                                strTmp1 = Mid(strWeek1, 1, InStr(strWeek1, ":") - 1) * 60 + Mid(strWeek1, InStr(strWeek1, ":") + 1)
'                                If strTime > strTmp1 Then
'                                    strChkResult = strChkResult + 1
'                                    If spdTemp.MaxRows < intSpd + 1 Then
'                                        spdTemp.MaxRows = intSpd + 1
'                                    End If
'                                    With spdTemp
'                                        .GetText 1, intSpd + 1, RS1.Fields("ptid")
'                                        .GetText 2, intSpd + 1, RS1.Fields("workarea")
'                                        .GetText 3, intSpd + 1, RS1.Fields("accdt")
'                                        .GetText 4, intSpd + 1, RS1.Fields("accseq")
'                                        .GetText 5, intSpd + 1, RS1.Fields("testcd")
'                                        .GetText 6, intSpd + 1, RS1.Fields("rcvdt")
'                                        .GetText 7, intSpd + 1, RS1.Fields("rcvtm")
'                                        .GetText 8, intSpd + 1, RS1.Fields("vfydt")
'                                        .GetText 9, intSpd + 1, RS1.Fields("vfytm")
'                                        intSpd = intSpd + 1
'                                    End With
'                                End If
'                        End Select
'                    End If
                    RS1.MoveNext
                Next
            End If
            .Col = 9: .Value = strChkResult
            tmpResult = (strChkResult / Val("" & RS.Fields("cnt"))) * 100
            tmpResult = 100 - Round(tmpResult, 1)
            strResult = tmpResult & " %"
            .Col = 10: .Value = strResult
            strChkResult = 0
            RS.MoveNext
        Loop
        Me.MousePointer = 0
        
        If .MaxRows > 0 Then    '합계처리하자 2014-09-18 PSK
           .MaxRows = .MaxRows + 1
           
           lngTot = 0: lngOUT = 0
           For lngRow = 1 To .MaxRows - 1
                .Row = lngRow: .Col = 8     '총건수
                lngTot = lngTot + CInt(.Value)
                
                .Row = lngRow: .Col = 9     '벗어난건수
                lngOUT = lngOUT + CInt(.Value)
           Next
           
'''           .MaxRows = .MaxRows + 1
           .Row = .MaxRows: .Col = 1: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 2: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 3: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 4: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 5: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 6: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 7: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 8: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 9: .BackColor = RGB(225, 225, 150)
           .Row = .MaxRows: .Col = 10: .BackColor = RGB(225, 225, 150)
           
           .Row = .MaxRows: .Col = 1: .Value = "합계/벗어난비율": .FontBold = True
           .Row = .MaxRows: .Col = 8: .Value = CStr(lngTot): .FontBold = True
           .Row = .MaxRows: .Col = 9: .Value = CStr(lngOUT): .FontBold = True
           .Row = .MaxRows: .Col = 10: .Value = 100 - CStr(Round((lngOUT / lngTot) * 100, 1)) & " % / " & CStr(Round((lngOUT / lngTot) * 100, 1)) & " %": .FontBold = True
        End If
    End With
    
    lblTotalCnt.Caption = ResultCnt & " 건"
    RS.Close
    Set RS = Nothing
    Set objShow = Nothing
'    Set objProBar = Nothing
    ReadData = True
End Function

Private Sub cboSort_Click(Index As Integer)

    Dim i As Integer
    Dim j As Integer

    j = Val(cboSort(Index).Tag)
    If cboSort(Index).ListIndex = 0 Then
        chkSubTot(Index).Value = 0
        Exit Sub
    End If

    cboSort(Index).Tag = cboSort(Index).ListIndex
    SortKeys(cboSort(Index).ListIndex) = Index + 1

    If MsgFg Then Exit Sub
    MsgFg = True
    For i = 0 To cboSort.Count - 1
        If i <> Index Then
            If Val(cboSort(i).Tag) = cboSort(Index).ListIndex Then
                If cboSort(i).ListIndex > 0 Then
                    cboSort(i).ListIndex = j
                Else
                    cboSort(i).Tag = j
                    SortKeys(j) = i + 1
                End If
            End If
        End If
    Next
    MsgFg = False

End Sub


Private Sub ShowData()
''    Dim K(6)    As String
''    Dim FirstFg As Boolean
''    Dim i       As Integer
''
''    FirstFg = True
''    K(1) = "": K(2) = "": K(3) = ""
''    K(4) = "": K(5) = "": K(6) = ""
''    SubTot(1) = 0: SubTot(2) = 0
''    SubTot(3) = 0: SubTot(4) = 0
''    totCnt = 0
''
''    With ssDataBuf
''        .Row = 1: .Row2 = .MaxRows
''        .Col = 1: .COL2 = .MaxCols
''        .BlockMode = True
''        .SortKey(1) = SortKeys(1)
''        .SortKeyOrder(1) = SortKeyOrderAscending
''        .SortKey(2) = SortKeys(2)
''        .SortKeyOrder(2) = SortKeyOrderAscending
''        .SortKey(3) = SortKeys(3)
''        .SortKeyOrder(3) = SortKeyOrderAscending
''        .SortKey(4) = SortKeys(4)
''        .SortKeyOrder(4) = SortKeyOrderAscending
''        .SortKey(5) = SortKeys(5)
''        .SortKeyOrder(5) = SortKeyOrderAscending
''        .SortKey(6) = SortKeys(6)
''        .SortKeyOrder(6) = SortKeyOrderAscending
''        .SortBy = SortByRow
''        .Action = ActionSort
''        .BlockMode = False
''
''        spdStat.MaxRows = 0
''        spdStat.Row = 0
''
''        For i = 0 To cboSort.Count - 1
''            spdStat.Col = Val(cboSort(i).Tag)
''            If cboSort(i).ListIndex = 0 Then
''                spdStat.ColHidden = True
''            Else
''                spdStat.ColHidden = False
''            End If
''            spdStat.ColWidth(spdStat.Col) = ColWid(i + 1)
''        Next
''
''        .Row = 0
''        Call SetValue(1, K(1))
''        Call SetValue(2, K(2))
''        Call SetValue(3, K(3))
''        Call SetValue(4, K(4))
''        Call SetValue(5, K(5))
''
''
''
''        For i = 1 To .MaxRows
''            .Row = i
''
''            If i > 0 Then
''                If chkAll(0).Value = 0 Then
''                    .Col = 5
''                    If .Value <> cboBuilding.Text Then GoTo Skip
''                End If
''                If chkAll(1).Value = 0 Then
''                    .Col = 1
''                    If .Value <> cboWorkArea.Text Then GoTo Skip
''
''                End If
''                If chkAll(2).Value = 0 Then
''                    .Col = 2
''                    If .Value <> cboEqpCd.Text Then GoTo Skip
''                End If
''
''                If chkAll(3).Value = 0 Then
''                    .Col = 3
''                    If .Value <> Trim(lblDeptNm.Caption) Then GoTo Skip
''                End If
''
''                If chkAll(4).Value = 0 Then
''                    .Col = 4
''                    If medGetP(.Value, 1, Space(8 - Len(Trim(txtTestCd.Text)))) <> txtTestCd.Text Then GoTo Skip
''                End If
''
''
''            End If
''
''            .Col = SortKeys(1)
''            If K(1) <> .Value Then
''                If Not FirstFg Then
''                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
''                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
''                    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
''                    If chkSubTot(SortKeys(1)).Value = 1 Then Call SetSubTot(1)
''                End If
''                If cboSort(SortKeys(1) - 1).ListIndex > 0 Then
''                    spdStat.MaxRows = spdStat.MaxRows + 1
''                    spdStat.Row = spdStat.MaxRows
''                    Call SetValue(1, K(1))
''                    Call SetValue(2, K(2))
''                    Call SetValue(3, K(3))
''                    Call SetValue(4, K(4))
''                    Call SetValue(5, K(5))
''                End If
''            End If
''            .Col = SortKeys(2)
''            If K(2) <> .Value Then
''                If Not FirstFg Then
''                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
''                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
''                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
''                    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
''                End If
''                If cboSort(SortKeys(2) - 1).ListIndex > 0 Then
''                    spdStat.MaxRows = spdStat.MaxRows + 1
''                    spdStat.Row = spdStat.MaxRows
''                    Call SetValue(2, K(2))
''                    Call SetValue(3, K(3))
''                    Call SetValue(4, K(4))
''                    Call SetValue(5, K(5))
''                End If
''            End If
''            .Col = SortKeys(3)
''            If K(3) <> .Value Then
''                If Not FirstFg Then
''                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
''                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
''                    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
''                End If
''                If cboSort(SortKeys(3) - 1).ListIndex > 0 Then
''                    spdStat.MaxRows = spdStat.MaxRows + 1
''                    spdStat.Row = spdStat.MaxRows
''                    Call SetValue(3, K(3))
''                    Call SetValue(4, K(4))
''                    Call SetValue(5, K(5))
''                End If
''            End If
''            .Col = SortKeys(4)
''            If K(4) <> .Value Then
''                If Not FirstFg Then
''                    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
''                    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
''                End If
''                If cboSort(.Col - 1).ListIndex > 0 Then
''                    spdStat.MaxRows = spdStat.MaxRows + 1
''                    spdStat.Row = spdStat.MaxRows
''                    Call SetValue(4, K(4))
''                    Call SetValue(5, K(5))
''                End If
''            End If
''
''            .Col = SortKeys(5)
''            If K(5) <> .Value Then
''                'If chkSubTot(.Col - 1).Value = 1 Then Call SetSubTot(5)
''                If cboSort(.Col - 1).ListIndex > 0 Then
''                    spdStat.MaxRows = spdStat.MaxRows + 1
''                    spdStat.Row = spdStat.MaxRows
''                    Call SetValue(5, K(5))
''                End If
''            End If
''
''            .Col = 6: spdStat.Col = 6
''            spdStat.Value = Val(spdStat.Value) + Val(.Value)
''
''
''            SubTot(1) = SubTot(1) + Val(.Value)
''            SubTot(2) = SubTot(2) + Val(.Value)
''            SubTot(3) = SubTot(3) + Val(.Value)
''            SubTot(4) = SubTot(4) + Val(.Value)
''            SubTot(5) = SubTot(5) + Val(.Value)
''            FirstFg = False
''
''            totCnt = totCnt + Val(.Value)
''
''Skip:
''        Next
''
''    End With
''    'If chkSubTot(SortKeys(5) - 1).Value = 1 Then Call SetSubTot(5)
''    If chkSubTot(SortKeys(4)).Value = 1 Then Call SetSubTot(4)
''    If chkSubTot(SortKeys(3)).Value = 1 Then Call SetSubTot(3)
''    If chkSubTot(SortKeys(2)).Value = 1 Then Call SetSubTot(2)
''    If chkSubTot(SortKeys(1)).Value = 1 Then Call SetSubTot(1)
''
''    lblTotalCnt.Caption = Format(totCnt, "###,###,###,###")
''    tabView.Tab = 0
''    spdStat.SetFocus

End Sub

Private Sub SetSubTot(ByVal Col As Integer)
    Dim lngColor As Long
    
    With spdStat
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = Col
        .Value = "소  계"
        lngColor = .BackColor
        .Col = 6
        .Value = SubTot(Col)
        .Col = Col: .COL2 = .MaxCols
        .Row = .Row: .Row2 = .Row
        .BlockMode = True
        .BackColor = &HEEEEEE        'lngColor
        .ForeColor = &HB9602F
        .CellBorderStyle = CellBorderStyleDot
        .CellBorderType = 8  '16
        .Action = ActionSetCellBorder
        '.FontBold = True
        .BlockMode = False
        SubTot(Col) = 0
    End With
End Sub

Private Sub Option5_Click()
    If Option5.Value = True Then
        spdStat.SetText 7, 0, "목표달성율% (채혈-결과보고)"
        spdRstList.SetText 5, 0, "채혈일"
        spdRstList.SetText 6, 0, "채혈시간"
    End If
End Sub

Private Sub Option6_Click()
    If Option6.Value = True Then
        spdStat.SetText 7, 0, "목표달성율% (접수-결과보고)"
        spdRstList.SetText 5, 0, "접수일"
        spdRstList.SetText 6, 0, "접수시간"
    End If
End Sub

Private Sub spdRstList_DblClick(ByVal Col As Long, ByVal Row As Long)
    spdRstList.MaxRows = 0
'    spdRstList.Visible = False
    Frame2.Visible = False
End Sub

Private Sub spdStat_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim strTestCd As String
    Dim varTmp
    Dim strTestCd1 As String
    Dim i, iCnt    As Integer
    Dim strWorkNo, strAccDt, strAccSeq As String
    
    If Col = 9 Then
        spdStat.GetText 1, Row, varTmp
        strTestCd = varTmp
        With spdTemp
            For i = 1 To spdTemp.MaxRows
                .GetText 5, i, varTmp: strTestCd1 = varTmp
                If strTestCd = strTestCd1 Then
                    If spdRstList.MaxRows < iCnt + 1 Then
                        spdRstList.MaxRows = spdRstList.MaxRows + 1
                    End If
                    spdTemp.GetText 1, i, varTmp: spdRstList.SetText 1, iCnt + 1, GetPtNm(varTmp): spdRstList.SetText 2, iCnt + 1, varTmp
                    spdTemp.GetText 2, i, varTmp: strWorkNo = varTmp
                    spdTemp.GetText 3, i, varTmp: strAccDt = varTmp
                    spdTemp.GetText 4, i, varTmp: strAccSeq = varTmp
                    spdRstList.SetText 3, iCnt + 1, strWorkNo & "-" & strAccDt & "-" & strAccSeq
                    spdTemp.GetText 5, i, varTmp: spdRstList.SetText 4, iCnt + 1, varTmp
                    spdTemp.GetText 6, i, varTmp: spdRstList.SetText 5, iCnt + 1, Format(varTmp, "####-##-##")
                    spdTemp.GetText 7, i, varTmp: spdRstList.SetText 6, iCnt + 1, Format(varTmp, "00:00:00")
                    spdTemp.GetText 8, i, varTmp: spdRstList.SetText 7, iCnt + 1, Format(varTmp, "####-##-##")
                    spdTemp.GetText 9, i, varTmp: spdRstList.SetText 8, iCnt + 1, Format(varTmp, "00:00:00")
                    spdTemp.GetText 10, i, varTmp: spdRstList.SetText 9, iCnt + 1, Trim(varTmp)
                    iCnt = iCnt + 1
                End If
            Next
        End With
        Frame2.Visible = True
'        spdRstList.Visible = True
    End If
End Sub

Private Sub spdStat_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Row = 0 Then Exit Sub
    If Col = Val(cboSort(4).Tag) Then
        spdStat.Row = Row
        spdStat.Col = Col
        If spdStat.Value = "소  계" Or Trim(spdStat.Value) = "" Then
            ShowTip = False
            Exit Sub
        End If
        MultiLine = 1
        TipText = "  " & spdStat.Value
        TipWidth = 3000
        spdStat.TextTipDelay = 200
        'Call spdStat.SetTextTipAppearance("굴림", 9, False, False, &HEEFDF2, vbBlue)    '&H996666)
        Call spdStat.SetTextTipAppearance("Arial", 11, False, False, vbWhite, vbBlue)    '&H996666)
        ShowTip = True
    Else
        ShowTip = False
    End If
End Sub

Private Sub ClearRtn()

    Dim i As Integer

    spdStat.MaxRows = 0
    spdTemp.MaxRows = 0
    spdRstList.MaxRows = 0
    spdTemp.Visible = False
'    spdRstList.Visible = False
    Frame2.Visible = False
    lblTotalCnt.Caption = ""

    dtpStart.Enabled = True
    dtpEnd.Enabled = True
    cmdStart.Enabled = True
    cmdRefresh.Enabled = False
    cmdPrint.Enabled = False
    cmdExcel.Enabled = False
       
End Sub

Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
End Sub

Private Sub cmdPrint_Click()
    Dim objSpread    As vaSpread
    Dim strTitle     As String
    Dim strPrintDate As String
    Dim strTestNm    As String
    Dim strPDate     As String
    Dim tmpTitle     As String
    Dim strDate      As String
    Dim strGb        As String
    
    strGb = ""
    strPDate = Format(Now, "yyyy-mm-dd hh:mm:ss")
       
    If Option1.Value = True Then
        strGb = "일반"
    ElseIf Option2.Value = True Then
        strGb = "응급"
    ElseIf Option4.Value = True Then
        strGb = "외래"
    End If
    
    With spdStat
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .FontBold = False
        .FontSize = 9
'        .ColWidth(1) = 10.5
'        .ColWidth(2) = 13.75
'        .ColWidth(3) = 11.75
'        .ColWidth(4) = 11.75
'        .ColWidth(5) = 13.25
'        .ColWidth(6) = 7.75
'        .ColWidth(7) = 7.75
'        .ColWidth(8) = 12.13
'        .ColWidth(9) = 12.13
        .BlockMode = False
               
        .PrintJobName = "검사항목 별 TAT건수통계"

        .PrintAbortMsg = "검사항목 별 TAT건수 통계를 출력중입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1
        
        tmpTitle = "검사항목 별 TAT건수 통계"
'        strTitle = "/fn""굴림체""/fz""18""/fb1/fi0/fu1/fk0/fs1" _
'              & "/f1/c" & tmpTitle & "/n/n/n"
        strTitle = "/fn""굴림체"" /fz""18"" /fb1/fi0/fu0/fk0/fs1" _
                  & "/f1/c" & tmpTitle & "/n/n/n"
        strPrintDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "출력일자 : " & strPDate & "/n/n"
        strTestNm = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "WorkArea : " & cboWorkArea.Text & "/n"
        strDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                  & "/f1/l" & "조회기간 : " & Format(dtpStart.Value, "yyyy-mm-dd") & " ~ " & Format(dtpEnd.Value, "yyyy-mm-dd") & "   조회유형 : " & strGb & "/n"
        .PrintHeader = strTitle & strTestNm & strDate 'strPrintDate
        .PrintMarginLeft = 10
'        .PrintMarginRight = 10
        .PrintOrientation = PrintOrientationPortrait 'PrintOrientationLandscape
'        .PrintOrientation = PrintOrientationLandscape 'PrintOrientationLandscape
        
        
'        P_HOSPITALNAME = "한마음혈액원"
        .PrintFooter = " /l " & String(130, Chr(6)) & "/n/l " & P_HOSPITALNAME & "/c/p/fb1"
     
        .PrintMarginBottom = 100
        .PrintShadows = True
        .PrintMarginTop = 300
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint

'        .GridSolid = True
'        .Row = 1: .Row2 = .MaxRows
'        .Col = 1: .COL2 = .MaxCols
'        .BlockMode = True
'        .FontSize = 9
'        .FontBold = False
''        .ColWidth(1) = 12.63
''        .ColWidth(2) = 16.75
''        .ColWidth(3) = 10.63
''        .ColWidth(4) = 11
''        .ColWidth(5) = 11.13
''        .ColWidth(6) = 14.38
''        .ColWidth(7) = 9.5
''        .ColWidth(8) = 9.13
''        .ColWidth(9) = 14.5
'        .BlockMode = False
    End With
End Sub



