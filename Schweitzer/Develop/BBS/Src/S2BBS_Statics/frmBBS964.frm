VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBBS964 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "진료과별수혈통계"
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optDW 
      BackColor       =   &H00800000&
      Caption         =   "병동"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   1
      Left            =   10095
      TabIndex        =   28
      Top             =   105
      Width           =   705
   End
   Begin VB.OptionButton optDW 
      BackColor       =   &H00800000&
      Caption         =   "진료과"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   0
      Left            =   9180
      TabIndex        =   27
      Top             =   105
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00EAE7E3&
      Caption         =   "Excel(&E)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   22
      Tag             =   "127"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   8400
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   -75
      TabIndex        =   2
      Top             =   2355
      Visible         =   0   'False
      Width           =   675
      _Version        =   196608
      _ExtentX        =   1191
      _ExtentY        =   1191
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
      SpreadDesigner  =   "frmBBS964.frx":0000
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   435
      Left            =   90
      ScaleHeight     =   435
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   1560
      Width           =   10515
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   6705
      Left            =   75
      TabIndex        =   1
      Top             =   1545
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   11827
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   27
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   13818331
      SpreadDesigner  =   "frmBBS964.frx":01CC
      TextTip         =   4
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
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
      Caption         =   "진료과별 출고내역"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1200
      Left            =   75
      TabIndex        =   6
      Top             =   285
      Width           =   10770
      Begin VB.Frame fraDt 
         BackColor       =   &H00DBE6E6&
         Height          =   1050
         Left            =   1245
         TabIndex        =   7
         Top             =   105
         Width           =   3435
         Begin MSComCtl2.DTPicker dtpFromDt 
            Height          =   315
            Left            =   180
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   210
            Width           =   2595
            _ExtentX        =   4577
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
            Format          =   70057984
            CurrentDate     =   36342.5951388889
         End
         Begin MSComCtl2.DTPicker dtpToDt 
            Height          =   315
            Left            =   180
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   660
            Width           =   2595
            _ExtentX        =   4577
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
            Format          =   70057984
            CurrentDate     =   36342.5951388889
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "From"
            Height          =   180
            Left            =   2880
            TabIndex        =   11
            Tag             =   "15104"
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "To"
            Height          =   180
            Left            =   2880
            TabIndex        =   10
            Tag             =   "15104"
            Top             =   765
            Width           =   225
         End
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "폐기"
         Height          =   270
         Index           =   2
         Left            =   7710
         TabIndex        =   25
         Top             =   900
         Width           =   795
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "반환"
         Height          =   270
         Index           =   1
         Left            =   6840
         TabIndex        =   24
         Top             =   900
         Width           =   795
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "출고"
         Height          =   270
         Index           =   0
         Left            =   5970
         TabIndex        =   23
         Top             =   885
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton cmdQuary 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조 회(&Q)"
         Height          =   510
         Left            =   9300
         Style           =   1  '그래픽
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   465
         Width           =   1320
      End
      Begin VB.CommandButton cmdListPop 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   6990
         Style           =   1  '그래픽
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "PtID"
         Top             =   510
         Width           =   360
      End
      Begin VB.CommandButton cmdListPop 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   330
         Index           =   0
         Left            =   6990
         Style           =   1  '그래픽
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "PtID"
         Top             =   150
         Width           =   360
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   7365
         TabIndex        =   14
         Top             =   150
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         BackColor       =   15463405
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
      Begin DRcontrol1.DrText txtWard 
         Height          =   330
         Left            =   5955
         TabIndex        =   15
         Top             =   150
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Text            =   "79203847"
         MaxLength       =   9
         Appearance      =   1
         Alignment       =   2
         BorderColor     =   4210752
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   315
         Left            =   7365
         TabIndex        =   16
         Top             =   510
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         BackColor       =   15463405
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
      Begin DRcontrol1.DrText txtDoct 
         Height          =   315
         Left            =   5955
         TabIndex        =   17
         Top             =   510
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Text            =   "79203847"
         MaxLength       =   9
         Appearance      =   1
         Alignment       =   2
         BorderColor     =   4210752
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   4875
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   150
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
         Caption         =   "진료과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   4875
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   495
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
         Caption         =   "주치의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   165
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   195
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
         Caption         =   "조회기간"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   4875
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   840
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
         Caption         =   "조회구분"
         Appearance      =   0
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   855
      Top             =   735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBBS964"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum tblColumn
    tcOrdDt = 1
    tcDept
    tcDoct
    tcPtid
    tcPtNm
    
    tcBldNo
    tcComp
    tcABO
    tcVOL
    tcDelDt
End Enum

Private WithEvents objListPop   As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1
Private SortTF As Boolean

Private Sub cmdClear_Click()
    Call Form_Clear
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblList.DataRowCnt = 0 And tblList.DataRowCnt = 0 Then Exit Sub
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblList.MaxRows + 1
        .MaxCols = tblList.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = tblList.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = LisLabel11.Caption
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Clear()
    Call medClearTable(tblList)
    dtpFromDt.Value = GetSystemDate
    dtpToDt.Value = GetSystemDate
    txtDoct.Text = "":  txtWard.Text = ""
    lblDoctNm.Caption = "":  lblWardNm.Caption = ""
    tblList.ZOrder 0
End Sub

Private Sub cmdListPop_Click(Index As Integer)
    '리스트 팝업을 불러오자...
    Set objListPop = New clsPopUpList
    With objListPop
        .Connection = DBConn
'        .BackColor = Me.BackColor
        Select Case Index
            '병동 불러오기
            Case 0:
                .FormCaption = IIf(optDW(0).Value, "진료과조회", "병동조회")
                .ColumnHeaderText = "코드;코드명"
'                .Width = .Width + 300
'                .ColSize(0) = 1000
                If optDW(0).Value Then
                    Call .LoadPopUp(GetSQLDeptList) ', 2350, 7650) ', ObjLISComCode.DeptCd)
                ElseIf optDW(1).Value Then
                    Call .LoadPopUp(GetSQLWardList)
                End If
                txtWard.Text = medGetP(.SelectedString, 1, ";")
                lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
            '주치의코드 불러오기
            Case 1:
                .FormCaption = "주치의 조회"
                .ColumnHeaderText = "코드;코드명"
'                .Width = .Width + 700
                Call .LoadPopUp(GetSQLDoctList) ', 2850, 7650)
                txtDoct.Text = medGetP(.SelectedString, 1, ";")
                lblDoctNm.Caption = medGetP(.SelectedString, 2, ";")
        End Select
    End With
    Set objListPop = Nothing
End Sub


Private Sub Form_Load()
    Call Form_Clear
End Sub
Private Sub cmdQuary_Click()
    Dim SSQL    As String
    Dim RS      As Recordset
    Dim ii      As Long
    Dim qStart  As String
    Dim qEnd    As String
    Dim objPro  As New clsProgress
    
    With objPro
        .Container = Me
        .Left = pic.Left
        .Top = pic.Top
        .Width = pic.Width
        .Height = pic.Height
        .Message = "자료를 수집하고 있습니다..."
'        .Choice = True
'        .SetMyForm Me
'        .XPos = pic.Left
'        .YPos = pic.Top
'        .XWidth = pic.Width
'        .YHeight = pic.Height
'        .Appearance = aPlate
'        .Msg = "자료를 수집하고 있습니다."
    End With
    
    Me.MousePointer = 11
    
    qStart = Format(dtpFromDt.Value, "yyyymmdd")
    qEnd = Format(dtpToDt.Value, "yyyymmdd")
    
    Call medClearTable(tblList)
    
    If optDiv(0).Value Then
        SSQL = " SELECT e.ptid,e.orddt,e.deptcd,e.wardid,e.orddoct,b.workarea,b.accdt,b.accseq,b.deliverydt," & _
               " a.rh,a.bldno , a.bldyy, a.bldsrc, a.volumn, a.abo, a.compocd" & _
               " FROM " & T_LAB101 & " e," & T_LAB102 & " d," & T_BBS402 & " b," & T_BBS401 & " a" & _
               " WHERE " & DBW("a.stscd=", "3") & _
               " AND   " & DBW("b.deliverydt>=", qStart) & _
               " AND   " & DBW("b.deliverydt<=", qEnd) & _
               " AND   a.bldsrc=b.bldsrc AND   a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd =b.compocd" & _
               " AND   b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq" & _
               " AND   d.ptid=e.ptid AND d.orddt =e.orddt AND d.ordno=e.ordno" '
    
        If optDW(0).Value Then
            If txtWard.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.deptcd=", Trim(txtWard.Text))
        ElseIf optDW(1).Value Then
            If txtWard.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.wardid=", Trim(txtWard.Text))
        End If
        If txtDoct.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.orddoct=", Trim(txtDoct.Text))
               
        SSQL = SSQL & " ORDER BY e.orddt,e.deptcd,e.orddoct"
        tblList.Row = 0: tblList.Col = tblColumn.tcDelDt: tblList.Value = "출고일"
    ElseIf optDiv(1).Value Then
        SSQL = " SELECT e.ptid,e.orddt,e.deptcd,e.wardid,e.orddoct,b.workarea,b.accdt,b.accseq,b.deliverydt," & _
               " a.rh,a.bldno , a.bldyy, a.bldsrc, a.volumn, a.abo, a.compocd" & _
               " FROM " & T_LAB101 & " e," & T_LAB102 & " d," & T_BBS402 & " b," & T_BBS401 & " a" & _
               " WHERE " & _
                DBW("b.retdt>=", qStart) & _
               " AND   " & DBW("b.retdt<=", qEnd) & _
               " AND   a.bldsrc=b.bldsrc AND   a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd =b.compocd" & _
               " AND   b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq" & _
               " AND   d.ptid=e.ptid AND d.orddt =e.orddt AND d.ordno=e.ordno" '
    
        
        If optDW(0).Value Then
            If txtWard.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.deptcd=", Trim(txtWard.Text))
        ElseIf optDW(1).Value Then
            If txtWard.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.wardid=", Trim(txtWard.Text))
        End If
        If txtDoct.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.orddoct=", Trim(txtDoct.Text))
               
        SSQL = SSQL & " ORDER BY e.orddt,e.deptcd,e.orddoct"
        tblList.Row = 0: tblList.Col = tblColumn.tcDelDt: tblList.Value = "반환일"
    Else
        SSQL = " SELECT e.ptid,e.orddt,e.deptcd,e.wardid,e.orddoct,b.workarea,b.accdt,b.accseq,b.deliverydt," & _
               " a.rh,a.bldno , a.bldyy, a.bldsrc, a.volumn, a.abo, a.compocd" & _
               " FROM " & T_LAB101 & " e," & T_LAB102 & " d," & T_BBS402 & " b," & T_BBS401 & " a" & _
               " WHERE " & DBW("a.stscd=", "4") & _
               " AND   " & DBW("a.realexpdt>=", qStart) & _
               " AND   " & DBW("a.realexpdt<=", qEnd) & _
               " AND   a.bldsrc=b.bldsrc AND   a.bldyy=b.bldyy AND a.bldno=b.bldno AND a.compocd =b.compocd" & _
               " AND   b.workarea=d.workarea AND b.accdt=d.accdt AND b.accseq=d.accseq" & _
               " AND   d.ptid=e.ptid AND d.orddt =e.orddt AND d.ordno=e.ordno" '
    
        
        If optDW(0).Value Then
            If txtWard.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.deptcd=", Trim(txtWard.Text))
        ElseIf optDW(1).Value Then
            If txtWard.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.wardid=", Trim(txtWard.Text))
        End If
        If txtDoct.Text <> "" Then SSQL = SSQL & " AND " & DBW("e.orddoct=", Trim(txtDoct.Text))
               
        SSQL = SSQL & " ORDER BY e.orddt,e.deptcd,e.orddoct"
        tblList.Row = 0: tblList.Col = tblColumn.tcDelDt: tblList.Value = "폐기일"
    End If
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        objPro.Max = RS.RecordCount
        
        With tblList
            .ReDraw = False
            .MaxRows = RS.RecordCount
            Do Until RS.EOF
                ii = ii + 1
                objPro.Value = ii
                .Row = ii
                .Col = tblColumn.tcOrdDt: .Value = Format(RS.Fields("orddt").Value & "", "####-##-##")
                
                .Col = tblColumn.tcDept:
'                            Call ObjComCode.DeptCd.KeyChange(RS.Fields("deptcd").Value & "")
                            If optDW(0).Value Then
                                .Value = GetDeptNm(RS.Fields("deptcd").Value & "") 'ObjComCode.DeptCd.Fields("deptnm")
                            ElseIf optDW(1).Value Then
                                .Value = GetDeptNm(RS.Fields("wardid").Value & "") 'ObjComCode.DeptCd.Fields("deptnm")
                            End If
                
                .Col = tblColumn.tcDoct:
'                            Call ObjComCode.Doct.KeyChange(RS.Fields("orddoct").Value & "")
                            .Value = GetDoctNm(RS.Fields("orddoct").Value & "") 'ObjComCode.Doct.Fields("doctnm")
                .Col = tblColumn.tcPtid: .Value = RS.Fields("ptid").Value & ""
                .Col = tblColumn.tcPtNm: .Value = GetPtNm(RS.Fields("ptid").Value & "")
                .Col = tblColumn.tcComp: .Value = GetCompoNm(RS.Fields("compocd").Value & "")
                .Col = tblColumn.tcBldNo: .Value = RS.Fields("bldsrc").Value & "" & "-" & RS.Fields("bldyy").Value & "" & "-" & Format(RS.Fields("bldno").Value & "", "000000")
                .Col = tblColumn.tcABO: .Value = RS.Fields("abo").Value & "" & RS.Fields("rh").Value & ""
                .Col = tblColumn.tcVOL: .Value = RS.Fields("volumn").Value & "" & "cc"
                .Col = tblColumn.tcDelDt: .Value = Format(RS.Fields("deliverydt").Value & "", "####-##-##")
                RS.MoveNext
            Loop
            .ReDraw = True
            If .MaxRows < 28 Then .MaxRows = 26
        End With
    Else
    
        MsgBox "조건에 맞는자료가 없습니다.", vbInformation + vbOKOnly, "진료과별 수혈내역"
    End If
    Me.MousePointer = 0
    Set RS = Nothing
    Set objPro = Nothing
End Sub
Private Function GetCompoNm(ByVal Compocd As String) As String
    Dim SSQL As String
    Dim RS  As Recordset
    
    SSQL = "SELECT * FROM " & T_BBS006 & " WHERE " & DBW("compocd=", Compocd)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        GetCompoNm = RS.Fields("abbrnm").Value & ""
    End If
    Set RS = Nothing
End Function

Private Function GetPtNm(ByVal PtId As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    
    
    SSQL = "SELECT " & F_PTNM & " as ptnm FROM " & T_HIS001 & " WHERE " & DBW(F_PTID, PtId, 2)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        GetPtNm = RS.Fields("ptnm").Value & ""
    End If
    Set RS = Nothing
End Function
Private Sub SpreadSort(ByVal Col As Integer)
    With tblList
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
        .Col = 1:  .Col2 = .MaxCols
        .Row = 1:  .Row2 = .MaxRows
        .BlockMode = True
        .Action = 25
        .BlockMode = False
        .ReDraw = True
    End With
End Sub

Private Sub optDiv_Click(Index As Integer)
    If optDW(0).Value Then
        Select Case Index
            Case 0: LisLabel11.Caption = "진료과별 출고내역"
            Case 1: LisLabel11.Caption = "진료과별 반환내역"
            Case 2: LisLabel11.Caption = "진료과별 폐기내역"
        End Select
    ElseIf optDW(1).Value Then
        Select Case Index
            Case 0: LisLabel11.Caption = "병동별 출고내역"
            Case 1: LisLabel11.Caption = "병동별 반환내역"
            Case 2: LisLabel11.Caption = "병동별 폐기내역"
        End Select
    End If
End Sub

Private Sub optDW_Click(Index As Integer)
    LisLabel4(0).Caption = IIf(Index = 0, "진료과", "병동")
    
    If txtWard.Text <> "" Then txtWard.Text = "": lblWardNm.Caption = ""
    
    Call tblList.SetText(2, 0, LisLabel4(0).Caption)
    Call optDiv_Click(IIf(optDiv(0).Value, 0, IIf(optDiv(1).Value, 1, 2)))
End Sub

Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then
        Call SpreadSort(Col)
        Exit Sub
    End If
End Sub
