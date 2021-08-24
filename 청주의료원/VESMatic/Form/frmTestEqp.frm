VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmTestEqp 
   BackColor       =   &H80000005&
   Caption         =   " 장비 VS 검사코드 설정"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15315
   Icon            =   "frmTestEqp.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   15315
   WindowState     =   2  '최대화
   Begin HSCotrol.UserPanel pnlEqpitem 
      Height          =   6750
      Left            =   10920
      TabIndex        =   21
      Top             =   1470
      Visible         =   0   'False
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   11906
      Bevel           =   2
      Moveble         =   -1  'True
      CloseEnabled    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spFind3 
         Height          =   6345
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   4035
         _Version        =   393216
         _ExtentX        =   7117
         _ExtentY        =   11192
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
         SpreadDesigner  =   "frmTestEqp.frx":08CA
      End
   End
   Begin Threed.SSPanel sspOpt 
      Height          =   1005
      Left            =   90
      TabIndex        =   13
      Top             =   600
      Width           =   15135
      _Version        =   65536
      _ExtentX        =   26696
      _ExtentY        =   1773
      _StockProps     =   15
      BackColor       =   16244694
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.CommandButton cmdEqpItm_Del 
         Caption         =   "삭 제"
         Height          =   525
         Left            =   14130
         TabIndex        =   24
         Top             =   240
         Width           =   885
      End
      Begin VB.CommandButton cmdEqpItm_Add 
         Caption         =   "추 가"
         Height          =   525
         Left            =   13170
         TabIndex        =   23
         Top             =   240
         Width           =   885
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F7DFD6&
         Caption         =   "검사코드 입력"
         Height          =   885
         Left            =   3750
         TabIndex        =   18
         Top             =   60
         Width           =   7695
         Begin VB.TextBox txtFind2 
            Appearance      =   0  '평면
            Height          =   270
            Left            =   210
            TabIndex        =   2
            Top             =   210
            Width           =   2115
         End
         Begin VB.TextBox txtTestCD 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Height          =   270
            Left            =   210
            MaxLength       =   1000
            TabIndex        =   3
            Top             =   540
            Width           =   7350
         End
         Begin HSCotrol.CButton cmdSerch 
            Height          =   285
            Left            =   2370
            TabIndex        =   19
            Top             =   195
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmTestEqp.frx":0ACF
            MaskColor       =   0
            PicCapAlign     =   1
            BorderStyle     =   1
            BorderColor     =   -2147483632
         End
         Begin VB.Label Label4 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2970
            TabIndex        =   20
            Top             =   270
            Width           =   4455
         End
      End
      Begin VB.TextBox txtVIndex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   270
         Left            =   12570
         MaxLength       =   5
         TabIndex        =   4
         Top             =   390
         Width           =   465
      End
      Begin VB.TextBox txtTstcdEqp 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1410
         MaxLength       =   20
         TabIndex        =   0
         Top             =   225
         Width           =   2040
      End
      Begin VB.TextBox txtTstnmEqp 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1410
         MaxLength       =   20
         TabIndex        =   1
         Top             =   570
         Width           =   2040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00F7DFD6&
         Caption         =   "표시 순서 :"
         Height          =   180
         Left            =   11580
         TabIndex        =   16
         Top             =   450
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F7DFD6&
         Caption         =   "Syscode :"
         Height          =   180
         Left            =   450
         TabIndex        =   15
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00F7DFD6&
         Caption         =   "장비 검사명 :"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   585
         Width           =   1095
      End
   End
   Begin HSCotrol.UserPanel pnlTestitem 
      Height          =   5100
      Left            =   3600
      TabIndex        =   11
      Top             =   1530
      Visible         =   0   'False
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   8996
      Bevel           =   2
      Moveble         =   -1  'True
      CloseEnabled    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spFind2 
         Height          =   4695
         Left            =   90
         TabIndex        =   17
         Top             =   300
         Width           =   9015
         _Version        =   393216
         _ExtentX        =   15901
         _ExtentY        =   8281
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
         SpreadDesigner  =   "frmTestEqp.frx":310D
      End
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   1440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":3312
            Key             =   "TST_E"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":38AC
            Key             =   "TST_M"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   520
      Left            =   90
      TabIndex        =   10
      Top             =   8940
      Width           =   15150
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Print"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   1575
         TabIndex        =   5
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Save"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   3015
         TabIndex        =   6
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   4470
         TabIndex        =   7
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmTestEqp.frx":3E46
      Caption         =   " Instruments Test Item Link ."
      SubCaption      =   "검사실 검사항목과 장비 검사항목을 연결 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Image1 
         Height          =   315
         Left            =   15000
         Top             =   0
         Width           =   345
      End
   End
   Begin FPSpread.vaSpread spFind1 
      Height          =   7215
      Left            =   90
      TabIndex        =   12
      Top             =   1680
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   12726
      _StockProps     =   64
      ButtonDrawMode  =   4
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      ShadowText      =   0
      SpreadDesigner  =   "frmTestEqp.frx":50C8
   End
End
Attribute VB_Name = "frmTestEqp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Const OBJTAG_EQP    As String = "EQP"
Private Const OBJTAG_TST    As String = "TST"
'Private Const AUTO_VEFY     As String = "YES"
'Private Const AUTO_VEFN     As String = "NO"

Private CallForm    As String
Private mAdoRs              As ADODB.Recordset

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
        Case 0: Call cmdPrint
        Case 1: Call cmdSave
        Case 2: Call cmdClear
        Case 3: Call cmdClose
        Case Else
    End Select
    
End Sub

Private Sub cmdPrint() ' 일단 사용정지

'    Call PrintFrom(lvwTestListLab.ListItems)

End Sub

Private Sub frmSetSpread1()

    With spFind1
        .UnitType = UnitTypeTwips   '단위를 트윕으로
        .TypeMaxEditLen = 1000      '텍스트 길이
        .ColHeadersShow = True:    .RowHeadersShow = False
        .GrayAreaBackColor = &H80000005:    .ShadowColor = &HFFC0C0:    .ShadowDark = &HFF8080
        .ScrollBars = ScrollBarsBoth:   .ScrollBarExtMode = True
        .GridShowHoriz = True:  .GridShowVert = True:   .GridColor = &HFFC0C0:  .BackColorStyle = BackColorStyleUnderGrid:  .GridSolid = False
        .OperationMode = OperationModeNormal       'OperationModeSingle : 한행 전체 선택
        .SelectBlockOptions = 0    '블럭선택안함
        .TextTip = TextTipFixed
        .UserResizeRow = UserResizeOff:     .UserResizeCol = UserResizeOn
        .EditEnterAction = EditEnterActionNext
        .EditModeReplace = True
        .Visible = False
        .maxrows = 0:   .MaxCols = 4
        .ColsFrozen = 1
        .CursorStyle = CursorStyleArrow
        
        .Row = 0
        .Font = "굴림":   .FontSize = 9
        .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter
        
        .Col = 1:      .Text = "순서"
        .Col = 2:      .Text = "Syscode"
        .Col = 3:      .Text = "장비검사명"
        .Col = 4:      .Text = "검  사  코  드"
        .RowHeight(0) = 270

        .Row = -1
        .Col = -1:  .Font = "굴림":   .FontSize = 9
        .Col = 1:      .CellType = CellTypeStaticText:   .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeNone:          .ColWidth(.Col) = 800
        .Col = 2:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeNone:        .ColWidth(.Col) = 2000
        .Col = 3:      .CellType = CellTypeStaticText: .TypeVAlign = TypeHAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeNone:    .ColWidth(.Col) = 2250
        .Col = 4:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft
                                .ColMerge = MergeNone:    .ColWidth(.Col) = 9700

        .Visible = True
    End With
End Sub

Private Sub frmSetSpread2()
    With spFind2
        .UnitType = UnitTypeTwips   '단위를 트윕으로
        .ColHeadersShow = True:    .RowHeadersShow = False
        .GrayAreaBackColor = &H80000005:    .ShadowColor = &HFFC0C0:    .ShadowDark = &HFF8080
        .ScrollBars = ScrollBarsBoth:   .ScrollBarExtMode = True
        .GridShowHoriz = True:  .GridShowVert = True:   .GridColor = &HFFC0C0:  .BackColorStyle = BackColorStyleUnderGrid:  .GridSolid = False
        .OperationMode = OperationModeNormal
        .SelectBlockOptions = 0    '블럭선택안함
        .TextTip = TextTipFixed
        .UserResizeRow = UserResizeOff:     .UserResizeCol = UserResizeOn
        .EditEnterAction = EditEnterActionNext
        .EditModeReplace = True
        .Visible = False
        .maxrows = 0:   .MaxCols = 5
        .ColsFrozen = 1
        .CursorStyle = CursorStyleArrow
        
        .Row = 0
        .Font = "굴림":   .FontSize = 9
        .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter
        
        .Col = 0:      .Text = ""
        .Col = 1:      .Text = "CHK"
        .Col = 2:      .Text = "검사코드"
        .Col = 3:      .Text = "서브코드"
        .Col = 4:      .Text = "검사코드명"
        .Col = 5:      .Text = "서브코드명"
        .RowHeight(0) = 270

        .Row = -1
        .Col = -1:  .Font = "굴림":   .FontSize = 9     'MergeAlways/MergeNone/MergeRestricted
        .Col = 0:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeNone:    .ColWidth(.Col) = 500
        .Col = 1:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeNone:    .ColWidth(.Col) = 500
        .Col = 2:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeRestricted:    .ColWidth(.Col) = 1200
        .Col = 3:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeRestricted:    .ColWidth(.Col) = 900
        .Col = 4:      .CellType = CellTypeStaticText:       .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft
                                .ColMerge = MergeRestricted:          .ColWidth(.Col) = 4000
        .Col = 5:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft
                                .ColMerge = MergeRestricted:          .ColWidth(.Col) = 1500
 
        .Visible = True
    End With

End Sub

Private Sub frmSetSpread3()
    With spFind3
        .UnitType = UnitTypeTwips   '단위를 트윕으로
        .ColHeadersShow = True:    .RowHeadersShow = False
        .GrayAreaBackColor = &H80000005:    .ShadowColor = &HFFC0C0:    .ShadowDark = &HFF8080
        .ScrollBars = ScrollBarsBoth:   .ScrollBarExtMode = True
        .GridShowHoriz = True:  .GridShowVert = True:   .GridColor = &HFFC0C0:  .BackColorStyle = BackColorStyleUnderGrid:  .GridSolid = False
        .OperationMode = OperationModeNormal
        .SelectBlockOptions = 0    '블럭선택안함
        .TextTip = TextTipFixed
        .UserResizeRow = UserResizeOff:     .UserResizeCol = UserResizeOn
        .EditEnterAction = EditEnterActionNext
        .EditModeReplace = True
        .Visible = False
        .maxrows = 0:   .MaxCols = 2
        .ColsFrozen = 1
        .CursorStyle = CursorStyleArrow
        
        .Row = 0
        .Font = "굴림":   .FontSize = 9
        .TypeHAlign = TypeHAlignCenter: .TypeVAlign = TypeVAlignCenter
        
        .Col = 0:      .Text = ""
        .Col = 1:      .Text = "LMCID"
        .Col = 2:      .Text = "검사코드"

        .RowHeight(0) = 270

        .Row = -1
        .Col = -1:  .Font = "굴림":   .FontSize = 9     'MergeAlways/MergeNone/MergeRestricted
        .Col = 0:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeNone:    .ColWidth(.Col) = 500
        .Col = 1:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeNone:    .ColWidth(.Col) = 1500
        .Col = 2:      .CellType = CellTypeStaticText: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
                                .ColMerge = MergeRestricted:    .ColWidth(.Col) = 2000
 
        .Visible = True
    End With

End Sub

Private Sub cmdClose()
    Unload Me
End Sub

Private Sub cmdClear()
    Call f_subClear_Form
End Sub

Private Sub cmdEqpItm_Add_Click()

    Dim TstcdEqp     As Boolean     'A:추가, S:수정, D:삭제
    Dim i     As Integer
    Dim Vartemp As Variant
    Dim YesNo   As String
        
    If Trim(txtTstcdEqp) = "" Then
        Call ShowMessage("장비 검사코드가 없습니다. 코드를 입력 하시오.   ")
        txtTstcdEqp.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTstnmEqp) = "" Then
        Call ShowMessage("장비 검사명이 없습니다. 검사명을 입력 하시오.   ")
        txtTstnmEqp.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtTestCD) = "" Then
'        Call ShowMessage("검사코드가 없습니다. 코드를 입력 하세요.   ")
'        Exit Sub
'    End If
    
    With spFind1
        TstcdEqp = False
        For i = 1 To .maxrows
            .Row = i: .Col = 2
            If Trim(.Text) = Trim(txtTstcdEqp.Text) Then
            YesNo = MsgBox(Trim(txtTstcdEqp.Text) & " 장비검사코드 내용을 수정하시겠습니까?", vbExclamation + vbYesNo, INS_NAME)
                If vbYes = YesNo Then
                    TstcdEqp = True
                ElseIf vbNo = YesNo Then
                    Exit Sub
                End If
            End If
        Next i
        
        If txtVIndex.Text = "" Then txtVIndex.Text = 0
        
        If TstcdEqp = True Then
        
            .Col = 0: .Row = .ActiveRow     '추가 후 수정시 추가 안되는것 방지
            If .Text <> "A" Then .SetText 0, .ActiveRow, "S"
            
            .SetText 1, .ActiveRow, Trim(txtVIndex.Text)
            .SetText 2, .ActiveRow, Trim(txtTstcdEqp.Text)
            .SetText 3, .ActiveRow, Trim(txtTstnmEqp.Text)
            .SetText 4, .ActiveRow, Trim(txtTestCD.Text)
        
            Call ShowMessage("수정 되었습니다! 저장을 해주세요.")
        Else
            .maxrows = .maxrows + 1
            .SetText 0, .maxrows, "A"
            .SetText 1, .maxrows, Trim(txtVIndex.Text)
            .SetText 2, .maxrows, Trim(txtTstcdEqp.Text)
            .SetText 3, .maxrows, Trim(txtTstnmEqp.Text)
            .SetText 4, .maxrows, Trim(txtTestCD.Text)
            
            Call ShowMessage("추가 되었습니다! 저장을 해주세요.")
        End If
    End With
    
    txtTstcdEqp = ""
    txtTstnmEqp = ""
    txtTestCD = ""
    txtVIndex = ""
    
    txtTstcdEqp.SetFocus

End Sub

Private Sub cmdEqpItm_Del_Click()

        If Trim(txtTstcdEqp) = "" And Trim(txtTstnmEqp) = "" Then
            Call ShowMessage("선택된 항목이 없습니다. 삭제 하려면 항목을 선택후 삭제하세요.")
            Exit Sub
        End If

        With spFind1
            If vbYes = MsgBox(Trim(txtTstcdEqp.Text) & " 장비검사코드 내용을 삭제하시겠습니까?", vbExclamation + vbYesNo, INS_NAME) Then
                If .maxrows < 1 Then Exit Sub
                .Row = .ActiveRow
                .RowHidden = True
                .SetText 0, .Row, "D"
                
'                .DeleteRows .ActiveRow, 1
'                .MaxRows = .MaxRows - 1
            End If
        End With

        txtTstcdEqp = ""
        txtTstnmEqp = ""
        txtTestCD = ""
        txtVIndex = ""

        txtTstcdEqp.SetFocus
End Sub

Private Sub cmdSave()

    On Error GoTo frmTestEqp_Add_Error

    Dim varTmp As Variant
    Dim sqlDoc  As String
    Dim Ctemp As String
    Dim OUT_SEQ As String, TESTCD_EQP As String, TESTNM_EQP As String, Testcd As String
    Dim i As Integer, y As Integer
    
    Dim adoRS   As New ADODB.Recordset
    Dim Temp_Testcd   As String '수정전 검사코드 값
        
    With spFind1
        For i = 1 To .maxrows
            .GetText 0, i, varTmp:     Ctemp = Trim(varTmp)
            .GetText 1, i, varTmp:     OUT_SEQ = Trim(varTmp)
            .GetText 2, i, varTmp:     TESTCD_EQP = Trim(varTmp)
            .GetText 3, i, varTmp:     TESTNM_EQP = Trim(varTmp)
            .GetText 4, i, varTmp:     Testcd = Trim(varTmp)
            
            If Ctemp = "S" Then '수정
            
                sqlDoc = "SELECT TESTCD FROM INTERFACE002" & _
                         " where EQP_CD = '" & INS_CODE & "' and TESTCD_EQP = '" & Trim(TESTCD_EQP) & "'"
                adoRS.CursorLocation = adUseClient
                adoRS.Open sqlDoc, AdoCn_Jet
                If adoRS.RecordCount > 0 Then adoRS.MoveFirst
                If Not adoRS.EOF Then Temp_Testcd = adoRS(0) & ""
                adoRS.Close:    Set adoRS = Nothing
                
                sqlDoc = "Update INTERFACE002" & _
                         " set TESTNM_EQP = '" & Trim(TESTNM_EQP) & "'," & _
                         "     TESTNM = '" & Trim(TESTNM_EQP) & "'," & _
                         "     OUT_SEQ = '" & Trim(OUT_SEQ) & "'," & _
                         "     TESTCD = '" & Trim(Testcd) & "'" & _
                         " where EQP_CD = '" & INS_CODE & "' and TESTCD_EQP = '" & Trim(TESTCD_EQP) & "'"
                
                AdoCn_Jet.Execute sqlDoc
                
                Call LAB_Machine_Coda_SQL(Ctemp, Trim(Testcd), Trim(Temp_Testcd))
                
            ElseIf Ctemp = "A" Then '추가
                sqlDoc = "Insert into INTERFACE002(" & _
                         "         EQP_CD, TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM)" & _
                         "    values( '" & INS_CODE & "'," & _
                         "            '" & Trim(TESTCD_EQP) & "'," & _
                         "            '" & Trim(TESTNM_EQP) & "'," & _
                         "             " & Val(OUT_SEQ) & "," & _
                         "            '" & Trim(Testcd) & "'," & _
                         "            '" & Trim(TESTNM_EQP) & "')"
                         
                AdoCn_Jet.Execute sqlDoc
                
                Call LAB_Machine_Coda_SQL(Ctemp, Trim(Testcd), "")
                
            ElseIf Ctemp = "D" Then '삭제
                sqlDoc = "delete from INTERFACE002" & _
                         " where EQP_CD = '" & INS_CODE & "' and TESTCD_EQP = '" & Trim(TESTCD_EQP) & "'"
                
                AdoCn_Jet.Execute sqlDoc
                
                Call LAB_Machine_Coda_SQL(Ctemp, Trim(Testcd), "")
                
            End If
        Next i
    End With

    Call f_subSet_EqpData(INS_CODE)

    Exit Sub
frmTestEqp_Add_Error:

    Call ErrMsgProc("frmTestEqp - Private Sub cmdSave()")

End Sub

Private Sub LAB_Machine_Coda_SQL(ByVal Ctemp As String, ByVal Testcd As String, ByVal Temp_Testcd As String)

    Dim tmpTestcd   As Variant, tmpTestCd1   As Variant
    Dim DelTestCd   As Variant, DelTestCd1   As Variant  '수정전 검사코드들 중 삭제된 코드들

    Dim y   As Integer, D   As Integer
    
    If Testcd <> "" Then
        If InStr(Testcd, ",") > 0 Or InStr(Temp_Testcd, ",") > 0 Then
            tmpTestcd = Split(Testcd, ",")
            For y = 0 To UBound(tmpTestcd)
                tmpTestCd1 = Split(tmpTestcd(y), "/")
                If Ctemp = "S" Then
                
                    If InStr(Temp_Testcd, tmpTestCd1(0) & "/" & tmpTestCd1(1)) = 0 Then
                        Call LAB_Machine_Coda_Insert(tmpTestCd1(0), tmpTestCd1(1))
                    Else
                        Temp_Testcd = Trim(Replace(Temp_Testcd, tmpTestCd1(0) & "/" & tmpTestCd1(1), ""))
                    End If
    
                ElseIf Ctemp = "A" Then
                    Call LAB_Machine_Coda_Insert(tmpTestCd1(0), tmpTestCd1(1))
                ElseIf Ctemp = "D" Then
                    Call LAB_Machine_Coda_DELETE(tmpTestCd1(0), tmpTestCd1(1))
                End If
            Next y
            
            '수정후 남은 검사코드(이미 DB에 있는 코드) 삭제
            If Temp_Testcd <> "" And Ctemp = "S" Then
            
                DelTestCd = Split(Temp_Testcd, ",")
                
                For D = 0 To UBound(DelTestCd)
                    If DelTestCd(D) <> "" Then
                        DelTestCd1 = Split(DelTestCd(D), "/")
                        Call LAB_Machine_Coda_DELETE(DelTestCd1(0), DelTestCd1(1))
                    End If
                Next D
            End If
            
            Call f_subSet_Eqpitem(INS_CODE)
        Else
            If Testcd <> "" Then
                tmpTestCd1 = Split(Testcd, "/")
                
                If Ctemp = "S" Then
                    Call LAB_Machine_Coda_Update(tmpTestCd1(0), tmpTestCd1(1), Trim(Temp_Testcd))
                ElseIf Ctemp = "A" Then
                    Call LAB_Machine_Coda_Insert(tmpTestCd1(0), tmpTestCd1(1))
                ElseIf Ctemp = "D" Then
                    Call LAB_Machine_Coda_DELETE(tmpTestCd1(0), tmpTestCd1(1))
                End If
                
                Call f_subSet_Eqpitem(INS_CODE)
            End If
        End If
    Else
        If Temp_Testcd <> "" Then       '검사코드 완전 삭제시
            If InStr(Temp_Testcd, ",") > 0 Then
                tmpTestcd = Split(Temp_Testcd, ",")
                For y = 0 To UBound(tmpTestcd)
                    tmpTestCd1 = Split(tmpTestcd(y), "/")
                    Call LAB_Machine_Coda_DELETE(tmpTestCd1(0), tmpTestCd1(1))
                Next y
            Else
                tmpTestCd1 = Split(Temp_Testcd, "/")
                Call LAB_Machine_Coda_DELETE(tmpTestCd1(0), tmpTestCd1(1))
            End If
        End If
    End If
    
End Sub

Private Sub LAB_Machine_Coda_Update(ByVal tmpCoda As String, ByVal tmpSubCoda As String, ByVal Temp_Testcd As String)

    Dim Old_TestCd  As Variant
    Dim strSql   As String
    
        If Temp_Testcd = "" Then
            Call LAB_Machine_Coda_Insert(tmpCoda, tmpSubCoda)
        Else
            Old_TestCd = Split(Temp_Testcd, "/")
        
            strSql = "Update LAB_Machine_Coda" & _
                     " SET Sys_Code = ''," & _
                     "     Coda = '" & Trim(tmpCoda) & "'," & _
                     "  SubCoda = '" & Trim(tmpSubCoda) & "'" & _
                     " WHERE MachineCode = '" & INS_CODE & "'" & _
                     "          And Coda = '" & Trim(Old_TestCd(0)) & "'" & _
                     "       And SubCoda = '" & Trim(Old_TestCd(1)) & "'"
        
            AdoCn_SQL.Execute strSql
        End If
    
End Sub

Private Sub LAB_Machine_Coda_Insert(ByVal tmpCoda As String, ByVal tmpSubCoda As String)
    
    Dim Z   As Integer
    Dim Coda_chk    As Boolean
    Dim strSql   As String
    
        With spFind3
            Coda_chk = False
             For Z = 1 To .maxrows   'DB에 있는 코드면 Insert 하지 않는다.

                 .Col = 2: .Row = Z
                 If InStr(.Text, Trim(tmpCoda) & "/" & Trim(tmpSubCoda)) > 0 Then
                     Coda_chk = True
                 End If
             Next Z
             
             If Coda_chk = False Then
                 strSql = "Insert LAB_Machine_Coda(MachineCode, Coda, SubCoda, Sys_Code)" & _
                          "VALUES ('" & INS_CODE & "', '" & Trim(tmpCoda) & "', '" & Trim(tmpSubCoda) & "', '')"
    
                 AdoCn_SQL.Execute strSql
    
             End If
    
        End With
End Sub

Private Sub LAB_Machine_Coda_DELETE(ByVal tmpCoda As String, ByVal tmpSubCoda As String)
    
    Dim strSql   As String

        strSql = "DELETE FROM LAB_Machine_Coda" & _
                 " WHERE MachineCode = '" & INS_CODE & "'" & _
                 "          And Coda = '" & Trim(tmpCoda) & "'" & _
                 "       And SubCoda = '" & Trim(tmpSubCoda) & "'"
    
        AdoCn_SQL.Execute strSql
    
End Sub

Private Sub cmdSerch_Click()

    Dim objTestItem As clsCommon
    Dim vntRs       As Variant
    Dim intRow      As Integer, i     As Integer

    With spFind2
        .maxrows = 0
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With

    Set objTestItem = New clsCommon
       
    With objTestItem
        Call .SetAdoCn(AdoCn_SQL)
        vntRs = .Get_TestItem(Trim(txtFind2.Text))
    End With

    Set objTestItem = Nothing

    If IsNull(vntRs) = False Then
        With spFind2
            .Visible = False
            For intRow = 0 To UBound(vntRs, 2)
                .maxrows = UBound(vntRs, 2) + 1

                i = intRow + 1
                .SetText 2, i, Trim(vntRs(0, intRow) & "")  '검사코드
                .SetText 3, i, Trim(vntRs(1, intRow) & "")  '서브코드
                .SetText 4, i, Trim(vntRs(2, intRow) & "")  '검사코드명
                .SetText 5, i, Trim(vntRs(3, intRow) & "")  '서브코드명
                
                If InStr(txtTestCD.Text, Trim(vntRs(0, intRow) & "") & "/" & Trim(vntRs(1, intRow) & "")) > 0 Then
                    .SetText 1, i, "☞"
                End If
             Next intRow
            .Visible = True
        End With

        With pnlTestitem
            Call .Move(cmdSerch.left + 1300, cmdSerch.Top + 1300)
            .Visible = True
            .ZOrder
        End With

    Else

        Call ShowMessage("등록된 검사항목이 없습니다.")

        With pnlTestitem
             .Visible = False
             .ZOrder (0)
        End With
    
    End If
    txtFind2.SetFocus
    
End Sub

Private Sub Form_Load()
    
    CaptionBar1.Caption = INS_NAME & " Instruments Test Item Link ."
    Call cmdClear
    Call frmSetSpread1
    Call frmSetSpread2
    Call frmSetSpread3
    Call f_subSet_EqpData(INS_CODE)
'    Call f_subSet_Eqpitem(INS_CODE)
     
    Label4.BorderStyle = 0
    Label4.Caption = "[검색방법 : B3698, C1234%, %4567, %1321%, C3%]"
   
    With pnlTestitem
        .Moveble = True
        .ZOrder (0)
    End With
    
    With pnlEqpitem
        .Moveble = True
        .ZOrder (0)
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If PopUp_List Is Nothing Then Set PopUp_List = Nothing
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
'    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height + 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 300) / 2, 1300, 360)
    Next
    
End Sub

Private Sub Image1_DblClick()
    Call f_subSet_Eqpitem(INS_CODE)
    
    With pnlEqpitem
        Call .Move(Image1.left - 4000, Image1.Top + 1600)
        .Visible = True
        .ZOrder
    End With
End Sub

'컬럼 누르면 정렬
'Private Sub lvwTestListLab_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Call SetListView_Sort(lvwTestListLab, ColumnHeader)
'End Sub

'Private Sub lvwTstListEqp_DblClick()
'
'    On Error GoTo lvwTstListEqp_DblClick
'
'    If MsgBox("장비검사코드를 삭제하시겠습니까?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
'
'    Dim itemX       As ListItem
'    Dim strTestEqp  As String, intRow   As Integer
'
'    Set itemX = lvwTstListEqp.SelectedItem
'
'    If Not itemX Is Nothing Then
'        AdoCn_Jet.Execute "delete from INTERFACE002 where EQP_CD = '" & INS_CODE & "' and TESTCD_EQP = '" & Trim$(itemX.Text) & "'"
'
'        lblTstcdEqp = "":   lblTstnmEqp = ""
'    End If
'    Set itemX = Nothing
'
'    Call f_subSet_EqpData(INS_CODE)
'
'    Exit Sub
'
'lvwTstListEqp_DblClick:
'    Set itemX = Nothing
'    Call ErrMsgProc("frmTestEqp - Private Sub lvwTstListEqp_DblClick()")
'
'End Sub
'
'Private Sub lvwTstListEqp_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        Call lvwTstListEqp_Click
'        KeyAscii = 0
'        Exit Sub
'    End If
'
'End Sub

Private Sub pnlTestitem_CloseMe()
    pnlTestitem.Visible = False
    txtFind2.Text = ""
    txtFind2.SetFocus
End Sub

Private Sub pnlEqpitem_CloseMe()
    pnlEqpitem.Visible = False
End Sub

Private Sub f_subSet_EqpData(ByVal strEqp_Cd As String)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim i As Integer
    
    On Error GoTo ErrorTrap
 
        With spFind1
            
            .maxrows = 0
            .Col = 1:   .Col2 = .MaxCols
            .Row = 1:   .Row2 = .maxrows
            .BlockMode = True
            .Action = ActionClearText
            .BlockMode = False
            
            sqlDoc = "select TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM" & _
                    " from INTERFACE002" & _
                    " where EQP_CD = '" & strEqp_Cd & "'" & _
                    " order by OUT_SEQ, TESTCD"

            adoRS.CursorLocation = adUseClient
            adoRS.Open sqlDoc, AdoCn_Jet
            
            If adoRS.RecordCount > 0 Then adoRS.MoveFirst

            .Visible = False
            If .maxrows > 0 Then .ClearRange 1, 1, .MaxCols, .maxrows, False
            i = 1

            Do While Not adoRS.EOF
                .maxrows = i
                .SetText 1, i, adoRS!OUT_SEQ & ""
                .SetText 2, i, adoRS!TESTCD_EQP & ""
                .SetText 3, i, adoRS!TESTNM_EQP & ""
                .SetText 4, i, adoRS!Testcd & ""

                adoRS.MoveNext
                i = i + 1
            Loop
            
            adoRS.Close:    Set adoRS = Nothing
            .maxrows = i - 1
            .Visible = True

        End With

Exit Sub
ErrorTrap:
    Set AdoRs_Jet = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub f_subSet_Eqpitem(ByVal strEqp_Cd As String)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim i As Integer
    
    On Error GoTo ErrorTrap
 
        With spFind3
            
            .maxrows = 0
            .Col = 1:   .Col2 = .MaxCols
            .Row = 1:   .Row2 = .maxrows
            .BlockMode = True
            .Action = ActionClearText
            .BlockMode = False
            
            sqlDoc = "SELECT LMCID, Coda, SubCoda" & _
                    " FROM LAB_Machine_Coda" & _
                    " WHERE MachineCode = '" & strEqp_Cd & "'" & _
                    " Order by Coda"

            adoRS.CursorLocation = adUseClient
            adoRS.Open sqlDoc, AdoCn_SQL
            
            If adoRS.RecordCount > 0 Then adoRS.MoveFirst

            .Visible = False
            If .maxrows > 0 Then .ClearRange 1, 1, .MaxCols, .maxrows, False
            i = 1

            Do While Not adoRS.EOF
                .maxrows = i
                .SetText 1, i, adoRS!LMCID & ""
                .SetText 2, i, adoRS!Coda & "/" & adoRS!SubCoda & ""

                adoRS.MoveNext
                i = i + 1
            Loop
            
            adoRS.Close:    Set adoRS = Nothing
            .maxrows = i - 1
            .Visible = True

        End With

Exit Sub
ErrorTrap:
    Set AdoRs_Jet = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub f_subClear_Form()
    txtTstcdEqp = ""
    txtTstnmEqp = ""
    txtFind2 = ""
    txtTestCD = ""
    txtVIndex = ""
End Sub

Private Sub spFind1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim varTmp As Variant
    Dim i As Integer

        With spFind1
    
            For i = 1 To .maxrows
                .Col = -1
                .Row = i
                .BackColor = &H80000005
            Next i

            .Col = -1   '한행 전체 선택
            .Row = .ActiveRow
            .BackColor = &HC0FFC0
 
            .GetText 1, .ActiveRow, varTmp:     txtVIndex.Text = Trim(varTmp)
            .GetText 2, .ActiveRow, varTmp:     txtTstcdEqp.Text = Trim(varTmp)
            .GetText 3, .ActiveRow, varTmp:     txtTstnmEqp.Text = Trim(varTmp)
            .GetText 4, .ActiveRow, varTmp:     txtTestCD.Text = Trim(varTmp)
        End With
    
End Sub

Private Sub spFind2_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    Dim varTmp As Variant
    Dim Codetemp As String
    Dim SubCodetemp As String
    Dim CHK As String

    On Error GoTo ErrorTrap
    
        With spFind2
            
            .GetText 2, .ActiveRow, varTmp:     Codetemp = Trim(varTmp)
            .GetText 3, .ActiveRow, varTmp:     SubCodetemp = Trim(varTmp)

            '검사코드 추가 및 삭제
            CHK = "0"
            If txtTestCD.Text = "" Then
                txtTestCD.Text = Codetemp & "/" & SubCodetemp
                CHK = "1"
            ElseIf InStr(txtTestCD.Text, Codetemp & "/" & SubCodetemp) > 0 And CHK = "0" Then
                txtTestCD.Text = Replace(txtTestCD.Text, Codetemp & "/" & SubCodetemp, "")
                txtTestCD.Text = Replace(txtTestCD.Text, ",,", ",")
                
                If Mid(txtTestCD.Text, 1, 1) = "," Then txtTestCD.Text = Mid(txtTestCD.Text, 2)
                
                If Len(txtTestCD.Text) <> 0 Then
                    If Mid(txtTestCD.Text, Len(txtTestCD.Text), 1) = "," Then _
                        txtTestCD.Text = Mid(txtTestCD.Text, 1, Len(txtTestCD.Text) - 1)
                End If
            Else
                txtTestCD.Text = txtTestCD.Text & "," & Codetemp & "/" & SubCodetemp
                txtTestCD.Text = Replace(txtTestCD.Text, ",,", ",")
                CHK = "1"
            End If
            
            If CHK = "0" Then
                .SetText 1, .ActiveRow, ""
            Else
                .SetText 1, .ActiveRow, "☞"
            End If

        End With
        
Exit Sub
ErrorTrap:
    Set AdoRs_Jet = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub txtFind2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtFind2.Text <> "" Then
            Call cmdSerch_Click
            KeyAscii = 0
            Exit Sub
        Else
            Call ShowMessage("검사코드를 입력해 주세요.")
        End If
    End If
End Sub

'",," , ", " 제거
Private Sub txtTestCD_Change()

    Dim L_tmp   As Integer

    If InStr(txtTestCD.Text, ",,") > 0 Then
        L_tmp = InStr(txtTestCD.Text, ",,")
        txtTestCD.Text = Replace(txtTestCD.Text, ",,", ",")
        txtTestCD.SelStart = L_tmp
    ElseIf InStr(txtTestCD.Text, ", ") > 0 Then
        L_tmp = InStr(txtTestCD.Text, ", ")
        txtTestCD.Text = Replace(txtTestCD.Text, ",,", ",")
        txtTestCD.SelStart = L_tmp
    End If

End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        txtTestCD.Locked = False
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtTstcdEqp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        txtTestCD.Locked = False
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtTstnmEqp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
'        txtTestCD.Locked = False
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtVIndex_GotFocus()
    Call TextBoxs_GotFocus(txtVIndex)
End Sub

Private Sub txtVIndex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0

End Sub

Private Sub txtVIndex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtVIndex.IMEMode = 8
End Sub
