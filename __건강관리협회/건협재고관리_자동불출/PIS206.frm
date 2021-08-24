VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS206 
   BackColor       =   &H00FFFFFF&
   Caption         =   "일마감"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   14880
   WindowState     =   2  '최대화
   Begin XLibrary_XGroupBox.XGroupBox grpMain 
      Height          =   9675
      Left            =   30
      Top             =   30
      Width           =   14790
      _ExtentX        =   26088
      _ExtentY        =   17066
      BackColor       =   16777215
      BorderColor     =   10070188
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TextPosition    =   0
      TextCustomMargin=   4
      GroupBoxStyle   =   0
      TextBarColor1   =   12757903
      TextBarStyle    =   3
      TextBarColor2   =   11767328
      TextBarSymbol   =   0   'False
      TextBarSymbolColor=   16777215
      TextBarHeightMargin=   10
      MouseCursor     =   0
      TextBarMouseCursor=   0
      IconandTextMargin=   4
      BodyColor       =   16777215
      Enabled         =   -1  'True
      Begin BHButton.BHImageButton cmdData 
         Height          =   375
         Left            =   12270
         TabIndex        =   11
         Top             =   330
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "마감자료"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS206.frx":0000
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdProcess 
         Height          =   375
         Left            =   11040
         TabIndex        =   12
         Top             =   330
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "일마감"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS206.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdCancel 
         Height          =   375
         Left            =   9810
         TabIndex        =   15
         Top             =   330
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "마감취소"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS206.frx":2F84
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13500
         TabIndex        =   14
         Top             =   330
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "닫 기"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS206.frx":4746
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   7350
         TabIndex        =   13
         Top             =   330
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "화면지움"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS206.frx":5F08
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   8580
         TabIndex        =   10
         Top             =   330
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "조 회"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS206.frx":76CA
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XLabel.XLabel lblLastDt 
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         Top             =   390
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         BackColor       =   16777215
         Text            =   "최종마감일자 : 2015-05-30"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   255
         IconAndTextMargin=   8
         TextAlign       =   0
         TextAlignMargin =   0
         Focus           =   0   'False
         MouseCursor     =   0
         ToolTipIcon     =   0
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   -1
         ToolTipBackColor=   14811135
         ToolTipForeColor=   0
         ToolTipOpacity  =   100
         ToolTipStyle    =   0
         ToolTipCentered =   0   'False
         ToolTipTitleText=   "Title"
         ToolTipBodyText =   ""
         TextBackColor1  =   -2147483629
         TextBackColor2  =   -2147483629
         TextBackMargin  =   4
         TextBackStyle   =   0
         Enabled         =   -1  'True
      End
      Begin XLibrary_XGroupBox.XGroupBox grpBar 
         Height          =   1395
         Left            =   1800
         Top             =   4230
         Visible         =   0   'False
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   2461
         BackColor       =   12632319
         BorderColor     =   10070188
         BorderRoundNum  =   0
         BorderStyle     =   1
         TextColor       =   0
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         TextPosition    =   0
         TextCustomMargin=   4
         GroupBoxStyle   =   0
         TextBarColor1   =   12757903
         TextBarStyle    =   3
         TextBarColor2   =   11767328
         TextBarSymbol   =   0   'False
         TextBarSymbolColor=   16777215
         TextBarHeightMargin=   10
         MouseCursor     =   0
         TextBarMouseCursor=   0
         IconandTextMargin=   4
         BodyColor       =   16777215
         Enabled         =   -1  'True
         Begin MSComctlLib.ProgressBar prgBar 
            Height          =   405
            Left            =   540
            TabIndex        =   7
            Top             =   750
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   714
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblMsg 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   540
            TabIndex        =   8
            Top             =   300
            Width           =   9885
         End
      End
      Begin XLibrary_XTextBox.XTextBox XTextBox3 
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   5010
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   609
         BackColor       =   16249839
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "☞ 품목별 출고(사용) 내역"
         BorderTextMargin=   4
         PasswordChar    =   ""
         MaxLength       =   0
         MouseCursor     =   4
         TextColor       =   0
         ToolTipOpacity  =   100
         ToolTipIcon     =   2
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   -1
         ToolTipBackColor=   16777215
         ToolTipForeColor=   0
         ToolTipStyle    =   3
         ToolTipCentered =   0   'False
         ToolTipTitleText=   ""
         ToolTipBodyText =   ""
         Locked          =   -1  'True
         Mask            =   0
         PromptChar      =   "_"
         WrongSound      =   0
         CustomSound     =   ""
         MaskShow        =   0   'False
         MaskColor       =   33023
         CustomMask      =   ""
         TextAlign       =   2
         Enabled         =   -1  'True
      End
      Begin XLibrary_XTextBox.XTextBox XTextBox2 
         Height          =   345
         Left            =   8880
         TabIndex        =   5
         Top             =   810
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   609
         BackColor       =   16249839
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "☞ 장비운영내역"
         BorderTextMargin=   4
         PasswordChar    =   ""
         MaxLength       =   0
         MouseCursor     =   4
         TextColor       =   0
         ToolTipOpacity  =   100
         ToolTipIcon     =   2
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   -1
         ToolTipBackColor=   16777215
         ToolTipForeColor=   0
         ToolTipStyle    =   3
         ToolTipCentered =   0   'False
         ToolTipTitleText=   ""
         ToolTipBodyText =   ""
         Locked          =   -1  'True
         Mask            =   0
         PromptChar      =   "_"
         WrongSound      =   0
         CustomSound     =   ""
         MaskShow        =   0   'False
         MaskColor       =   33023
         CustomMask      =   ""
         TextAlign       =   2
         Enabled         =   -1  'True
      End
      Begin XLibrary_XTextBox.XTextBox XTextBox1 
         Height          =   345
         Left            =   90
         TabIndex        =   4
         Top             =   810
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   609
         BackColor       =   16249839
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "☞ 마감검사건수"
         BorderTextMargin=   4
         PasswordChar    =   ""
         MaxLength       =   0
         MouseCursor     =   4
         TextColor       =   0
         ToolTipOpacity  =   100
         ToolTipIcon     =   2
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   -1
         ToolTipBackColor=   16777215
         ToolTipForeColor=   0
         ToolTipStyle    =   3
         ToolTipCentered =   0   'False
         ToolTipTitleText=   ""
         ToolTipBodyText =   ""
         Locked          =   -1  'True
         Mask            =   0
         PromptChar      =   "_"
         WrongSound      =   0
         CustomSound     =   ""
         MaskShow        =   0   'False
         MaskColor       =   33023
         CustomMask      =   ""
         TextAlign       =   2
         Enabled         =   -1  'True
      End
      Begin FPSpreadADO.fpSpread spMach 
         CausesValidation=   0   'False
         Height          =   3765
         Left            =   8880
         TabIndex        =   3
         Tag             =   "20001"
         Top             =   1170
         Width           =   5850
         _Version        =   524288
         _ExtentX        =   10319
         _ExtentY        =   6641
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   3
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS206.frx":8E8C
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
      Begin FPSpreadADO.fpSpread spStk 
         CausesValidation=   0   'False
         Height          =   4245
         Left            =   90
         TabIndex        =   2
         Tag             =   "20001"
         Top             =   5370
         Width           =   14640
         _Version        =   524288
         _ExtentX        =   25823
         _ExtentY        =   7488
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   11
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS206.frx":942E
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox3 
         Height          =   615
         Left            =   90
         Top             =   90
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   1085
         BackColor       =   16311512
         BorderColor     =   10070188
         BorderRoundNum  =   0
         BorderStyle     =   1
         TextColor       =   0
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         TextPosition    =   0
         TextCustomMargin=   4
         GroupBoxStyle   =   0
         TextBarColor1   =   12757903
         TextBarStyle    =   3
         TextBarColor2   =   11767328
         TextBarSymbol   =   0   'False
         TextBarSymbolColor=   16777215
         TextBarHeightMargin=   10
         MouseCursor     =   0
         TextBarMouseCursor=   0
         IconandTextMargin=   4
         BodyColor       =   16777215
         Enabled         =   -1  'True
         Begin TDBDate6Ctl.TDBDate dtpDt 
            Height          =   315
            Left            =   1740
            TabIndex        =   16
            Top             =   150
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "PIS206.frx":9D03
            Caption         =   "PIS206.frx":9DEA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS206.frx":9E4D
            Keys            =   "PIS206.frx":9E6B
            Spin            =   "PIS206.frx":9EC9
            AlignHorizontal =   0
            AlignVertical   =   2
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "yyyy-mm-dd"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   -2147483640
            Format          =   "yyyy-mm-dd"
            HighlightText   =   0
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   5
            MarginRight     =   5
            MarginTop       =   1
            MaxDate         =   2958465
            MinDate         =   -657434
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "2015-07-15"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   42200
            CenturyMode     =   2
         End
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   300
            TabIndex        =   1
            Top             =   150
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "마감일자"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   0
            IconAndTextMargin=   8
            TextAlign       =   0
            TextAlignMargin =   0
            Focus           =   0   'False
            MouseCursor     =   0
            ToolTipIcon     =   0
            ToolTipPopupTime=   -1
            ToolTipHoverTime=   -1
            ToolTipBackColor=   14811135
            ToolTipForeColor=   0
            ToolTipOpacity  =   100
            ToolTipStyle    =   0
            ToolTipCentered =   0   'False
            ToolTipTitleText=   "Title"
            ToolTipBodyText =   ""
            TextBackColor1  =   -2147483629
            TextBackColor2  =   -2147483629
            TextBackMargin  =   4
            TextBackStyle   =   0
            Enabled         =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread spTest 
         CausesValidation=   0   'False
         Height          =   3765
         Left            =   90
         TabIndex        =   0
         Tag             =   "20001"
         Top             =   1170
         Width           =   8610
         _Version        =   524288
         _ExtentX        =   15187
         _ExtentY        =   6641
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   8
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS206.frx":9EF1
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Dim cPis311 As clsPis311, cPis312 As clsPis312, cPis313 As clsPis313, cPis314 As clsPis314
Dim cErp As clsErpLeave, cCtr As clsPisCenter
Dim sDate As String, sReturn As Boolean

    MousePointer = vbHourglass
    If MsgBox("해당일자의 마감을 취소하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        sDate = Format(dtpDt.Value, "yyyyMMdd")
        sReturn = gfMagamCheck(sDate)
        
        If sReturn = False Then
            MsgBox "이후일자에 마감된 자료가 있습니다. 마감취소 하실 수 없습니다.", vbCritical
            MousePointer = vbDefault
            Exit Sub
        End If
        
        Call cDb.csBegin
        
        ' 기존 마감자료 삭제
        Set cPis311 = New clsPis311
        Set cPis312 = New clsPis312
        Set cPis313 = New clsPis313
        Set cPis314 = New clsPis314
        
        Set cErp = New clsErpLeave
        Set cCtr = New clsPisCenter
    
        ' 장비운영일정 자동등록분 삭제
        gSql = "DELETE FROM S2PIS302 WHERE WORKDT='" & sDate & "' AND AUTOFG='1'"
        sReturn = cDb.cfExecute(gSql)
        
        cPis311.workdt = sDate
        sReturn = cPis311.cfDelete
        If sReturn Then
            cPis312.workdt = sDate
            sReturn = cPis312.cfDelete
        End If
        If sReturn Then
            cPis313.workdt = sDate
            sReturn = cPis313.cfDelete
        End If
        
        If sReturn Then
            cPis314.workdt = sDate
            sReturn = cPis314.cfDelete
        End If
        
        If sReturn Then
            cErp.DT_IO = sDate
            sReturn = cErp.cfDeleteAll
        End If
        
        If sReturn Then
            cCtr.areacd = gAreaCd
            cCtr.iodt = sDate
            sReturn = cCtr.cfDeleteAll
        End If
        
        If sReturn Then
            Call cDb.csCommit
            
            Call gsSpreadClear(spTest, 0, True)
            Call gsSpreadClear(spMach, 0, True)
            Call gsSpreadClear(spStk, 0, True)
        
            Call gsButtonEnable(cmdProcess, True)
            Call gsButtonEnable(cmdCancel, False)
            Call gsButtonEnable(cmdData, False)
        Else
            Call cDb.csRollback
        End If
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub cmdClear_Click()
Dim sLastDt As String
    
    sLastDt = gfMagamMaxDate
    lblLastDt.Text = "※ 최종마감일자 : " & Format(sLastDt, "####년 ##월 ##일")

    dtpDt.MinDate = Format(sLastDt, "####-##-##")
    dtpDt.Value = dtpDt.MinDate + 1
    dtpDt.Enabled = True
    
    Call gsSpreadClear(spTest, 0, True)
    Call gsSpreadClear(spMach, 0, True)
    Call gsSpreadClear(spStk, 0, True)
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdCancel, False)
    Call gsButtonEnable(cmdProcess, False)
    Call gsButtonEnable(cmdData, False)

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdData_Click()
Dim cErp As clsErpLeave, cCtr As clsPisCenter
Dim sReturn As Boolean, sUnitRate As Double, sYear As String, sDate As String, sUnitAmt As Currency
Dim sLastNo As Long, sStkCd As String, sStkSumQty As Double

    MousePointer = vbHourglass
    If MsgBox("마감 수불자료를 생성하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        prgBar.Value = 0
        grpBar.Visible = True:      grpBar.ZOrder 0
            
        sDate = Format(dtpDt.Value, "yyyyMMdd")
        sYear = Format(dtpDt.Value, "yyyy")
    
        sReturn = True
        
        gSql = "SELECT COUNT(*) AS RECCNT FROM " & gTBLleave & " WHERE DT_IO='" & sDate & "' AND YN_ERP='Y' GROUP BY DT_IO"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    If Val("" & .Fields("RECCNT").Value) > 0 Then
                        sReturn = False
                    End If
                End If
                .Close
            End If
        End With
    
        If sReturn = False Then
            MsgBox "ERP에서 전표처리된 마감자료가 존재합니다. ERP 전표취소작업을 진행하신 후 다시 생성하시기 바랍니다.!", vbInformation
        End If
    
        Call cDb.csBegin
        
        Set cErp = New clsErpLeave
        Set cCtr = New clsPisCenter
        
        If sReturn Then
            lblMsg.Caption = "... 마감수불자료(ERP 연동자료) 삭제 중 ...":        lblMsg.Refresh
            cErp.DT_IO = sDate
            sReturn = cErp.cfDeleteAll
        End If
        
        If sReturn Then
            lblMsg.Caption = "... 마감수불자료(본부 보고자료) 삭제 중 ...":        lblMsg.Refresh
            cCtr.areacd = gAreaCd
            cCtr.iodt = sDate
            sReturn = cCtr.cfDeleteAll
        End If
        
        If sReturn Then
            lblMsg.Caption = "... 마감수불자료(ERP 연동자료) 생성중 ...":        lblMsg.Refresh
            
            gSql = "SELECT A.ENTDT, A.ENTSEQ, A.OUTQTY, C.*, X.UNIT_SO_FACT AS UNITRATE" & vbNewLine & _
                   "  FROM (" & vbNewLine & _
                   "        SELECT B.ENTDT,B.ENTSEQ, SUM(A.OUTQTY) AS OUTQTY FROM S2PIS314 A" & vbNewLine & _
                   "               INNER JOIN S2PIS401 B ON A.CHULDT=B.CHULDT AND A.CHULSEQ=B.CHULSEQ" & vbNewLine & _
                   "         WHERE A.WORKDT='" & sDate & "' GROUP BY B.ENTDT, B.ENTSEQ " & vbNewLine & _
                   "       ) A INNER JOIN S2PIS201 B ON A.ENTDT=B.ENTDT AND A.ENTSEQ=B.ENTSEQ" & vbNewLine & _
                   "  INNER JOIN " & gTBLenter & " C ON B.CD_COMPANY=C.CD_COMPANY AND B.CD_PLANT=C.CD_PLANT AND B.CD_BIZAREA=C.CD_BIZAREA" & vbNewLine & _
                   "                                    AND B.NO_IO=C.NO_IO AND B.NO_IOLINE=C.NO_IOLINE" & vbNewLine & _
                   "      LEFT JOIN " & gTBLstk & " X ON C.CD_ITEM=X.CD_ITEM" & gERPStkCondition & vbNewLine & _
                   "  ORDER BY C.CD_ITEM"
            With cDb.cfRecordSet(gSql)
                If .State = adStateOpen Then
                    If Not .EOF Then
                        prgBar.Max = .RecordCount
                        While (Not .EOF) And sReturn
                            prgBar.Value = prgBar.Value + 1
                            
                            If sStkCd <> ("" & .Fields("CD_ITEM").Value) Then
                                ' 동일품목에 대해서는 출고항번을 같은번호로 처리
                                sStkCd = ("" & .Fields("CD_ITEM").Value)
                                sLastNo = 0
                                sStkSumQty = 0
                                
                                ' 품목코드기준 출고총수량 집계
                                gSql = "SELECT SUM(OUTQTY) AS SUMQTY FROM S2PIS314 WHERE WORKDT='" & sDate & "' AND STKCD='" & sStkCd & "' GROUP BY STKCD"
                                With cDb.cfRecordSet(gSql)
                                    If .State = adStateOpen Then
                                        If Not .EOF Then
                                            sStkSumQty = Val("" & .Fields("SUMQTY").Value)
                                        End If
                                        .Close
                                    End If
                                End With
                            End If
                            
                            cErp.CD_COMPANY = "" & .Fields("CD_COMPANY").Value
                            cErp.CD_PLANT = "" & .Fields("CD_PLANT").Value
                            cErp.CD_BIZAREA = "" & .Fields("CD_BIZAREA").Value
                            cErp.NO_IO = Format(dtpDt.Value, "yyyyMMdd")
                            cErp.NO_IOLINE = sLastNo
                            cErp.NO_IOLINE2 = 0
                            
                            cErp.CD_SL = "" & .Fields("CD_SL").Value
                            cErp.DT_IO = sDate
                            cErp.CD_ITEM = "" & .Fields("CD_ITEM").Value
                            If Val("" & .Fields("UNITRATE").Value) <> 0 Then
                                cErp.QT_IO = Val("" & .Fields("OUTQTY").Value) / Val("" & .Fields("UNITRATE").Value)
                            Else
                                cErp.QT_IO = Val("" & .Fields("OUTQTY").Value)
                            End If
                            cErp.NO_LOT = "" & .Fields("NO_LOT").Value
                            cErp.DT_LIMIT = "" & .Fields("DT_LIMIT").Value
                            cErp.NO_EMP = gUserId
                            
                            cErp.QT_IO_SUM = sStkSumQty
                            
                            cErp.NO_IO_MGMT = Trim("" & .Fields("NO_IO").Value)
                            cErp.NO_IOLINE_MGMT = Val("" & .Fields("NO_IOLINE").Value)
                            cErp.NO_IOLINE_MGMT2 = 0
'                            cErp.NO_IOLINE_MGMT2 = Val("" & .Fields("NO_IOLINE2").Value)   ' 입고자료에서 no_ioline2 필드 삭제-> 0으로 강제처리
                            
                            sReturn = cErp.cfSave
                            If sReturn Then
                                sLastNo = cErp.NO_IOLINE
                            Else
                                .MoveLast
                            End If
                            
                            .MoveNext
                        Wend
                    End If
                    .Close
                End If
            End With
        End If
        
        If sReturn Then
            lblMsg.Caption = "... 마감수불자료(본부집계자료) 생성중 ...":        lblMsg.Refresh
        
            gSql = "SELECT Y.STKCD, SUM(Y.PREVQTY) AS PREVQTY, SUM(Y.ENTQTY) AS ENTQTY, SUM(Y.OUTQTY) AS OUTQTY, MAX(X.UNIT_SO_FACT) AS UNITRATE                    " & vbNewLine & _
                   "  FROM (                                                                                                                                        " & vbNewLine & _
                   "        SELECT A.STKCD, (SUM(A.PREVQTY)+SUM(A.ENTQTY)-SUM(A.OUTQTY)) AS PREVQTY, 0 AS ENTQTY, 0 AS OUTQTY                                       " & vbNewLine & _
                   "          FROM (                                                                                                                                " & vbNewLine & _
                   "                SELECT STKCD,PREVQTY,0 AS ENTQTY,0 AS OUTQTY FROM S2PIS409 WHERE RMDYEAR='" & sYear & "'                                        " & vbNewLine & _
                   "                UNION All                                                                                                                       " & vbNewLine & _
                   "                SELECT STKCD,0 AS PREVQTY,SUM(IQTY_SO) AS ENTQTY,0 AS OUTQTY FROM S2PIS201                                                      " & vbNewLine & _
                   "                 WHERE SUBSTR(ENTDT,1,4)='" & sYear & "' AND ENTDT<'" & sDate & "' GROUP BY STKCD                                               " & vbNewLine & _
                   "                UNION All                                                                                                                       " & vbNewLine & _
                   "                SELECT STKCD,0 AS PREVQTY, 0 AS ENTQTY,                                                                                         " & vbNewLine & _
                   "                       SUM(NVL(TESTQTY,0)+NVL(FREEQTY,0)+NVL(QCQTY,0)+NVL(RETESTQTY,0)+NVL(MANUQTY,0)+NVL(MACHQTY,0)+NVL(HANDQTY,0)) AS OUTQTY  " & vbNewLine & _
                   "                  FROM S2PIS313 WHERE SUBSTR(WORKDT,1,4)='" & sYear & "' AND WORKDT<'" & sDate & "' GROUP BY STKCD                              " & vbNewLine & _
                   "        ) A GROUP BY A.STKCD                                                                                                                    " & vbNewLine & _
                   "        UNION ALL                                                                                                                               " & vbNewLine & _
                   "        SELECT A.STKCD, 0 AS PREVQTY, SUM(A.ENTQTY) AS ENTQTY, SUM(A.OUTQTY) AS OUTQTY                                                          " & vbNewLine & _
                   "          FROM (                                                                                                                                " & vbNewLine & _
                   "                SELECT STKCD,SUM(IQTY_SO) AS ENTQTY, 0 AS OUTQTY FROM S2PIS201 WHERE ENTDT='" & sDate & "' GROUP BY STKCD                       " & vbNewLine & _
                   "                UNION ALL                                                                                                                       " & vbNewLine & _
                   "                SELECT STKCD,0 AS ENTQTY, SUM(NVL(TESTQTY,0)+NVL(FREEQTY,0)+NVL(QCQTY,0)+NVL(RETESTQTY,0)+NVL(MANUQTY,0)+NVL(MACHQTY,0)         " & vbNewLine & _
                   "                                             +NVL(HANDQTY,0)) AS OUTQTY FROM S2PIS313 WHERE WORKDT='" & sDate & "' GROUP BY STKCD               " & vbNewLine & _
                   "        ) A GROUP BY A.STKCD                                                                                                                    " & vbNewLine & _
                   ") Y LEFT JOIN " & gTBLstk & " X ON Y.STKCD=X.CD_ITEM " & gERPStkCondition & " GROUP BY Y.STKCD"
            With cDb.cfRecordSet(gSql)
                If .State = adStateOpen Then
                    If Not .EOF Then
                        prgBar.Max = .RecordCount
                        prgBar.Value = 0
                        While (Not .EOF) And sReturn
                            prgBar.Value = prgBar.Value + 1
                            
                            sUnitAmt = 0
                            sUnitRate = Val("" & .Fields("UNITRATE").Value)
                            
                            ' 최종입고단가 확인
                            gSql = "SELECT UNITAMT FROM S2PIS201 WHERE STKCD='" & .Fields("STKCD").Value & "' " & vbNewLine & _
                                   "   AND ENTDT<='" & sDate & "' AND ROWNUM<=1 ORDER BY ENTDT DESC"
                            With cDb.cfRecordSet(gSql)
                                If .State = adStateOpen Then
                                    If Not .EOF Then
                                        sUnitAmt = Val("" & .Fields("UNITAMT").Value)
                                    End If
                                End If
                            End With
                            
                            cCtr.areacd = gAreaCd
                            cCtr.iodt = sDate
                            cCtr.stkcd = "" & .Fields("STKCD").Value
                            cCtr.prevqty_o = Val("" & .Fields("PREVQTY").Value)
                            cCtr.inqty_o = Val("" & .Fields("ENTQTY").Value)
                            cCtr.outqty_o = Val("" & .Fields("OUTQTY").Value)
                            If sUnitRate = 0 Then sUnitRate = 1
                            
                            cCtr.prevqty_i = cCtr.prevqty_o / sUnitRate
                            cCtr.inqty_i = cCtr.inqty_o / sUnitRate
                            cCtr.outqty_i = cCtr.outqty_o / sUnitRate
                            
                            cCtr.unitamt = sUnitAmt
                            cCtr.empid = gUserId
                            
                            sReturn = cCtr.cfSave
                            
                            .MoveNext
                        Wend
                    End If
                    .Close
                End If
            End With
        End If
        
        grpBar.Visible = False
        
        If sReturn Then
            Call cDb.csCommit
            MsgBox "마감수불자료의 생성이 완료되었습니다.!", vbInformation
        Else
            Call cDb.csRollback
            MsgBox "마감수불자료 생성중 오류가 발생되어 중단되었습니다.!", vbCritical
        End If
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdFind_Click()
Dim cPis311 As clsPis311, cPis312 As clsPis312, cPis313 As clsPis313
Dim sRow As Long, sSum As Double, sReturn As Boolean

    MousePointer = vbHourglass
    sReturn = False
    
    Set cPis311 = New clsPis311
    With cPis311.cfList(Format(dtpDt.Value, "yyyyMMdd"))
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = sReturn Or True
                Call gsSpreadClear(spTest, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spTest.SetText 1, sRow, "" & .Fields("TESTCD").Value
                    spTest.SetText 2, sRow, "" & .Fields("TESTNM").Value
                    spTest.SetText 3, sRow, Format(Val("" & .Fields("TESTCNT").Value), "#,##0")
                    spTest.SetText 4, sRow, Format(Val("" & .Fields("FREECNT").Value), "#,##0")
                    spTest.SetText 5, sRow, Format(Val("" & .Fields("QCCNT").Value), "#,##0")
                    spTest.SetText 6, sRow, Format(Val("" & .Fields("RETESTCNT").Value), "#,##0")
                    spTest.SetText 7, sRow, Format(Val("" & .Fields("MANUCNT").Value), "#,##0")
                
                    sSum = Val("" & .Fields("TESTCNT").Value) + Val("" & .Fields("FREECNT").Value) _
                           + Val("" & .Fields("QCCNT").Value) + Val("" & .Fields("RETESTCNT").Value) _
                           + Val("" & .Fields("MANUCNT").Value)
                    spTest.SetText 8, sRow, Format(sSum, "#,##0")
                    
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With

    sRow = 0
    Set cPis312 = New clsPis312
    With cPis312.cfList(Format(dtpDt.Value, "yyyyMMdd"))
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = sReturn Or True
                Call gsSpreadClear(spMach, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spMach.SetText 1, sRow, "" & .Fields("EQPNM").Value
                    spMach.SetText 2, sRow, "" & .Fields("OPERNM").Value
                    spMach.SetText 3, sRow, Format(Val("" & .Fields("OPERCNT").Value), "#,##0")
                
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With

    sRow = 0
    Set cPis313 = New clsPis313
    With cPis313.cfList(Format(dtpDt.Value, "yyyyMMdd"))
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = sReturn Or True
                Call gsSpreadClear(spStk, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spStk.SetText 1, sRow, "" & .Fields("STKCD").Value
                    spStk.SetText 2, sRow, "" & .Fields("STKNM").Value
                    spStk.SetText 3, sRow, "" & .Fields("UNIT").Value
                    spStk.SetText 4, sRow, gfQtyOutputStr(Val("" & .Fields("TESTQTY").Value))
                    spStk.SetText 5, sRow, gfQtyOutputStr(Val("" & .Fields("FREEQTY").Value))
                    spStk.SetText 6, sRow, gfQtyOutputStr(Val("" & .Fields("QCQTY").Value))
                    spStk.SetText 7, sRow, gfQtyOutputStr(Val("" & .Fields("RETESTQTY").Value))
                    spStk.SetText 8, sRow, gfQtyOutputStr(Val("" & .Fields("MANUQTY").Value))
                    spStk.SetText 9, sRow, gfQtyOutputStr(Val("" & .Fields("MACHQTY").Value))
                    spStk.SetText 10, sRow, gfQtyOutputStr(Val("" & .Fields("HANDQTY").Value) + Val("" & .Fields("LOTNOQTY").Value))
                
                    sSum = Val("" & .Fields("TESTQTY").Value) + Val("" & .Fields("FREEQTY").Value) + Val("" & .Fields("QCQTY").Value) _
                           + Val("" & .Fields("RETESTQTY").Value) + Val("" & .Fields("MANUQTY").Value) + Val("" & .Fields("MACHQTY").Value) _
                           + Val("" & .Fields("HANDQTY").Value) + Val("" & .Fields("LOTNOQTY").Value)
                    spStk.SetText 11, sRow, gfQtyOutputStr(sSum)
                    
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    
    If sReturn Then
        Call gsButtonEnable(cmdCancel, True)
        Call gsButtonEnable(cmdProcess, False)
        Call gsButtonEnable(cmdData, True)
    Else
        Call gsButtonEnable(cmdCancel, False)
        Call gsButtonEnable(cmdProcess, True)
        Call gsButtonEnable(cmdData, False)
        
        MsgBox "해당일자에 마감된 자료가 없습니다.!", vbInformation
    End If
    
    Call gsButtonEnable(cmdFind, False)
    dtpDt.Enabled = False
    
    MousePointer = vbDefault

End Sub

Private Sub cmdProcess_Click()
Dim cPis311 As clsPis311, cPis312 As clsPis312, cPis313 As clsPis313, cPis314 As clsPis314
Dim cPis302 As clsPis302
Dim sTestCnt As Long, sFreeCnt As Long, sQcCnt As Long, sRetestCnt As Long, sManuCnt As Long, sReturn As Boolean
Dim sDate As String, sGapCnt As Long
Dim sWeek As Integer, sDay As Integer, sUseQty As Double

    Set cPis311 = New clsPis311     ' 검사건수
    Set cPis312 = New clsPis312     ' 장비운영
    Set cPis313 = New clsPis313     ' 품목출고
    Set cPis314 = New clsPis314     ' LOT출고
    
    If MsgBox("해당일자의 마감처리를 진행하시겠습니까 ?", vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If
    
    sDate = Format(dtpDt.Value, "yyyyMMdd")
    sReturn = gfMagamCheck(sDate)
    If sReturn = False Then
        MsgBox "해당일자 이후에 마감된 자료가 있습니다. 마감할 수 없습니다.!", vbCritical
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    Call cDb.csBegin
    
    prgBar.Value = 0
    grpBar.Visible = True:      grpBar.ZOrder 0
    Me.Refresh
    
    If gWorkArea Then
        ' 중앙검사센터 ---------------------------------------------------------------------------
        ' 일반 검사건수 집계
        lblMsg.Caption = "... 일반검사 집계 처리중 ...":        lblMsg.Refresh
        gSql = "-- 진검검사항목 집계                                                                            " & vbNewLine & _
               "SELECT A.TESTCD AS ORDCD, COUNT(*) AS CNT FROM S2LAB302 A                                       " & vbNewLine & _
               "       INNER JOIN S2ORD101 B ON A.PTID=B.PTID AND A.ORDDT=B.ORDDT AND A.ORDNO=B.ORDNO           " & vbNewLine & _
               " WHERE A.VFYDT='" & sDate & "'                                                                  " & vbNewLine & _
               "   AND (B.FREE IS NULL OR B.FREE<>'1') AND (B.TRANSDIV IS NULL OR B.TRANSDIV<>'9')              " & vbNewLine & _
               "   AND EXISTS(SELECT C.TESTCD FROM S2LAB001 C WHERE A.TESTCD=C.TESTCD)                          " & vbNewLine & _
               " GROUP BY A.VFYDT, A.TESTCD                                                                     " & vbNewLine & _
               "UNION ALL                                                                                       " & vbNewLine & _
               "-- 병리검사항목 집계                                                                            " & vbNewLine & _
               "SELECT A.TESTCD AS ORDCD, COUNT(*) AS CNT FROM S2ANA103 A                                       " & vbNewLine & _
               "       INNER JOIN S2ORD101 B ON A.KSAMPLENO=B.PTID AND A.KORDDT=B.ORDDT AND A.KORDNO=B.ORDNO    " & vbNewLine & _
               " WHERE B.ORDDT='" & sDate & "'                                                                  " & vbNewLine & _
               "   AND (B.FREE IS NULL OR B.FREE<>'1') AND (B.TRANSDIV IS NULL OR B.TRANSDIV<>'9')              " & vbNewLine & _
               "   AND EXISTS(SELECT C.TESTCD FROM S2LAB001 C WHERE A.TESTCD=C.TESTCD)                          " & vbNewLine & _
               " GROUP BY A.VFYDT, A.TESTCD                                                                     "
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    prgBar.Max = .RecordCount
                    While (Not .EOF) And sReturn
                        prgBar.Value = prgBar.Value + 1
                        
                        cPis311.workdt = sDate
                        cPis311.testcd = "" & .Fields("ORDCD").Value
                        cPis311.testcnt = Val("" & .Fields("CNT").Value)
                        cPis311.empid = gUserId
                        
                        sReturn = cPis311.cfSave(0)
                            
                        .MoveNext
                    Wend
                End If
                .Close
            End If
        End With
        If sReturn = False Then GoTo errMagam
        
        ' 무료검사건수 집계
        lblMsg.Caption = "... 무료검사 집계 처리중 ...":        lblMsg.Refresh
        gSql = "-- 진검검사항목 집계                                                                            " & vbNewLine & _
               "SELECT A.TESTCD AS ORDCD, COUNT(*) AS CNT FROM S2LAB302 A                                       " & vbNewLine & _
               "       INNER JOIN S2ORD101 B ON A.PTID=B.PTID AND A.ORDDT=B.ORDDT AND A.ORDNO=B.ORDNO           " & vbNewLine & _
               " WHERE A.VFYDT='" & sDate & "'                                                                  " & vbNewLine & _
               "   AND B.FREE='1' AND (B.TRANSDIV IS NULL OR B.TRANSDIV<>'9')                                   " & vbNewLine & _
               "   AND EXISTS(SELECT C.TESTCD FROM S2LAB001 C WHERE A.TESTCD=C.TESTCD)                          " & vbNewLine & _
               " GROUP BY A.VFYDT, A.TESTCD                                                                     " & vbNewLine & _
               "UNION ALL                                                                                       " & vbNewLine & _
               "-- 병리검사항목 집계                                                                            " & vbNewLine & _
               "SELECT A.TESTCD AS ORDCD, COUNT(*) AS CNT FROM S2ANA103 A                                       " & vbNewLine & _
               "       INNER JOIN S2ORD101 B ON A.KSAMPLENO=B.PTID AND A.KORDDT=B.ORDDT AND A.KORDNO=B.ORDNO    " & vbNewLine & _
               " WHERE A.VFYDT='" & sDate & "'                                                                  " & vbNewLine & _
               "   AND B.FREE='1' AND (B.TRANSDIV IS NULL OR B.TRANSDIV<>'9')                                   " & vbNewLine & _
               "   AND EXISTS(SELECT C.TESTCD FROM S2LAB001 C WHERE A.TESTCD=C.TESTCD)                          " & vbNewLine & _
               " GROUP BY A.VFYDT, A.TESTCD                                                                     "
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    prgBar.Value = 0:       prgBar.Max = .RecordCount
                    While (Not .EOF) And sReturn
                        prgBar.Value = prgBar.Value + 1
                        
                        cPis311.workdt = sDate
                        cPis311.testcd = "" & .Fields("ORDCD").Value
                        cPis311.freecnt = Val("" & .Fields("CNT").Value)
                        cPis311.empid = gUserId
                        
                        sReturn = cPis311.cfSave(1)
                        
                        .MoveNext
                    Wend
                End If
                .Close
            End If
        End With
        If sReturn = False Then GoTo errMagam
        
        ' 재검검사 집계
        lblMsg.Caption = "... 재검검사(장비인터페이스) 집계 처리중 ...":        lblMsg.Refresh
        gSql = "-- 진검재검사항목 집계                                                                          " & vbNewLine & _
               "SELECT A.TESTCD AS ORDCD, COUNT(*) AS CNT FROM S2LAB302 A                                       " & vbNewLine & _
               "       INNER JOIN S2ORD101 B ON A.PTID=B.PTID AND A.ORDDT=B.ORDDT AND A.ORDNO=B.ORDNO           " & vbNewLine & _
               " WHERE A.RETESTDT='" & sDate & "'                                                               " & vbNewLine & _
               "   AND (B.FREE IS NULL OR B.FREE<>'1') AND (B.TRANSDIV IS NULL OR B.TRANSDIV<>'9')              " & vbNewLine & _
               "   AND EXISTS(SELECT C.TESTCD FROM S2LAB001 C WHERE A.TESTCD=C.TESTCD)                          " & vbNewLine & _
               " GROUP BY A.RETESTDT, A.TESTCD                                                                  "
'        gSql = "SELECT A.TESTCD AS ORDCD, COUNT(*) AS CNT FROM S2IFSRST A               " & vbNewLine & _
'               " WHERE A.INST_DT='" & sDate & "'                                        " & vbNewLine & _
'               " GROUP BY A.INST_DT, A.TESTCD HAVING COUNT(*)>1"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    prgBar.Value = 0:       prgBar.Max = .RecordCount
                    While (Not .EOF) And sReturn
                        prgBar.Value = prgBar.Value + 1
                        
                        cPis311.workdt = sDate
                        cPis311.testcd = "" & .Fields("ORDCD").Value
                        cPis311.retestcnt = Val("" & .Fields("CNT").Value)
                        cPis311.empid = gUserId
                        
                        sReturn = cPis311.cfSave(3)
                        
                        .MoveNext
                    Wend
                End If
                .Close
            End If
        End With
        If sReturn = False Then GoTo errMagam
    Else
        ' 건협지부 ---------------------------------------------------------------------------
        ' 일반 검사건수 집계
        lblMsg.Caption = "... 일반검사 집계 처리중 ...":        lblMsg.Refresh
        gSql = "SELECT A.ITEMCODE AS ORDCD, COUNT(*) AS CNT FROM " & gKahpUser & "TWMED_RESULT1 A                                   " & vbNewLine & _
               "       INNER JOIN " & gKahpUser & "TWMED_KEYTBL B ON A.JDATE=B.JDATE AND A.CORPCODE=B.CORPCODE AND A.SEQNO=B.SEQNO  " & vbNewLine & _
               "       INNER JOIN " & gKahpUser & "TWMED_ITEM C ON A.ITEMCODE=C.ITEMCODE                                            " & vbNewLine & _
               " WHERE TO_CHAR(A.REGILSI,'YYYYMMDD')='" & sDate & "'                                                                " & vbNewLine & _
               "   AND NOT(C.LPARTCODE IS NULL) AND (B.LAB_FREE IS NULL OR B.LAB_FREE<>'1')                                         " & vbNewLine & _
               " GROUP BY A.ITEMCODE"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    prgBar.Max = .RecordCount
                    While (Not .EOF) And sReturn
                        prgBar.Value = prgBar.Value + 1
                        
                        ' 동일인이 동일검사로 2개이상 결과등록될 경우 중복횟수 제외
                        sGapCnt = 0
                        gSql = "SELECT A.JUMIN, COUNT(*) AS GAPCNT FROM " & gKahpUser & "TWMED_RESULT1 A                            " & vbNewLine & _
                               "       INNER JOIN " & gKahpUser & "TWMED_KEYTBL B ON A.JDATE=B.JDATE                                " & vbNewLine & _
                               "                  AND A.CORPCODE=B.CORPCODE AND A.SEQNO=B.SEQNO                                     " & vbNewLine & _
                               " WHERE TO_CHAR(A.REGILSI,'YYYYMMDD')='" & sDate & "' AND A.ITEMCODE='" & .Fields("ORDCD").Value & "'" & vbNewLine & _
                               "   AND (B.LAB_FREE IS NULL OR B.LAB_FREE<>'1') GROUP BY A.JUMIN HAVING COUNT(*)>1"
                        With cDb.cfRecordSet(gSql)
                            If .State = adStateOpen Then
                                While (Not .EOF) And sReturn
                                    sGapCnt = sGapCnt + (Val("" & .Fields("GAPCNT").Value)) - 1
                                    
                                    .MoveNext
                                Wend
                                .Close
                            End If
                        End With
                        
                        cPis311.workdt = sDate
                        cPis311.testcd = "" & .Fields("ORDCD").Value
                        cPis311.testcnt = Val("" & .Fields("CNT").Value) - sGapCnt
                        cPis311.empid = gUserId
                        
                        sReturn = cPis311.cfSave(0)
                            
                        .MoveNext
                    Wend
                End If
                .Close
            End If
        End With
        If sReturn = False Then GoTo errMagam
        
        ' 무료검사건수 집계
        lblMsg.Caption = "... 무료검사 집계 처리중 ...":        lblMsg.Refresh
        gSql = "SELECT A.ITEMCODE AS ORDCD, COUNT(*) AS CNT FROM " & gKahpUser & "TWMED_RESULT1 A                                   " & vbNewLine & _
               "       INNER JOIN " & gKahpUser & "TWMED_KEYTBL B ON A.JDATE=B.JDATE AND A.CORPCODE=B.CORPCODE AND A.SEQNO=B.SEQNO  " & vbNewLine & _
               "       INNER JOIN " & gKahpUser & "TWMED_ITEM C ON A.ITEMCODE=C.ITEMCODE                                            " & vbNewLine & _
               " WHERE TO_CHAR(A.REGILSI,'YYYYMMDD')='" & sDate & "'                                                                " & vbNewLine & _
               "   AND NOT(C.LPARTCODE IS NULL) AND B.LAB_FREE='1'                                                                  " & vbNewLine & _
               "  GROUP BY A.ITEMCODE"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    prgBar.Value = 0:       prgBar.Max = .RecordCount
                    While (Not .EOF) And sReturn
                        prgBar.Value = prgBar.Value + 1
                        
                        ' 동일인이 동일검사로 2개이상 결과등록될 경우 중복횟수 제외
                        sGapCnt = 0
                        gSql = "SELECT A.JUMIN, COUNT(*) AS GAPCNT FROM TWMED_RESULT1 A                                             " & vbNewLine & _
                               "       INNER JOIN TWMED_KEYTBL B ON A.JDATE=B.JDATE AND A.CORPCODE=B.CORPCODE AND A.SEQNO=B.SEQNO   " & vbNewLine & _
                               " WHERE TO_CHAR(A.REGILSI,'YYYYMMDD')='" & sDate & "' AND A.ITEMCODE='" & .Fields("ORDCD").Value & "'" & vbNewLine & _
                               "   AND B.LAB_FREE='1' GROUP BY A.JUMIN HAVING COUNT(*)>1"
                        With cDb.cfRecordSet(gSql)
                            If .State = adStateOpen Then
                                While (Not .EOF) And sReturn
                                    sGapCnt = sGapCnt + (Val("" & .Fields("GAPCNT").Value)) - 1
                                    
                                    .MoveNext
                                Wend
                                .Close
                            End If
                        End With
                        
                        cPis311.workdt = sDate
                        cPis311.testcd = "" & .Fields("ORDCD").Value
                        cPis311.freecnt = Val("" & .Fields("CNT").Value) - sGapCnt
                        cPis311.empid = gUserId
                        
                        sReturn = cPis311.cfSave(1)
                        
                        .MoveNext
                    Wend
                End If
                .Close
            End If
        End With
        If sReturn = False Then GoTo errMagam
        ' 재검검사 집계(건협지부는 재검에 대한 기록이 없으므로 집계에서 제외)
    End If
    
    ' 정도관리 집계
    lblMsg.Caption = "... 정도관리 검사집계 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT ITEM_CD, COUNT(*) AS CNT FROM S2QCS101 WHERE RST_INP_DT='" & sDate & "'" & vbNewLine & _
           "   AND USE_YN='1' GROUP BY ITEM_CD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                prgBar.Value = 0:       prgBar.Max = .RecordCount
                While (Not .EOF) And sReturn
                    prgBar.Value = prgBar.Value + 1
                    
                    cPis311.workdt = sDate
                    cPis311.testcd = "" & .Fields("ITEM_CD").Value
                    cPis311.qccnt = Val("" & .Fields("CNT").Value)
                    cPis311.empid = gUserId
                    
                    sReturn = cPis311.cfSave(2)
                        
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam
    
    ' 수기검사 집계
    lblMsg.Caption = "... 수기검사 집계 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT TESTCD, SUM(TESTCNT) AS CNT FROM S2PIS301 WHERE WORKDT='" & sDate & "'" & vbNewLine & _
           " GROUP BY WORKDT,TESTCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                prgBar.Value = 0:       prgBar.Max = .RecordCount
                While (Not .EOF) And sReturn
                    prgBar.Value = prgBar.Value + 1
                    
                    cPis311.workdt = sDate
                    cPis311.testcd = "" & .Fields("TESTCD").Value
                    cPis311.manucnt = Val("" & .Fields("CNT").Value)
                    cPis311.empid = gUserId
                    
                    sReturn = cPis311.cfSave(4)
                        
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam
    
    ' 검사별 소요품목
    lblMsg.Caption = "... 검사항목별 소요품목 출고 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT B.STKCD, SUM(A.TESTCNT*(B.QTY+(B.QTY*(B.LOSS/100)))) AS TESTQTY      " & vbNewLine & _
           "              , SUM(A.FREECNT*(B.QTY+(B.QTY*(B.LOSS/100)))) AS FREEQTY      " & vbNewLine & _
           "              , SUM(A.QCCNT*(B.QTY+(B.QTY*(B.LOSS/100)))) AS QCQTY          " & vbNewLine & _
           "              , SUM(A.RETESTCNT*(B.QTY+(B.QTY*(B.LOSS/100)))) AS RETESTQTY  " & vbNewLine & _
           "              , SUM(A.MANUCNT*(B.QTY+(B.QTY*(B.LOSS/100)))) AS MANUQTY      " & vbNewLine & _
           "  FROM S2PIS311 A INNER JOIN S2PIS101 B ON A.TESTCD=B.TESTCD                " & vbNewLine & _
           " WHERE A.WORKDT='" & sDate & "' GROUP BY B.STKCD ORDER BY B.STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                prgBar.Value = 0:       prgBar.Max = .RecordCount
                While (Not .EOF) And sReturn
                    prgBar.Value = prgBar.Value + 1
                    
                    cPis313.workdt = sDate
                    cPis313.stkcd = "" & .Fields("STKCD").Value
                    cPis313.testqty = Val("" & .Fields("TESTQTY").Value)
                    cPis313.freeqty = Val("" & .Fields("FREEQTY").Value)
                    cPis313.qcqty = Val("" & .Fields("QCQTY").Value)
                    cPis313.retestqty = Val("" & .Fields("RETESTQTY").Value)
                    cPis313.manuqty = Val("" & .Fields("MANUQTY").Value)
                    cPis313.empid = gUserId
                    
                    sReturn = cPis313.cfSave(0)
        
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam

    ' 장비운영 집계
    Set cPis302 = New clsPis302
    sDay = Format(dtpDt.Value, "dd")
    sWeek = Weekday(dtpDt.Value)

    lblMsg.Caption = "... 장비운영일정 등록 처리중 ...":        lblMsg.Refresh
    ' 장비운영일정에 등록된 자료를 기준으로 장비운영자료를 자동 생성한다.
    gSql = "SELECT A.* FROM S2PIS103 A                                                  " & vbNewLine & _
           " WHERE (A.CYCLEFG='0') OR (A.CYCLEFG='1' AND A.CYCLEDAY='" & sWeek & "')    " & vbNewLine & _
           "       OR (A.CYCLEFG='2' AND CYCLEDAY='" & sDay & "')                       " & vbNewLine & _
           "   AND A.PAUSEFG='0' AND A.STARTDT <='" & sDate & "'                        " & vbNewLine & _
           "   AND (A.ENDDT IS NULL OR A.ENDDT>='" & sDate & "')"
    If gWorkArea Then
        gSql = gSql & " AND EXISTS(SELECT B.EQPCD FROM S2LAB006 B WHERE A.EQPCD=B.EQPCD)"
    Else
        gSql = gSql & " AND EXISTS(SELECT B.MACHCODE FROM " & gKahpUser & "TWMED_MACHINE B WHERE A.EQPCD=B.MACHCODE)"
    End If
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                prgBar.Value = 0:       prgBar.Max = .RecordCount
                While (Not .EOF) And sReturn
                    prgBar.Value = prgBar.Value + 1
                    
                    cPis302.workdt = sDate
                    cPis302.eqpcd = "" & .Fields("EQPCD").Value
                    cPis302.seq = 0
                    cPis302.opercd = "" & .Fields("OPERCD").Value
                    cPis302.workcnt = Val("" & .Fields("OPERCNT").Value)
                    cPis302.autofg = "1"
                    cPis302.remark = "(AUTO) 장비운영일정"
                    cPis302.empid = gUserId
                    
                    sReturn = cPis302.cfSave
                        
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam
    
    lblMsg.Caption = "... 장비운영 집계 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT A.EQPCD,A.OPERCD,SUM(A.WORKCNT) AS CNT FROM S2PIS302 A WHERE A.WORKDT='" & sDate & "'" & vbNewLine
    If gWorkArea Then
        gSql = gSql & " AND EXISTS(SELECT B.EQPCD FROM S2LAB006 B WHERE A.EQPCD=B.EQPCD)"
    Else
        gSql = gSql & " AND EXISTS(SELECT B.MACHCODE FROM " & gKahpUser & "TWMED_MACHINE B WHERE A.EQPCD=B.MACHCODE)"
    End If
    gSql = gSql & " GROUP BY A.EQPCD,A.OPERCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                prgBar.Value = 0:       prgBar.Max = .RecordCount
                While (Not .EOF) And sReturn
                    prgBar.Value = prgBar.Value + 1
                    
                    cPis312.workdt = sDate
                    cPis312.eqpcd = "" & .Fields("EQPCD").Value
                    cPis312.opercd = "" & .Fields("OPERCD").Value
                    cPis312.opercnt = "" & .Fields("CNT").Value
                    cPis312.empid = gUserId
                    
                    sReturn = cPis312.cfSave
                        
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam
    
    ' 장비별 소요품목
    lblMsg.Caption = "... 장비별 소요품목 출고 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT B.STKCD, SUM(A.OPERCNT*(B.QTY+(B.QTY*(B.LOSS/100)))) AS QTY FROM S2PIS312 A      " & vbNewLine & _
           "       INNER JOIN S2PIS102 B ON A.EQPCD=B.EQPCD AND A.OPERCD=B.OPERCD                   " & vbNewLine & _
           " WHERE A.WORKDT='" & sDate & "' GROUP BY B.STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                prgBar.Value = 0:       prgBar.Max = .RecordCount
                While (Not .EOF) And sReturn
                    prgBar.Value = prgBar.Value + 1
                    
                    cPis313.workdt = sDate
                    cPis313.stkcd = "" & .Fields("STKCD").Value
                    cPis313.machqty = Val("" & .Fields("QTY").Value)
                    cPis313.empid = gUserId
                    
                    sReturn = cPis313.cfSave(4)
        
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam
    
    ' 수기출고 집계
    lblMsg.Caption = "... 수기 출고 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT STKCD,SUM(QTY) AS QTY FROM S2PIS303 WHERE WORKDT='" & sDate & "'" & vbNewLine & _
           "   AND LOTFG='0' GROUP BY WORKDT,STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                prgBar.Value = 0:       prgBar.Max = .RecordCount
                While (Not .EOF) And sReturn
                    prgBar.Value = prgBar.Value + 1
                    
                    cPis313.workdt = sDate
                    cPis313.stkcd = "" & .Fields("STKCD").Value
                    cPis313.handqty = Val("" & .Fields("QTY").Value)
                    cPis313.empid = gUserId
                    
                    sReturn = cPis313.cfSave(5)
                        
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam
    
    ' 수기 LOT출고처리
    lblMsg.Caption = "... 수기 LOT출고 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT STKCD,SUM(QTY) AS QTY FROM S2PIS303 WHERE WORKDT='" & sDate & "'" & vbNewLine & _
           "   AND LOTFG='1' GROUP BY WORKDT,STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                prgBar.Value = 0:       prgBar.Max = .RecordCount
                While (Not .EOF) And sReturn
                    prgBar.Value = prgBar.Value + 1
                    
                    cPis313.workdt = sDate
                    cPis313.stkcd = "" & .Fields("STKCD").Value
                    cPis313.lotnoqty = Val("" & .Fields("QTY").Value)
                    cPis313.empid = gUserId
                    
                    sReturn = cPis313.cfSave(6)
                        
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam

    ' LOT지정출고 처리
    lblMsg.Caption = "... LOT지정 출고 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT A.*, B.ENTDT, B.ENTSEQ FROM S2PIS303 A                                       " & vbNewLine & _
           "       INNER JOIN S2PIS401 B ON A.CHULDT=B.CHULDT AND A.CHULSEQ=B.CHULSEQ           " & vbNewLine & _
           " WHERE A.WORKDT='" & sDate & "' AND A.LOTFG='1' ORDER BY A.WORKDT,A.STKCD,A.SEQ"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF) And sReturn
                cPis314.workdt = sDate
                cPis314.stkcd = "" & .Fields("STKCD").Value
                cPis314.seq = 0
                cPis314.outqty = Val("" & .Fields("QTY").Value)
                cPis314.empid = gUserId
                cPis314.chuldt = "" & .Fields("CHULDT").Value
                cPis314.chulseq = Val("" & .Fields("CHULSEQ").Value)
                sReturn = cPis314.cfSave
                
                If sReturn Then
                    gSql = "UPDATE S2PIS401 SET USEQTY=NVL(USEQTY,0)+" & Val("" & .Fields("QTY").Value) & _
                           " WHERE CHULDT='" & .Fields("CHULDT").Value & "' AND CHULSEQ=" & .Fields("CHULSEQ").Value
                    sReturn = cDb.cfExecute(gSql)
                End If
                
                .MoveNext
            Wend
            .Close
        End If
    End With
    
    ' 품목 LOT 출고처리
    lblMsg.Caption = "... 재고 출고 처리중 ...":        lblMsg.Refresh
    gSql = "SELECT * FROM S2PIS313 WHERE WORKDT='" & sDate & "' ORDER BY STKCD"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                While (Not .EOF) And sReturn
                    sUseQty = Val("" & .Fields("TESTQTY").Value) + Val("" & .Fields("FREEQTY").Value) + Val("" & .Fields("QCQTY").Value) _
                              + Val("" & .Fields("RETESTQTY").Value) + Val("" & .Fields("MANUQTY").Value) + Val("" & .Fields("MACHQTY").Value) _
                              + Val("" & .Fields("HANDQTY").Value)
                              
                    ' 현장재고 불출처리
                    sReturn = pfStkLotChulgo(sDate, "" & .Fields("STKCD").Value, sUseQty)
                    
                    .MoveNext
                Wend
            End If
            .Close
        End If
    End With
    If sReturn = False Then GoTo errMagam

    Call cDb.csCommit
    
    grpBar.Visible = False

    Call cmdFind_Click

    MousePointer = vbDefault
    MsgBox "마감이 완료되었습니다.!", vbInformation
    Exit Sub
    
errMagam:
    Call cDb.csRollback
    
    lblMsg.Caption = ""
    MousePointer = vbDefault
    MsgBox "마감중 오류가 발생되어 마감이 중단되었습니다.!", vbCritical
    grpBar.Visible = False

End Sub

Private Function pfStkLotChulgo(ByVal brDt As String, ByVal brStk As String, ByVal brQty As Double) As Boolean
Dim cPis314 As clsPis314
Dim sRmdQty As Double, sUseQty As Double, sJaegoQty As Double, sReturn As Boolean, sEndFgStr As String
    
    ' 출고량을 기준으로 LOT(현장불출)기준으로 유효기간이 짧은 순차적으로 재고를 불출한다
    sRmdQty = brQty
    sReturn = True
    
    Set cPis314 = New clsPis314
    gSql = "SELECT A.*, B.EXPIRYDT FROM S2PIS401 A                                                              " & vbNewLine & _
           "       INNER JOIN S2PIS201 B ON A.ENTDT=B.ENTDT AND A.ENTSEQ=B.ENTSEQ                               " & vbNewLine & _
           " WHERE A.STKCD='" & brStk & "' AND (A.ENDFG IS NULL OR A.ENDFG<>'1' OR A.ENTQTY>NVL(A.USEQTY,0))    " & vbNewLine & _
           "   AND A.CHULDT<='" & brDt & "' ORDER BY B.EXPIRYDT,A.CHULDT,A.CHULSEQ"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            While (Not .EOF) And (sRmdQty > 0) And sReturn
                sEndFgStr = ""
                sJaegoQty = 0
                sUseQty = 0
            
                sJaegoQty = Val("" & .Fields("ENTQTY").Value) - Val("" & .Fields("USEQTY").Value)
                If sRmdQty <= sJaegoQty Then
                    ' 사용량이 재고보다 이하일경우
                    sUseQty = sRmdQty
                    sRmdQty = 0
                Else
                    sUseQty = sJaegoQty
                    sRmdQty = sRmdQty - sUseQty
                    
                    sEndFgStr = ",ENDFG='1'"
                End If
                
                cPis314.workdt = brDt
                cPis314.stkcd = brStk
                cPis314.seq = 0
                cPis314.outqty = sUseQty
                cPis314.empid = gUserId
                cPis314.chuldt = "" & .Fields("CHULDT").Value
                cPis314.chulseq = Val("" & .Fields("CHULSEQ").Value)
                
                sReturn = cPis314.cfSave
                
                If sReturn Then
                    gSql = "UPDATE S2PIS401 SET USEQTY=NVL(USEQTY,0)+" & sUseQty & sEndFgStr & _
                           " WHERE CHULDT='" & .Fields("CHULDT").Value & "' AND CHULSEQ=" & .Fields("CHULSEQ").Value
                    sReturn = cDb.cfExecute(gSql)
                End If
                
                .MoveNext
            Wend
        End If
    End With
    
    If sRmdQty > 0 Then
        ' 현장불출된 재고가 부족한 경우
        sReturn = False
        MsgBox brStk & " 품목의 현장재고가 " & gfQtyOutputStr(sRmdQty) & " 만큼 부족한 상태입니다.!" & _
               vbNewLine & "창고 현장불출을 추가하신 후 마감하세요.!", vbCritical
    End If
    
    pfStkLotChulgo = sReturn

End Function

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    Call cmdClear_Click

End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2
    grpMain.Height = Me.ScaleHeight - 50
    spStk.Height = (grpMain.Height - spStk.Top) - 50
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

