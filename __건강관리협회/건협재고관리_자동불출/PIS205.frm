VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "PVNum.ocx"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XCombobox.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS205 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LOT별 출고내역등록"
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
      Begin BHButton.BHImageButton cmdChulgoList 
         Height          =   375
         Left            =   12300
         TabIndex        =   39
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "출고내역"
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
         TransparentPicture=   "PIS205.frx":0000
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdChulgo 
         Height          =   375
         Left            =   11070
         TabIndex        =   10
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "출고처리"
         CaptionChecked  =   "BHImageButton2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS205.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13530
         TabIndex        =   9
         Top             =   840
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
         TransparentPicture=   "PIS205.frx":2F84
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   9840
         TabIndex        =   8
         Top             =   840
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
         TransparentPicture=   "PIS205.frx":4746
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   8610
         TabIndex        =   7
         Top             =   840
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
         TransparentPicture=   "PIS205.frx":5F08
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XGroupBox.XGroupBox grpFind 
         Height          =   675
         Left            =   90
         Top             =   90
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1191
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
         Begin TDBDate6Ctl.TDBDate dtpTodt 
            Height          =   315
            Left            =   12300
            TabIndex        =   51
            Top             =   180
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "PIS205.frx":76CA
            Caption         =   "PIS205.frx":77B1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS205.frx":7814
            Keys            =   "PIS205.frx":7832
            Spin            =   "PIS205.frx":7890
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
         Begin TDBDate6Ctl.TDBDate dtpFrdt 
            Height          =   315
            Left            =   10860
            TabIndex        =   50
            Top             =   180
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "PIS205.frx":78B8
            Caption         =   "PIS205.frx":799F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS205.frx":7A02
            Keys            =   "PIS205.frx":7A20
            Spin            =   "PIS205.frx":7A7E
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
         Begin BHButton.BHImageButton cmdStk 
            Height          =   315
            Left            =   2670
            TabIndex        =   11
            Top             =   180
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            Caption         =   "..."
            CaptionChecked  =   "BHImageButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "PIS205.frx":7AA6
            BackColor       =   14737632
            AlphaColor      =   16777215
            ImgOutLineSize  =   3
         End
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   9450
            TabIndex        =   6
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "유효기한"
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
         Begin XLibrary_XTextBox.XTextBox txtLotno 
            Height          =   315
            Left            =   7710
            TabIndex        =   5
            Top             =   180
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            BackColor       =   16777215
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
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
            Locked          =   0   'False
            Mask            =   0
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XTextBox.XTextBox txtNm 
            Height          =   315
            Left            =   3030
            TabIndex        =   4
            Top             =   180
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   556
            BackColor       =   14737632
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
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
            TextAlign       =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   180
            TabIndex        =   3
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "입고품목"
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
         Begin XLibrary_XTextBox.XTextBox txtCd 
            Height          =   315
            Left            =   1590
            TabIndex        =   2
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            BackColor       =   16777215
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
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
            Locked          =   0   'False
            Mask            =   0
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel1 
            Height          =   315
            Left            =   6300
            TabIndex        =   1
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "입고 LOT NO"
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
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   8295
         Left            =   90
         TabIndex        =   0
         Tag             =   "20001"
         Top             =   1290
         Width           =   14640
         _Version        =   524288
         _ExtentX        =   25823
         _ExtentY        =   14631
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         EditModeReplace =   -1  'True
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
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS205.frx":9268
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
   Begin XLibrary_XGroupBox.XGroupBox grpChulgo 
      Height          =   3285
      Left            =   4290
      Top             =   3180
      Visible         =   0   'False
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   5794
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
      Begin PVNumericLib.PVNumeric txtQty 
         Height          =   315
         Left            =   4290
         TabIndex        =   37
         Top             =   1620
         Width           =   1665
         _Version        =   524288
         _ExtentX        =   2937
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   1
         EditMode        =   0
         SpinButtons     =   0
         SuppressThousand=   0   'False
         LimitValue      =   -1  'True
      End
      Begin TDBDate6Ctl.TDBDate dtpDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   52
         Top             =   1620
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   556
         Calendar        =   "PIS205.frx":9AC2
         Caption         =   "PIS205.frx":9BA9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PIS205.frx":9C0C
         Keys            =   "PIS205.frx":9C2A
         Spin            =   "PIS205.frx":9C88
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
      Begin BHButton.BHImageButton cmdCancel 
         Height          =   375
         Left            =   4770
         TabIndex        =   36
         Top             =   2820
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "취 소"
         CaptionChecked  =   "BHImageButton3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS205.frx":9CB0
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   3510
         TabIndex        =   35
         Top             =   2820
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "저 장"
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
         TransparentPicture=   "PIS205.frx":B472
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XTextBox.XTextBox XTextBox1 
         Height          =   405
         Left            =   30
         TabIndex        =   34
         Top             =   30
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   714
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
         Text            =   "LOT별 수기 출고"
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
      Begin XLibrary_XTextBox.XTextBox txtChulseq 
         Height          =   315
         Left            =   1440
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2790
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "123456789012345"
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
         TextAlign       =   0
         Enabled         =   -1  'True
      End
      Begin XLibrary_XTextBox.XTextBox txtChuldt 
         Height          =   315
         Left            =   360
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2790
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "123456789012345"
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
         TextAlign       =   0
         Enabled         =   -1  'True
      End
      Begin XLibrary_XTextBox.XTextBox txtStkcd 
         Height          =   315
         Left            =   1320
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   540
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "123456789012345"
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
         TextAlign       =   0
         Enabled         =   -1  'True
      End
      Begin XLibrary_XLabel.XLabel XLabel2 
         Height          =   315
         Left            =   150
         TabIndex        =   30
         Top             =   540
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "품목코드"
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
      Begin XLibrary_XTextBox.XTextBox txtStknm 
         Height          =   315
         Left            =   2400
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   540
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "123456789012345"
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
         TextAlign       =   0
         Enabled         =   -1  'True
      End
      Begin XLibrary_XTextBox.XTextBox txtLot 
         Height          =   315
         Left            =   1320
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   900
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "123456789012345"
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
         TextAlign       =   0
         Enabled         =   -1  'True
      End
      Begin XLibrary_XLabel.XLabel XLabel3 
         Height          =   315
         Left            =   150
         TabIndex        =   27
         Top             =   900
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "LOT NO"
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
      Begin XLibrary_XTextBox.XTextBox txtExpirydt 
         Height          =   315
         Left            =   1320
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "123456789012345"
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
         TextAlign       =   0
         Enabled         =   -1  'True
      End
      Begin XLibrary_XLabel.XLabel XLabel5 
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   1260
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "유효기한"
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
      Begin XLibrary_XTextBox.XTextBox txtEntqty 
         Height          =   315
         Left            =   4290
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   900
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "123456789012345"
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
      Begin XLibrary_XLabel.XLabel XLabel7 
         Height          =   315
         Left            =   3120
         TabIndex        =   23
         Top             =   900
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "불출량"
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
      Begin XLibrary_XTextBox.XTextBox txtUseqty 
         Height          =   315
         Left            =   4290
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "123456789012345"
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
      Begin XLibrary_XLabel.XLabel XLabel8 
         Height          =   315
         Left            =   3120
         TabIndex        =   21
         Top             =   1260
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "총사용량"
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
      Begin XLibrary_XLabel.XLabel XLabel9 
         Height          =   315
         Left            =   150
         TabIndex        =   20
         Top             =   1620
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "사용일자"
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
      Begin XLibrary_XLabel.XLabel XLabel10 
         Height          =   315
         Left            =   3120
         TabIndex        =   19
         Top             =   1620
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "사용량"
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
      Begin XLibrary_XLabel.XLabel XLabel11 
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   1980
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "출고사유"
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
      Begin XLibrary_XTextBox.XTextBox txtRemark 
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   2340
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         BackColor       =   16777215
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
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
         Locked          =   0   'False
         Mask            =   0
         PromptChar      =   "_"
         WrongSound      =   0
         CustomSound     =   ""
         MaskShow        =   0   'False
         MaskColor       =   33023
         CustomMask      =   ""
         TextAlign       =   0
         Enabled         =   -1  'True
      End
      Begin XLibrary_XLabel.XLabel XLabel12 
         Height          =   315
         Left            =   150
         TabIndex        =   16
         Top             =   2340
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "비고사항"
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
      Begin XLibrary_XComboBox.XComboBox cboReason 
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Top             =   1980
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         BackColor       =   16777215
         BorderColor     =   16744576
         BtnBackColor1   =   16777215
         BtnBackStyle    =   3
         Text            =   ""
         BtnBorderColor  =   12632256
         BtnBorderStyle  =   1
         BtnBackColor2   =   15000804
         BtnSymbolColor  =   8388608
         BtnSymbolStyle  =   2
         UpListShow      =   0   'False
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowItemNum     =   5
         AutoSel         =   0   'False
         TextEdit        =   0   'False
         BtnMouseCursor  =   2
         ToolTipIcon     =   1
         ToolTipPopupTime=   -1
         ToolTipHoverTime=   800
         ToolTipBackColor=   16777215
         ToolTipForeColor=   0
         ToolTipOpacity  =   100
         ToolTipStyle    =   2
         ToolTipCentered =   0   'False
         ToolTipTitleText=   ""
         ToolTipBodyText =   ""
         TextColor       =   0
         ListBgColor     =   16777215
         ListTextColor   =   0
         Enabled         =   -1  'True
      End
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   4410
      ScaleHeight     =   3285
      ScaleWidth      =   6180
      TabIndex        =   38
      Top             =   3300
      Visible         =   0   'False
      Width           =   6180
   End
   Begin XLibrary_XGroupBox.XGroupBox grpList 
      Height          =   8295
      Left            =   1650
      Top             =   735
      Visible         =   0   'False
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   14631
      BackColor       =   12640511
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
      Begin Threed.SSFrame SSFrame1 
         Height          =   435
         Left            =   60
         TabIndex        =   43
         Top             =   510
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   767
         _Version        =   262144
         BackColor       =   16311512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSPanel SSPanel2 
            Height          =   300
            Left            =   90
            TabIndex        =   44
            Top             =   60
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16311512
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "불출일자"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblChulDt 
            Height          =   300
            Left            =   1020
            TabIndex        =   45
            Top             =   60
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "불출일자"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   300
            Left            =   2250
            TabIndex        =   46
            Top             =   60
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16311512
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "불출번호"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblChulNo 
            Height          =   300
            Left            =   3180
            TabIndex        =   47
            Top             =   60
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "불출일자"
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   300
            Left            =   3900
            TabIndex        =   48
            Top             =   60
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16311512
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "출고품목"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel lblChulStk 
            Height          =   300
            Left            =   4830
            TabIndex        =   49
            Top             =   60
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   529
            _Version        =   262144
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "불출일자"
            BevelOuter      =   0
            Alignment       =   1
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   435
         Left            =   60
         TabIndex        =   42
         Top             =   60
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   767
         _Version        =   262144
         BackColor       =   16249839
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "☞ 품목 입고 LOT별 출고내역 ☜"
         BorderWidth     =   1
         BevelInner      =   2
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   300
         Left            =   60
         TabIndex        =   41
         Top             =   7920
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   529
         _Version        =   262144
         ForeColor       =   255
         BackColor       =   16311512
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "※ 상기리스트는 최종마감일자 이후의 자료만 조회하여 작업하실 수 있습니다 !! ※"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin BHButton.BHImageButton cmdCupdate 
         Height          =   375
         Left            =   7950
         TabIndex        =   40
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "수 정"
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
         TransparentPicture=   "PIS205.frx":CC34
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin FPSpreadADO.fpSpread spChulgo 
         CausesValidation=   0   'False
         Height          =   6885
         Left            =   60
         TabIndex        =   14
         Tag             =   "20001"
         Top             =   960
         Width           =   11550
         _Version        =   524288
         _ExtentX        =   20373
         _ExtentY        =   12144
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         EditModeReplace =   -1  'True
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
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS205.frx":E3F6
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
      Begin BHButton.BHImageButton cmdCclose 
         Height          =   375
         Left            =   10410
         TabIndex        =   13
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "닫 기"
         CaptionChecked  =   "BHImageButton2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS205.frx":F075
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdCdelete 
         Height          =   375
         Left            =   9180
         TabIndex        =   12
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "삭 제"
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
         TransparentPicture=   "PIS205.frx":10837
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
   End
End
Attribute VB_Name = "PIS205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fReason() As String

Private Sub cmdCancel_Click()

    grpMain.Enabled = True
    picShadow.Visible = False
    grpChulgo.Visible = False

End Sub

Private Sub cmdCclose_Click()

    grpMain.Enabled = True
    grpList.Visible = False
    
    If cmdCclose.Tag = "1" Then
        Call cmdFind_Click
    End If
    cmdCclose.Tag = ""

End Sub

Private Sub cmdCdelete_Click()

    MousePointer = vbHourglass
    If MsgBox("선택하신 출고내역을 삭제하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Call psChulgoProcess(False)
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdChulgo_Click()
Dim sGetVal As Variant

    With spList
        If .ActiveCol > 0 And .ActiveRow > 0 Then
            .GetText 1, .ActiveRow, sGetVal:        txtChuldt.Text = Trim(sGetVal)
            .GetText 4, .ActiveRow, sGetVal:        txtStkcd.Text = Trim(sGetVal)
            .GetText 5, .ActiveRow, sGetVal:        txtStkNm.Text = Trim(sGetVal)
            .GetText 6, .ActiveRow, sGetVal:        txtLot.Text = Trim(sGetVal)
            .GetText 7, .ActiveRow, sGetVal:        txtExpirydt.Text = Trim(sGetVal)
            .GetText 9, .ActiveRow, sGetVal:        txtEntqty.Text = Trim(sGetVal):     txtEntqty.Tag = Val(Str(txtEntqty.Text))
            .GetText 10, .ActiveRow, sGetVal:       txtUseqty.Text = Trim(sGetVal):     txtUseqty.Tag = Val(Str(txtUseqty.Text))
            .GetText 11, .ActiveRow, sGetVal:       txtChulseq.Text = Val(sGetVal)
            
            txtQty.Text = Val(txtEntqty.Tag) - Val(txtUseqty.Tag)
            
            dtpDate.Value = gfSystemDate
            
            grpMain.Enabled = False
            
            Call gsReasonCombo(cboReason, gReasonManual)
            cboReason.ListIndex = -1
            picShadow.Visible = True:       picShadow.ZOrder 0
            grpChulgo.Visible = True:       grpChulgo.ZOrder 0
        End If
    End With
    
End Sub

Private Sub cmdChulgoList_Click()
Dim cPis303 As clsPis303, cPis006
Dim sDate As String, sSeq As Long, sGetVal As Variant, sRow As Long, sStr As String

    Call gsSpreadClear(spChulgo, 0, True)

    grpMain.Enabled = False
    grpList.Visible = True
    grpList.ZOrder 0
    
    sStr = "":      sRow = 0
    Set cPis006 = New clsPis006
    With cPis006.cfList(True, gReasonManual)
        If .State = adStateOpen Then
            If Not .EOF Then
                ReDim fReason(.RecordCount) As String
                
                While (Not .EOF)
                    sStr = sStr & .Fields("REASONNM").Value
                    fReason(sRow) = "" & .Fields("REASONCD").Value
                    sRow = sRow + 1
                    
                    .MoveNext
                    If Not .EOF Then
                        sStr = sStr & vbTab
                    End If
                Wend
            End If
            .Close
        End If
    End With
    
    spChulgo.Row = -1
    spChulgo.Col = 6
    spChulgo.TypeComboBoxList = sStr
    
    sRow = 0
    
    spList.GetText 1, spList.ActiveRow, sGetVal:    sDate = Format(sGetVal, "yyyyMMdd")
    lblChulDt.Caption = sGetVal
    spList.GetText 11, spList.ActiveRow, sGetVal:   sSeq = Val(sGetVal)
    lblChulNo.Caption = sGetVal
    spList.GetText 4, spList.ActiveRow, sGetVal:    lblChulStk.Tag = Trim(sGetVal)
    spList.GetText 5, spList.ActiveRow, sGetVal:    lblChulStk.Caption = Trim(sGetVal)
    
    Set cPis303 = New clsPis303
    With cPis303.cfLotList(sDate, sSeq, gfMagamMaxDate)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spChulgo, .RecordCount, True)
                
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spChulgo.SetText 1, sRow, ""
                    spChulgo.SetText 2, sRow, Format(.Fields("WORKDT").Value, "####-##-##")
                    spChulgo.SetText 3, sRow, "" & .Fields("SEQ").Value
                    spChulgo.SetText 4, sRow, "" & .Fields("QTY").Value
                    spChulgo.SetText 5, sRow, "" & .Fields("QTY").Value
                    spChulgo.SetText 6, sRow, "" & .Fields("REASONNM").Value
                    spChulgo.SetText 7, sRow, "" & .Fields("REASONCD").Value
                    spChulgo.SetText 8, sRow, "" & .Fields("REMARK").Value
                    spChulgo.SetText 9, sRow, "" & .Fields("EMPNM").Value
                    spChulgo.SetText 10, sRow, "" & .Fields("WRTDT").Value
                    spChulgo.SetText 11, sRow, "" & .Fields("MODDT").Value
                
                    .MoveNext
                Wend
            Else
                MsgBox "출고내역이 없습니다.!", vbCritical
            End If
            .Close
        End If
    End With
    
    If spList.MaxRows = 0 Then
        grpMain.Enabled = True
        grpList.Visible = False
    End If

End Sub

Private Sub cmdClear_Click()

    dtpFrdt.Value = Format(gfSystemDate, "yyyy-MM") & "-01"
    dtpTodt.Value = gfSystemDate
    dtpTodt.Value = ""
    
    Call gsSpreadClear(spList, 0, True)
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdChulgo, False)
    Call gsButtonEnable(cmdChulgoList, False)
    
    grpFind.Enabled = True
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub psChulgoProcess(ByVal brSave As Boolean)
Dim cPis303 As clsPis303
Dim sRow As Long, sGetVal As Variant, sReturn As Boolean

    Set cPis303 = New clsPis303
    With spChulgo
        For sRow = .MaxRows To 1 Step -1
            .GetText 2, sRow, sGetVal:          cPis303.workdt = Format(sGetVal, "yyyyMMdd")
            .GetText 1, sRow, sGetVal
            If Val(sGetVal) > 0 And Len(cPis303.workdt) > 0 Then
                .GetText 3, sRow, sGetVal:      cPis303.seq = Val(sGetVal)
                
                cPis303.chuldt = Format(lblChulDt.Caption, "yyyyMMdd")
                cPis303.chulseq = Val(lblChulNo.Caption)
                cPis303.stkcd = lblChulStk.Tag
                cPis303.empid = gUserId
                
                cDb.csBegin
                If brSave Then
                    .GetText 4, sRow, sGetVal:  cPis303.qty = Val(sGetVal)
                    .GetText 7, sRow, sGetVal:  cPis303.reasoncd = Trim(sGetVal)
                    .GetText 8, sRow, sGetVal:  cPis303.remark = Trim(sGetVal)
                    sReturn = cPis303.cfSave
                Else
                    .GetText 5, sRow, sGetVal:  cPis303.qty = Val(sGetVal)
                    sReturn = cPis303.cfDelete
                End If
                
                If sReturn Then
                    Call cDb.csCommit
                Else
                    Call cDb.csRollback
                    Exit For
                End If
            End If
        Next sRow
    End With
    
    If sReturn Then
        Call cmdChulgoList_Click
        cmdCclose.Tag = "1"
    End If
    
End Sub

Private Sub cmdCupdate_Click()

    MousePointer = vbHourglass
    Call psChulgoProcess(True)
    MousePointer = vbDefault

End Sub

Private Sub cmdFind_Click()
Dim sRow As Long

    MousePointer = vbHourglass
    gSql = "SELECT A.*, B.CSTCD, B.LOTNO, B.EXPIRYDT, B.IQTY_SO, X.NM_ITEM AS STKNM, D.CSTNM FROM S2PIS401 A    " & vbNewLine & _
           "      INNER JOIN S2PIS201 B ON A.ENTDT=B.ENTDT AND A.ENTSEQ=B.ENTSEQ                                " & vbNewLine & _
           "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM" & gERPStkCondition & "                     " & vbNewLine & _
           "       LEFT JOIN S2PIS002 D ON B.CSTCD=D.CSTCD                                                      " & vbNewLine & _
           " WHERE (A.ENDFG IS NULL OR A.ENDFG<>'1')"
    If Len(txtCd.Text) > 0 Then
        gSql = gSql & "  AND A.STKCD='" & Trim(txtCd.Text) & "'"
    End If
    If Len(txtLotno.Text) > 0 Then
        gSql = gSql & "  AND B.LOTNO='%" & Trim(txtLotno.Text) & "%'"
    End If
    If Not IsNull(dtpFrdt.Value) Then
        gSql = gSql & "  AND B.EXPIRYDT>='" & Format(dtpFrdt.Value, "yyyyMMdd") & "'"
    End If
    If Not IsNull(dtpTodt.Value) Then
        gSql = gSql & "  AND B.EXPIRYDT<='" & Format(dtpTodt.Value, "yyyyMMdd") & "'"
    End If
    gSql = gSql & " ORDER BY A.CHULDT,A.STKCD,B.LOTNO"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    
                    spList.SetText 1, sRow, Format("" & .Fields("CHULDT").Value, "####-##-##")
                    spList.SetText 2, sRow, "" & .Fields("CSTCD").Value
                    spList.SetText 3, sRow, "" & .Fields("CSTNM").Value
                    spList.SetText 4, sRow, "" & .Fields("STKCD").Value
                    spList.SetText 5, sRow, "" & .Fields("STKNM").Value
                    spList.SetText 6, sRow, "" & .Fields("LOTNO").Value
                    spList.SetText 7, sRow, Format("" & .Fields("EXPIRYDT").Value, "####-##-##")
                    spList.SetText 8, sRow, gfQtyOutputStr(Val("" & .Fields("IQTY_SO").Value))
                    spList.SetText 9, sRow, gfQtyOutputStr(Val("" & .Fields("ENTQTY").Value))
                    spList.SetText 10, sRow, gfQtyOutputStr(Val("" & .Fields("USEQTY").Value))
                    spList.SetText 11, sRow, "" & .Fields("CHULSEQ").Value
                
                    .MoveNext
                Wend
                Call gsButtonEnable(cmdFind, False)
                Call gsButtonEnable(cmdChulgo, True)
                Call gsButtonEnable(cmdChulgoList, True)
                
                grpFind.Enabled = False
            Else
                Call gsSpreadClear(spList, 0, True)
                MsgBox "조건에 해당하는 입고자료가 없습니다.!", vbCritical
            End If
            .Close
        End If
    End With
    MousePointer = vbDefault

End Sub

Private Sub cmdSave_Click()
Dim cPis303 As clsPis303, sReason() As String

    If cboReason.ListIndex < 0 Then
        MsgBox "출고사유를 선택하세요.!", vbCritical
        Exit Sub
    End If
    
    If gfMagamCheck(Format(dtpDate.Value, "yyyyMMdd"), True) = False Then
        MsgBox "마감된 일자입니다.!", vbCritical
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    If MsgBox("입력된 자료를 출고(사용)처리하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Set cPis303 = New clsPis303
        With cPis303
            .workdt = Format(dtpDate.Value, "yyyyMMdd")
            .seq = 0
            .stkcd = Trim(txtStkcd.Text)
            .qty = Val(txtQty.Text)
            If cboReason.ListIndex >= 0 Then
                sReason = Split(cboReason.Text, gCboSplitStr)
                .reasoncd = Trim(sReason(1))
            End If
            .remark = Trim(txtRemark.Text)
            .empid = gUserId
            .lotfg = "1"
            .chuldt = Format(txtChuldt.Text, "yyyyMMdd")
            .chulseq = Val(txtChulseq.Text)
            
            Call cDb.csBegin
            If .cfSave Then
                Call cDb.csCommit
                
                Call cmdFind_Click
                Call cmdCancel_Click
            Else
                Call cDb.csRollback
            End If
        End With
    End If
    MousePointer = vbDefault

End Sub

Private Sub cmdStk_Click()

    hlpStkList.Tag = "one"
    hlpStkList.Show vbModal
    
    If Len(gHelpCode) > 0 Then
        txtCd.Text = gHelpCode
        txtNm.Text = gfStkName(gHelpCode)
    Else
        txtCd.Text = ""
        txtNm.Text = ""
    End If
    
End Sub

Private Sub dtpDate_Change()
Dim sDate As String

    sDate = Format(gfMagamMaxDate, "####-##-##")
    
    If dtpDate.Value <= sDate Then
        MsgBox "마감완료된 일자입니다.!", vbCritical
        dtpDate.Value = DateAdd("d", 1, sDate)
    End If

End Sub

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    txtQty.DecimalMax = gDecimalQtyO
    txtQty.DecimalMin = gDecimalQtyO
    
'    spList.SetText 8, SpreadHeader, "입고량" & vbNewLine & "(출고단위)"
'    spList.SetText 9, SpreadHeader, "불출량" & vbNewLine & "(출고단위)"
'    spList.SetText 10, SpreadHeader, "사용량" & vbNewLine & "(출고단위)"
    
    Call cmdClear_Click
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2
    grpMain.Height = Me.ScaleHeight - 50
    spList.Height = (grpMain.Height - spList.Top) - 50
    
    grpChulgo.Left = (Me.ScaleWidth - grpChulgo.Width) / 2
    grpChulgo.Top = (Me.ScaleHeight - grpChulgo.Height) / 2
    
    picShadow.Left = grpChulgo.Left + 80
    picShadow.Top = grpChulgo.Top + 80
    
    grpList.Left = (Me.ScaleWidth - grpList.Width) / 2
    grpList.Top = (Me.ScaleHeight - grpList.Height) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

Private Sub spChulgo_Change(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 And Col > 1 Then
        spChulgo.SetText 1, Row, 1
    End If
    
End Sub

Private Sub spList_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 Then
        Call cmdChulgoList_Click
    End If
    
End Sub

Private Sub txtCd_LostFocus()

    txtNm.Text = gfStkName(Trim(txtCd.Text))
    If Len(txtNm.Text) = 0 Then txtCd.Text = ""

End Sub

Private Sub txtQty_lostFocus()

    If (Val(txtEntqty.Tag) - Val(txtUseqty.Tag)) < Val(txtQty.Text) Then
        MsgBox "잔량보다 많은 양을 사용할 수는 없습니다.!", vbCritical
        txtQty.Text = Val(txtEntqty.Tag) - Val(txtUseqty.Tag)
    End If

End Sub
