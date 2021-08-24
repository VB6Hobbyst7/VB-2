VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XCombobox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS914 
   BackColor       =   &H00FFFFFF&
   Caption         =   "∞À√º∆Û±‚(RACK¥‹¿ß)"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "±º∏≤√º"
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
   WindowState     =   2  '√÷¥Î»≠
   Begin XLibrary_XGroupBox.XGroupBox grpMain 
      Height          =   9675
      Left            =   30
      Top             =   30
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   17066
      BackColor       =   16777215
      BorderColor     =   10070188
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
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
      Begin BHButton.BHImageButton cmdScrap 
         Height          =   375
         Left            =   12330
         TabIndex        =   14
         Top             =   1170
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "∆Û±‚µÓ∑œ"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS914.frx":0000
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   9870
         TabIndex        =   13
         Top             =   1170
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "»≠∏È¡ˆøÚ"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS914.frx":17C2
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13560
         TabIndex        =   12
         Top             =   1170
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "¥› ±‚"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS914.frx":2F84
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   11100
         TabIndex        =   11
         Top             =   1170
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "¡∂ »∏"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS914.frx":4746
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   7455
         Left            =   90
         TabIndex        =   10
         Tag             =   "20001"
         Top             =   2160
         Width           =   14670
         _Version        =   524288
         _ExtentX        =   25876
         _ExtentY        =   13150
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   10
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS914.frx":5F08
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
      Begin XLibrary_XGroupBox.XGroupBox grpFind 
         Height          =   975
         Left            =   90
         Top             =   90
         Width           =   14670
         _ExtentX        =   25876
         _ExtentY        =   1720
         BackColor       =   16311512
         BorderColor     =   10070188
         BorderRoundNum  =   0
         BorderStyle     =   1
         TextColor       =   0
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤"
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
            Left            =   7320
            TabIndex        =   18
            Top             =   510
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calendar        =   "PIS914.frx":6A92
            Caption         =   "PIS914.frx":6B79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS914.frx":6BDC
            Keys            =   "PIS914.frx":6BFA
            Spin            =   "PIS914.frx":6C58
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
            Left            =   5760
            TabIndex        =   17
            Top             =   510
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calendar        =   "PIS914.frx":6C80
            Caption         =   "PIS914.frx":6D67
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS914.frx":6DCA
            Keys            =   "PIS914.frx":6DE8
            Spin            =   "PIS914.frx":6E46
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
         Begin TDBDate6Ctl.TDBDate dtpSvTo 
            Height          =   315
            Left            =   7320
            TabIndex        =   16
            Top             =   150
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calendar        =   "PIS914.frx":6E6E
            Caption         =   "PIS914.frx":6F55
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS914.frx":6FB8
            Keys            =   "PIS914.frx":6FD6
            Spin            =   "PIS914.frx":7034
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
         Begin TDBDate6Ctl.TDBDate dtpSvFr 
            Height          =   315
            Left            =   5760
            TabIndex        =   15
            Top             =   150
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calendar        =   "PIS914.frx":705C
            Caption         =   "PIS914.frx":7143
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS914.frx":71A6
            Keys            =   "PIS914.frx":71C4
            Spin            =   "PIS914.frx":7222
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
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   4440
            TabIndex        =   9
            Top             =   510
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "¿Ø»ø±‚«—"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
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
         Begin XLibrary_XTextBox.XTextBox txtRackno 
            Height          =   315
            Left            =   1620
            TabIndex        =   8
            Top             =   510
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            BackColor       =   16777215
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "10"
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
         Begin XLibrary_XLabel.XLabel XLabel17 
            Height          =   315
            Left            =   210
            TabIndex        =   7
            Top             =   510
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "RACK NO"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
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
            Left            =   210
            TabIndex        =   6
            Top             =   150
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "¿˙¿Â∞Ì"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
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
         Begin XLibrary_XComboBox.XComboBox cboDepot 
            Height          =   315
            Left            =   1620
            TabIndex        =   5
            Top             =   150
            Width           =   2625
            _ExtentX        =   4630
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
               Name            =   "±º∏≤√º"
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
         Begin XLibrary_XLabel.XLabel XLabel3 
            Height          =   315
            Left            =   4440
            TabIndex        =   4
            Top             =   150
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "∫∏∞¸¿œ¿⁄"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
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
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox1 
         Height          =   495
         Left            =   1470
         Top             =   1620
         Width           =   13290
         _ExtentX        =   23442
         _ExtentY        =   873
         BackColor       =   16311512
         BorderColor     =   10070188
         BorderRoundNum  =   0
         BorderStyle     =   1
         TextColor       =   0
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤"
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
         Begin TDBDate6Ctl.TDBDate dtpScrapdt 
            Height          =   315
            Left            =   1560
            TabIndex        =   19
            Top             =   90
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            Calendar        =   "PIS914.frx":724A
            Caption         =   "PIS914.frx":7331
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS914.frx":7394
            Keys            =   "PIS914.frx":73B2
            Spin            =   "PIS914.frx":7410
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
            Left            =   3000
            TabIndex        =   3
            Top             =   90
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "∆Û±‚ªÁ¿Ø"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
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
         Begin XLibrary_XTextBox.XTextBox txtReason 
            Height          =   315
            Left            =   4410
            TabIndex        =   2
            Top             =   90
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   556
            BackColor       =   16777215
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "10"
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
         Begin XLibrary_XLabel.XLabel XLabel8 
            Height          =   315
            Left            =   150
            TabIndex        =   1
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "∆Û±‚¿œ¿⁄"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤√º"
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
      Begin XLibrary_XTextBox.XTextBox XTextBox1 
         Height          =   495
         Left            =   90
         TabIndex        =   0
         Top             =   1620
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   873
         BackColor       =   16249839
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤√º"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "¢— ∆Û±‚¡§∫∏"
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
   End
End
Attribute VB_Name = "PIS914"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()

    Call gsSpreadClear(spList, 0, True)
    
    dtpFrdt.Value = gfSystemDate
    dtpFrdt.Value = ""
    dtpTodt.Value = gfSystemDate
    dtpTodt.Value = ""
    
    dtpSvFr.Value = gfSystemDate
'    dtpSvFr.Value = ""
    dtpSvTo.Value = gfSystemDate
'    dtpSvTo.Value = ""
    
    dtpScrapdt.Value = gfSystemDate
    
    cboDepot.ListIndex = 0
    txtRackno.Text = ""
    txtReason.Text = ""
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdScrap, False)
    
    grpFind.Enabled = True
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdFind_Click()
Dim sRow As Long, sDepot() As String

    MousePointer = vbHourglass
    If gWorkArea Then
        gSql = "SELECT A.*, B.DEPOTNM, C.EMPNM, D.ROWCNT, D.COLCNT                                              " & vbNewLine & _
               "     , (SELECT COUNT(*) FROM S2PIS902 Z WHERE Z.RACKNO=A.RACKNO GROUP BY Z.RACKNO) AS SPCCNT    " & vbNewLine & _
               "  FROM S2PIS901 A                                                                               " & vbNewLine & _
               "       LEFT JOIN S2PIS092 B ON A.DEPOTCD=B.DEPOTCD                                              " & vbNewLine & _
               "       LEFT JOIN S2COM006 C ON A.EMPID=C.EMPID                                                  " & vbNewLine & _
               "      INNER JOIN S2PIS091 D ON A.RACKNO=D.RACKNO                                                " & vbNewLine & _
               " WHERE a.SAVEDT <= '" & Format(dtpScrapdt.Value, "yyyyMMdd") & "' "
    Else
        gSql = "SELECT A.*, B.DEPOTNM, C.USER_NM AS EMPNM, D.ROWCNT, D.COLCNT                                   " & vbNewLine & _
               "     , (SELECT COUNT(*) FROM S2PIS902 Z WHERE Z.RACKNO=A.RACKNO GROUP BY Z.RACKNO) AS SPCCNT    " & vbNewLine & _
               "  FROM S2PIS901 A                                                                               " & vbNewLine & _
               "       LEFT JOIN S2PIS092 B ON A.DEPOTCD=B.DEPOTCD                                              " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " C ON A.EMPID=C.USERID                                   " & vbNewLine & _
               "      INNER JOIN S2PIS091 D ON A.RACKNO=D.RACKNO                                                " & vbNewLine & _
               " WHERE a.SAVEDT <= '" & Format(dtpScrapdt.Value, "yyyyMMdd") & "' "
    End If
    
    ' ∫∏∞¸¿œ¿⁄
    If Len(dtpSvFr.Value) > 0 Then
        gSql = gSql & " AND A.SAVEDT >= '" & Format(dtpSvFr.Value, "yyyyMMdd") & "'"
    End If
    If Len(dtpSvTo.Value) > 0 Then
        gSql = gSql & " AND A.SAVEDT <= '" & Format(dtpSvTo.Value, "yyyyMMdd") & "'"
    End If
    ' ∆Û±‚øπ¡§¿œ¿⁄
    If Len(dtpFrdt.Value) > 0 Then
        gSql = gSql & " AND A.EXPIRYDT >= '" & Format(dtpFrdt.Value, "yyyyMMdd") & "'"
    End If
    If Len(dtpTodt.Value) > 0 Then
        gSql = gSql & " AND A.EXPIRYDT <= '" & Format(dtpTodt.Value, "yyyyMMdd") & "'"
    End If
    ' Rack NO.
    If Len(txtRackno.Text) > 0 Then
        gSql = gSql & " AND A.RACKNO='" & Trim(txtRackno.Text) & "'"
    End If
    ' ¿˙¿Â∞Ì
    If cboDepot.ListIndex > 0 Then
        sDepot = Split(cboDepot.Text, gCboSplitStr)
        gSql = gSql & " AND A.DEPOTCD='" & Trim(sDepot(1)) & "'"
    End If
    gSql = gSql & " ORDER BY A.EXPIRYDT,A.SAVEDT,A.RACKNO"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                gPrgBar.Max = .RecordCount:     gPrgBar.Value = 0
                gPrgBar.Visible = True:         gPrgBar.Refresh
                
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1
                    gPrgBar = sRow
                    
                    spList.SetText 1, sRow, ""
                    spList.SetText 2, sRow, "" & .Fields("RACKNO").Value
                    spList.SetText 3, sRow, "" & .Fields("ROWCNT").Value & "/" & .Fields("COLCNT").Value
                    spList.SetText 4, sRow, Format("" & .Fields("SAVEDT").Value, "####-##-##")
                    spList.SetText 5, sRow, Format("" & .Fields("EXPIRYDT").Value, "####-##-##")
                    spList.SetText 6, sRow, "" & .Fields("DEPOTNM").Value
                    spList.SetText 7, sRow, "" & .Fields("SAVEFLOOR").Value
                    spList.SetText 8, sRow, Format(Val("" & .Fields("SPCCNT").Value), "#,##0")
                    spList.SetText 9, sRow, "" & .Fields("EMPNM").Value
                    spList.SetText 10, sRow, "" & .Fields("MODDT").Value
                    
                    .MoveNext
                Wend
                gPrgBar.Visible = False
                
                Call gsButtonEnable(cmdScrap, True)
                Call gsButtonEnable(cmdFind, False)
                
                grpFind.Enabled = False
            Else
                Call gsSpreadClear(spList, 0, True)
                MsgBox "¡∂∞«ø° «ÿ¥Á«œ¥¬ ¿⁄∑·∞° æ¯Ω¿¥œ¥Ÿ.!", vbCritical
            End If
            .Close
        End If
    End With
    MousePointer = vbDefault

End Sub

Private Sub cmdScrap_Click()
Dim cPis902 As clsPis902
Dim sRow As Long, sGetVal As Variant, sReturn As Boolean
    
    MousePointer = vbHourglass
    If MsgBox("º±≈√«œΩ≈ RACK¿ª ∆Û±‚√≥∏Æ«œΩ√∞⁄Ω¿¥œ±Ó ?", vbYesNo + vbQuestion) = vbYes Then
        Set cPis902 = New clsPis902
        With spList
            For sRow = 1 To .MaxRows
                .GetText 2, sRow, sGetVal:      cPis902.rackno = Trim(sGetVal)
                .GetText 1, sRow, sGetVal
                If Val(sGetVal) > 0 And Len(cPis902.rackno) > 0 Then
                    Call cDb.csBegin
                    sReturn = True
                    gSql = "SELECT SPCBARCD,SPCCD FROM S2PIS902 WHERE RACKNO='" & cPis902.rackno & "'"
                    With cDb.cfRecordSet(gSql)
                        If .State = adStateOpen Then
                            While (Not .EOF) And sReturn
                                cPis902.spcbarcd = "" & .Fields("SPCBARCD").Value
                                cPis902.spccd = "" & .Fields("SPCCD").Value
                                cPis902.scrapdt = Format(dtpScrapdt.Value, "yyyyMMdd")
                                cPis902.scrapreason = Trim(txtReason.Text)
                                cPis902.scrapempid = gUserId
                                
                                sReturn = cPis902.cfScraptUpdate(False)
                                
                                .MoveNext
                            Wend
                            .Close
                        End If
                    End With
                    
                    If sReturn Then
                        sReturn = cPis902.cfRackUpdate
                    End If
                    If sReturn Then
                        Call cDb.csCommit
                        spList.SetText 1, sRow, ""
                    Else
                        Call cDb.csRollback
                        Exit For
                    End If
                End If
            Next sRow
        End With
        
        If sReturn Then Call cmdFind_Click
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_Load()
Dim sCol As Integer

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    With spList
        .UserColAction = UserColActionSort
        For sCol = 1 To .MaxCols
            .ColUserSortIndicator(sCol) = ColUserSortIndicatorAscending
        Next sCol
    End With
    
    Call gsDepotCombo(cboDepot, True)
    cboDepot.ListIndex = 0
    
    Call cmdClear_Click

End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2
    grpMain.Height = Me.ScaleHeight - 50
    spList.Height = (grpMain.Height - spList.Top) - 50
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

