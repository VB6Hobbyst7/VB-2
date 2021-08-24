VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XCombobox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS913 
   BackColor       =   &H00FFFFFF&
   Caption         =   "검체폐기(검체단위)"
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
      Width           =   14820
      _ExtentX        =   26141
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
      Begin BHButton.BHImageButton cmdScrap 
         Height          =   375
         Left            =   12330
         TabIndex        =   18
         Top             =   1170
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "폐기등록"
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
         TransparentPicture=   "PIS913.frx":0000
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   9870
         TabIndex        =   17
         Top             =   1170
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
         TransparentPicture=   "PIS913.frx":17C2
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13560
         TabIndex        =   16
         Top             =   1170
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
         TransparentPicture=   "PIS913.frx":2F84
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   11100
         TabIndex        =   15
         Top             =   1170
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
         TransparentPicture=   "PIS913.frx":4746
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XTextBox.XTextBox XTextBox1 
         Height          =   495
         Left            =   90
         TabIndex        =   13
         Top             =   1620
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   873
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
         Text            =   "☞ 폐기정보"
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
         Begin TDBDate6Ctl.TDBDate dtpScrapdt 
            Height          =   315
            Left            =   1560
            TabIndex        =   23
            Top             =   90
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "PIS913.frx":5F08
            Caption         =   "PIS913.frx":5FEF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS913.frx":6052
            Keys            =   "PIS913.frx":6070
            Spin            =   "PIS913.frx":60CE
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
         Begin XLibrary_XLabel.XLabel XLabel8 
            Height          =   315
            Left            =   150
            TabIndex        =   12
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "폐기일자"
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
         Begin XLibrary_XTextBox.XTextBox txtReason 
            Height          =   315
            Left            =   4410
            TabIndex        =   11
            Top             =   90
            Width           =   8745
            _ExtentX        =   15425
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
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   3000
            TabIndex        =   10
            Top             =   90
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "폐기사유"
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
            Left            =   11550
            TabIndex        =   22
            Top             =   480
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calendar        =   "PIS913.frx":60F6
            Caption         =   "PIS913.frx":61DD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS913.frx":6240
            Keys            =   "PIS913.frx":625E
            Spin            =   "PIS913.frx":62BC
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
            Left            =   9990
            TabIndex        =   21
            Top             =   480
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calendar        =   "PIS913.frx":62E4
            Caption         =   "PIS913.frx":63CB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS913.frx":642E
            Keys            =   "PIS913.frx":644C
            Spin            =   "PIS913.frx":64AA
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
            Left            =   11550
            TabIndex        =   20
            Top             =   120
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calendar        =   "PIS913.frx":64D2
            Caption         =   "PIS913.frx":65B9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS913.frx":661C
            Keys            =   "PIS913.frx":663A
            Spin            =   "PIS913.frx":6698
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
            Left            =   9990
            TabIndex        =   19
            Top             =   120
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calendar        =   "PIS913.frx":66C0
            Caption         =   "PIS913.frx":67A7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS913.frx":680A
            Keys            =   "PIS913.frx":6828
            Spin            =   "PIS913.frx":6886
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
         Begin XLibrary_XLabel.XLabel XLabel3 
            Height          =   315
            Left            =   8580
            TabIndex        =   14
            Top             =   150
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "보관일자"
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
         Begin XLibrary_XLabel.XLabel XLabel2 
            Height          =   315
            Left            =   4380
            TabIndex        =   9
            Top             =   150
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "수검자명"
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
         Begin XLibrary_XTextBox.XTextBox txtPtnm 
            Height          =   315
            Left            =   5790
            TabIndex        =   8
            Top             =   150
            Width           =   2625
            _ExtentX        =   4630
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
         Begin XLibrary_XLabel.XLabel XLabel1 
            Height          =   315
            Left            =   4380
            TabIndex        =   7
            Top             =   510
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "검체번호"
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
         Begin XLibrary_XTextBox.XTextBox txtSpcno 
            Height          =   315
            Left            =   5790
            TabIndex        =   6
            Top             =   510
            Width           =   2625
            _ExtentX        =   4630
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
         Begin XLibrary_XComboBox.XComboBox cboSpc 
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
         Begin XLibrary_XLabel.XLabel XLabel11 
            Height          =   315
            Left            =   210
            TabIndex        =   4
            Top             =   150
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "검체종류"
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
         Begin XLibrary_XLabel.XLabel XLabel17 
            Height          =   315
            Left            =   210
            TabIndex        =   3
            Top             =   510
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "RACK NO"
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
         Begin XLibrary_XTextBox.XTextBox txtRackno 
            Height          =   315
            Left            =   1620
            TabIndex        =   2
            Top             =   510
            Width           =   2625
            _ExtentX        =   4630
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
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   8580
            TabIndex        =   1
            Top             =   510
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
      End
      Begin FPSpreadADO.fpSpread spList 
         CausesValidation=   0   'False
         Height          =   7455
         Left            =   90
         TabIndex        =   0
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
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   13
         MaxRows         =   489
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "PIS913.frx":68AE
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS913"
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
    'dtpSvFr.Value = ""
    dtpSvTo.Value = gfSystemDate
'    dtpSvTo.Value = ""
    
    dtpScrapdt.Value = gfSystemDate
    
    cboSpc.ListIndex = 0
    txtRackno.Text = ""
    txtSpcno.Text = ""
    txtPtnm.Text = ""
    txtReason.Text = ""
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdScrap, False)
    
    grpFind.Enabled = True

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdFind_Click()
Dim sRow As Long, sSpc() As String

    If cboSpc.ListIndex <= 0 Then
        MsgBox "검체종류를 선택하세요.!", vbCritical
        Exit Sub
    End If

    MousePointer = vbHourglass
    sSpc = Split(cboSpc.Text, gCboSplitStr)
    If gWorkArea Then
        gSql = "SELECT A.SPCBARCD,A.SPCCD,A.SAVEDT,A.EXPIRYDT,A.STATUS,A.DEPOTCD,A.SAVEFLOOR,A.RACKNO,A.SAVEROW,A.SAVECOL,A.SAVEEMPID               " & vbNewLine & _
               "     , B.FIELD3 AS SPCNM, C.DEPOTNM, D.EMPNM, (SELECT DISTINCT Z.PTNM FROM S2ORD101 Z WHERE SUBSTR(A.SPCBARCD,1,10)=Z.PTID) AS PTNM " & vbNewLine & _
               "  FROM S2PIS902 A                                                                                                                   " & vbNewLine & _
               "       LEFT JOIN S2LAB032 B ON A.SPCCD=B.CDVAL1 AND B.CDINDEX='C215'                                                                " & vbNewLine & _
               "       LEFT JOIN S2PIS092 C ON A.DEPOTCD=C.DEPOTCD                                                                                  " & vbNewLine & _
               "       LEFT JOIN S2COM006 D ON A.SAVEEMPID=D.EMPID                                                                                  " & vbNewLine & _
               " WHERE a.SAVEDT <= '" & Format(dtpScrapdt.Value, "yyyyMMdd") & "' AND A.SPCCD='" & Trim(sSpc(1)) & "'"
    Else
        gSql = "SELECT A.SPCBARCD,A.SPCCD,A.SAVEDT,A.EXPIRYDT,A.STATUS,A.DEPOTCD,A.SAVEFLOOR,A.RACKNO,A.SAVEROW,A.SAVECOL,A.SAVEEMPID               " & vbNewLine & _
               "     , B.SPECNAME AS SPCNM, C.DEPOTNM, D.USER_NM AS EMPNM, (SELECT DISTINCT Z.PTNM FROM S2ORD101 Z WHERE A.SPCBARCD=Z.PTID) AS PTNM " & vbNewLine & _
               "  FROM S2PIS902 A                                                                                                                   " & vbNewLine & _
               "       LEFT JOIN " & gKahpUser & "TWMED_SPEC B ON A.SPCCD=B.SPECCODE                                                                " & vbNewLine & _
               "       LEFT JOIN S2PIS092 C ON A.DEPOTCD=C.DEPOTCD                                                                                  " & vbNewLine & _
               "       LEFT JOIN " & gKahpUserTable & " D ON A.SAVEEMPID=D.USERID                                                                   " & vbNewLine & _
               " WHERE a.SAVEDT <= '" & Format(dtpScrapdt.Value, "yyyyMMdd") & "' AND A.SPCCD='" & Trim(sSpc(1)) & "'"
    End If
    
    ' 보관일자
    If Len(dtpSvFr.Value) > 0 Then
        gSql = gSql & " AND A.SAVEDT >= '" & Format(dtpSvFr.Value, "yyyyMMdd") & "'"
    End If
    If Len(dtpSvTo.Value) > 0 Then
        gSql = gSql & " AND A.SAVEDT <= '" & Format(dtpSvTo.Value, "yyyyMMdd") & "'"
    End If
    ' 폐기예정일자
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
    ' 검체번호(BARCODE)
    If Len(txtSpcno.Text) > 0 Then
        gSql = gSql & " AND A.SPCBARCD='" & Trim(txtSpcno.Text) & "'"
    End If
    ' 수검자명
    If Len(txtPtnm.Text) > 0 Then
        gSql = gSql & " AND EXISTS(SELECT Y.PTNM FROM S2ORD101 Y WHERE A.SPCBARCD=Y.PTID AND Y.PTNM LIKE '%" & Trim(txtPtnm.Text) & "%')"
    End If
    gSql = gSql & " ORDER BY A.RACKNO, A.EXPIRYDT, A.SAVEROW, A.SAVECOL"
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
                    spList.SetText 2, sRow, "" & .Fields("SPCBARCD").Value
                    spList.SetText 3, sRow, "" & .Fields("SPCCD").Value
                    spList.SetText 5, sRow, "" & .Fields("SPCNM").Value
                    spList.SetText 4, sRow, "" & .Fields("PTNM").Value
                    spList.SetText 6, sRow, Format("" & .Fields("SAVEDT").Value, "####-##-##")
                    spList.SetText 7, sRow, Format("" & .Fields("EXPIRYDT").Value, "####-##-##")
                    spList.SetText 8, sRow, gSpcStatus(Val("" & .Fields("STATUS").Value))
                    spList.SetText 9, sRow, "" & .Fields("DEPOTNM").Value
                    spList.SetText 10, sRow, "" & .Fields("SAVEFLOOR").Value
                    spList.SetText 11, sRow, "" & .Fields("RACKNO").Value
                    spList.SetText 12, sRow, "" & .Fields("SAVEROW").Value & "/" & .Fields("SAVECOL").Value
                    spList.SetText 13, sRow, "" & .Fields("EMPNM").Value
                    
                    .MoveNext
                Wend
                gPrgBar.Visible = False
                
                Call gsButtonEnable(cmdFind, False)
                Call gsButtonEnable(cmdScrap, True)
                grpFind.Enabled = False
            Else
                Call gsSpreadClear(spList, 0, True)
                MsgBox "조건에 해당하는 자료가 없습니다.!", vbCritical
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
    If MsgBox("선택하신 검체를 폐기처리하시겠습니까 ?", vbYesNo + vbQuestion) = vbYes Then
        Set cPis902 = New clsPis902
        With spList
            For sRow = 1 To .MaxRows
                .GetText 2, sRow, sGetVal:      cPis902.spcbarcd = Trim(sGetVal)
                .GetText 1, sRow, sGetVal
                If Val(sGetVal) > 0 And Len(cPis902.spcbarcd) > 0 Then
                    .GetText 3, sRow, sGetVal:      cPis902.spccd = Trim(sGetVal)
                    .GetText 11, sRow, sGetVal:     cPis902.rackno = Trim(sGetVal)
                    
                    cPis902.scrapdt = Format(dtpScrapdt.Value, "yyyyMMdd")
                    cPis902.scrapreason = Trim(txtReason.Text)
                    cPis902.scrapempid = gUserId
                    
                    Call cDb.csBegin
                    sReturn = cPis902.cfScraptUpdate
                    If sReturn = False Then
                        Call cDb.csRollback
                        Exit For
                    Else
                        Call cDb.csCommit
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
    
    Call gsSpcCombo(cboSpc, True)
    cboSpc.ListIndex = 0
    
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

Private Sub spList_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 And Col = 1 Then
        With spList
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = 1
            
            .BlockMode = True
            If Val(.CellTag) = 0 Then
                .CellTag = 1:       .Value = 1
            Else
                .CellTag = 0:       .Value = 0
            End If
            .BlockMode = False
        End With
    End If
    
End Sub

Private Sub txtRackno_LostFocus()

    txtRackno.Text = UCase(txtRackno.Text)

End Sub

Private Sub txtSpcno_LostFocus()

    txtSpcno.Text = UCase(txtSpcno.Text)


End Sub
