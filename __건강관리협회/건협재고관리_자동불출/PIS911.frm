VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XCombobox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS911 
   BackColor       =   &H00FFFFFF&
   Caption         =   "검체입고등록"
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
      Height          =   9700
      Left            =   30
      Top             =   30
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   17119
      BackColor       =   16777215
      BorderColor     =   10070188
      BorderRoundNum  =   0
      BorderStyle     =   1
      TextColor       =   0
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   720
         TabIndex        =   33
         Top             =   1860
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
         TransparentPicture=   "PIS911.frx":0000
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   1950
         TabIndex        =   32
         Top             =   1860
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
         TransparentPicture=   "PIS911.frx":17C2
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   3180
         TabIndex        =   31
         Top             =   1860
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
         TransparentPicture=   "PIS911.frx":2F84
         BackColor       =   12632319
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox5 
         Height          =   1035
         Left            =   4530
         Top             =   1200
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   1826
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
         Text            =   "Clinical Specimen"
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
         Begin TDBDate6Ctl.TDBDate dtpSavedt 
            Height          =   315
            Left            =   5460
            TabIndex        =   34
            Top             =   270
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "PIS911.frx":4746
            Caption         =   "PIS911.frx":482D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS911.frx":4890
            Keys            =   "PIS911.frx":48AE
            Spin            =   "PIS911.frx":490C
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
         Begin XLibrary_XLabel.XLabel XLabel9 
            Height          =   315
            Left            =   4170
            TabIndex        =   30
            Top             =   270
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "보관일자   :"
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
            Left            =   120
            TabIndex        =   29
            Top             =   270
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "유효기간   :"
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
         Begin XLibrary_XTextBox.XTextBox txtDay 
            Height          =   315
            Left            =   1410
            TabIndex        =   28
            Top             =   270
            Width           =   705
            _ExtentX        =   1244
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
            Mask            =   2
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   2
            Enabled         =   -1  'True
         End
         Begin XLibrary_XTextBox.XTextBox txtScrapDt 
            Height          =   315
            Left            =   2730
            TabIndex        =   27
            Top             =   270
            Width           =   1215
            _ExtentX        =   2143
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
            Text            =   "2015-06-30"
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
            TextAlign       =   2
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel3 
            Height          =   315
            Left            =   2160
            TabIndex        =   26
            Top             =   270
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "(일)"
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
         Begin XLibrary_XLabel.XLabel XLabel19 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   630
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "검체저장고 :"
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
         Begin XLibrary_XComboBox.XComboBox cboDepot 
            Height          =   315
            Left            =   1410
            TabIndex        =   24
            Top             =   630
            Width           =   2535
            _ExtentX        =   4471
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
         Begin XLibrary_XTextBox.XTextBox txtFloor 
            Height          =   315
            Left            =   5460
            TabIndex        =   23
            Top             =   630
            Width           =   915
            _ExtentX        =   1614
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
            Mask            =   2
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   2
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel5 
            Height          =   315
            Left            =   6480
            TabIndex        =   22
            Top             =   630
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "(층)"
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
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   4170
            TabIndex        =   21
            Top             =   630
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "저장고층수 :"
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
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox2 
         Height          =   585
         Left            =   90
         Top             =   9030
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1032
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
         Text            =   "검체정보"
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
         Begin XLibrary_XTextBox.XTextBox txtSpccst 
            Height          =   315
            Left            =   11100
            TabIndex        =   16
            Top             =   180
            Width           =   3375
            _ExtentX        =   5953
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
         Begin XLibrary_XTextBox.XTextBox txtSpcdt 
            Height          =   315
            Left            =   7920
            TabIndex        =   15
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
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
            Text            =   "2015-06-30"
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
         Begin XLibrary_XTextBox.XTextBox txtSpcnm 
            Height          =   315
            Left            =   4740
            TabIndex        =   14
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
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
            Text            =   "홍길동"
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
         Begin XLibrary_XLabel.XLabel XLabel18 
            Height          =   315
            Left            =   3330
            TabIndex        =   13
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16777215
            Text            =   "수검자명:"
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
         Begin XLibrary_XLabel.XLabel XLabel14 
            Height          =   315
            Left            =   6510
            TabIndex        =   12
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16777215
            Text            =   "접수일자:"
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
         Begin XLibrary_XLabel.XLabel XLabel12 
            Height          =   315
            Left            =   9690
            TabIndex        =   11
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16777215
            Text            =   "의뢰거래처:"
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
            Left            =   150
            TabIndex        =   10
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16777215
            Text            =   "검체번호:"
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
            Left            =   1560
            TabIndex        =   9
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
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
            Text            =   "123456789012"
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
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox4 
         Height          =   2145
         Left            =   11850
         Top             =   90
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   3784
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
         Text            =   "저장고정보"
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
         Begin FPSpreadADO.fpSpread spDepot 
            Height          =   1785
            Left            =   120
            TabIndex        =   20
            Top             =   270
            Width           =   2655
            _Version        =   524288
            _ExtentX        =   4683
            _ExtentY        =   3149
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            MaxCols         =   2
            OperationMode   =   1
            ScrollBars      =   2
            ShadowColor     =   16777215
            SpreadDesigner  =   "PIS911.frx":4934
            AppearanceStyle =   0
         End
      End
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox3 
         Height          =   765
         Left            =   90
         Top             =   90
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   1349
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
         Text            =   "Clinical Specimen"
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
         Begin XLibrary_XComboBox.XComboBox cboSpc 
            Height          =   315
            Left            =   1560
            TabIndex        =   18
            Top             =   270
            Width           =   2535
            _ExtentX        =   4471
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
            Left            =   150
            TabIndex        =   17
            Top             =   270
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "보관검체   :"
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
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox1 
         Height          =   1005
         Left            =   4530
         Top             =   90
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   1773
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
         Text            =   "Rack Information"
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
         Begin BHButton.BHImageButton cmdApply 
            Height          =   675
            Left            =   4080
            TabIndex        =   35
            Top             =   210
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   1191
            Caption         =   "RACK적용"
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
            TransparentPicture=   "PIS911.frx":4E91
            BackColor       =   12632319
            ImgOutLineSize  =   3
         End
         Begin XLibrary_XTextBox.XTextBox txtRackNo 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            Top             =   270
            Width           =   1875
            _ExtentX        =   3307
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
            Locked          =   -1  'True
            Mask            =   2
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   2
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   150
            TabIndex        =   7
            Top             =   270
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "Rack NO.   :"
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
         Begin XLibrary_XLabel.XLabel XLabel8 
            Height          =   315
            Left            =   3390
            TabIndex        =   6
            Top             =   600
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "(칸)"
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
         Begin XLibrary_XLabel.XLabel XLabel7 
            Height          =   315
            Left            =   2190
            TabIndex        =   5
            Top             =   600
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "(열)"
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
         Begin XLibrary_XTextBox.XTextBox txtCol 
            Height          =   315
            Left            =   2610
            TabIndex        =   4
            Top             =   600
            Width           =   705
            _ExtentX        =   1244
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
            Mask            =   2
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   2
            Enabled         =   -1  'True
         End
         Begin XLibrary_XTextBox.XTextBox txtRow 
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   600
            Width           =   705
            _ExtentX        =   1244
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
            Mask            =   2
            PromptChar      =   "_"
            WrongSound      =   0
            CustomSound     =   ""
            MaskShow        =   0   'False
            MaskColor       =   33023
            CustomMask      =   ""
            TextAlign       =   2
            Enabled         =   -1  'True
         End
         Begin XLibrary_XLabel.XLabel XLabel2 
            Height          =   315
            Left            =   150
            TabIndex        =   2
            Top             =   600
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "저장칸수   :"
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
      Begin XLibrary_XGroupBox.XGroupBox grpBarCode 
         Height          =   855
         Left            =   90
         Top             =   930
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   1508
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
         Text            =   "Barcode Read"
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
         Begin XLibrary_XLabel.XLabel XLabel1 
            Height          =   465
            Left            =   210
            TabIndex        =   1
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   820
            BackColor       =   16311512
            Text            =   "BARCODE:"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
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
         Begin XLibrary_XTextBox.XTextBox txtBarcd 
            Height          =   465
            Left            =   1590
            TabIndex        =   0
            Top             =   270
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   820
            BackColor       =   65535
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "123456789012"
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
            TextAlign       =   2
            Enabled         =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread spList 
         Height          =   6435
         Left            =   90
         TabIndex        =   19
         Top             =   2460
         Width           =   14640
         _Version        =   524288
         _ExtentX        =   25823
         _ExtentY        =   11351
         _StockProps     =   64
         ColHeaderDisplay=   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   10
         MaxRows         =   10
         Protect         =   0   'False
         ScrollBars      =   0
         SelectBlockOptions=   0
         SpreadDesigner  =   "PIS911.frx":6653
      End
   End
End
Attribute VB_Name = "PIS911"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const fSelectBackColor = vbYellow, fNoneBackColor = vbWhite

Private Sub psSpreadDisplay()
Dim sWidth As Single, sHeight As Single
Dim sRow As Integer, sCol As Integer, sCnt As Integer

    sWidth = 115:       sHeight = 260

    sRow = Val(txtRow.Text)
    sCol = Val(txtCol.Text)

    With spList
        .Redraw = False
        .MaxCols = sCol
        .MaxRows = sRow
        
        If sRow > 0 And sCol > 0 Then
            For sCnt = 1 To sRow
                .RowHeight(sCnt) = (sHeight / sRow)
            Next sCnt
            For sCnt = 1 To sCol
                .ColWidth(sCnt) = (sWidth / sCol)
            Next sCnt
            
            .RowHeight(.MaxRows) = .RowHeight(1) + (sHeight - (.RowHeight(1) * sRow))
            .ColWidth(.MaxCols) = .ColWidth(1) + (sWidth - (.ColWidth(1) * sCol))
        End If
        .Redraw = True
        .Enabled = True
        
        .Row = 1:       .Col = 1
        .Action = ActionActiveCell
        .BackColor = fSelectBackColor
    End With
    Call gsButtonEnable(cmdSave, True)
    
End Sub

Private Sub psRackInfo(ByVal brNo As String, Optional ByVal brPamTray As Boolean = False)
' Rack 마스터 유무 확인
Dim cPis091 As clsPis091

    Set cPis091 = New clsPis091
    With cPis091
        If .cfSeek(brNo) Then
            If .usefg <> "0" Then
                MsgBox "사용중지 등록된 RACK 입니다.!", vbCritical
            Else
                txtRackno.Text = brNo
            
                txtRow.Text = .rowcnt
                txtCol.Text = .colcnt
                
                txtRow.BackColor = gLockColor
                txtRow.Enabled = False
                txtCol.BackColor = gLockColor
                txtCol.Enabled = False
                
                Call psSpreadDisplay
            End If
        Else
            If brPamTray Then
                txtRackno.Text = brNo
            
                txtRow.Text = 10
                txtCol.Text = 5
                
                txtRow.BackColor = gLockColor
                txtRow.Enabled = False
                txtCol.BackColor = gLockColor
                txtCol.Enabled = False
            Else
                Call gsButtonEnable(cmdApply, True)
            
                txtRackno.Text = brNo
                
                txtRow.Text = 10
                txtCol.Text = 5
                
                txtRow.BackColor = gEditColor
                txtRow.Enabled = True
                txtCol.BackColor = gEditColor
                txtCol.Enabled = True
                
                Me.KeyPreview = False
                txtRow.SetFocus
            End If
        End If
    End With
    
End Sub

Private Function pfRackExistsCheck(ByVal brNo As String) As Boolean
' Rack 저장유무 확인
Dim sReturn As Boolean, sRow As Integer, sCol As Integer
Dim sDepot() As String, sSpc() As String

    gSql = "SELECT A.*, B.ROWCNT,B.COLCNT FROM S2PIS901 A LEFT JOIN S2PIS091 B ON A.RACKNO=B.RACKNO WHERE A.RACKNO='" & brNo & "'"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                sReturn = True
                
                txtRackno.Text = brNo
                txtRow.Text = Val("" & .Fields("ROWCNT").Value)
                txtCol.Text = Val("" & .Fields("COLCNT").Value)
                
                Call psSpreadDisplay
                
                dtpSavedt.Value = Format("" & .Fields("SAVEDT").Value, "####-##-##")
                txtScrapdt.Text = Format("" & .Fields("EXPIRYDT").Value, "####-##-##")
                txtFloor.Text = Val("" & .Fields("SAVEFLOOR").Value)
                txtDay.Text = DateDiff("d", dtpSavedt.Value, txtScrapdt.Text)
                
                ' 저장고 콤보선택
                For sCol = 0 To cboDepot.ListCount
                    sDepot = Split(cboDepot.List(sCol), gCboSplitStr)
                    If .Fields("DEPOTCD").Value = Trim(sDepot(1)) Then
                        cboDepot.ListIndex = sCol
                        Call cboDepot_Click
                        Exit For
                    End If
                Next sCol
                
                If gWorkArea Then
                    gSql = "SELECT A.SPCBARCD, A.SPCCD, B.FIELD3 AS SPCNM, A.SAVEROW, A.SAVECOL, " & vbNewLine & _
                           "       (SELECT DISTINCT C.PTNM FROM S2ORD101 C WHERE SUBSTR(A.SPCBARCD,1,10)=C.PTID) AS PTNM " & vbNewLine & _
                           "  FROM S2PIS902 A LEFT JOIN S2LAB032 B ON A.SPCCD=B.CDVAL1 AND B.CDINDEX='C215'" & vbNewLine & _
                           " WHERE A.RACKNO='" & brNo & "' AND A.STATUS='0'"
                Else
                    gSql = "SELECT A.SPCBARCD, A.SPCCD, B.SPECNAME AS SPCNM, A.SAVEROW, A.SAVECOL, " & vbNewLine & _
                           "       (SELECT DISTINCT C.PTNM FROM S2ORD101 C WHERE A.SPCBARCD=C.PTID) AS PTNM " & vbNewLine & _
                           "  FROM S2PIS902 A LEFT JOIN " & gKahpUser & "TWMED_SPEC B ON A.SPCCD=B.SPECCODE" & vbNewLine & _
                           " WHERE A.RACKNO='" & brNo & "' AND A.STATUS='0'"
                End If
                With cDb.cfRecordSet(gSql)
                    If .State = adStateOpen Then
                        If Not .EOF Then
                            ' 검체종류선택
                            For sCol = 0 To cboSpc.ListCount
                                sSpc = Split(cboSpc.List(sCol), gCboSplitStr)
                                If .Fields("SPCCD").Value = Trim(sSpc(1)) Then
                                    cboSpc.ListIndex = sCol
                                    Exit For
                                End If
                            Next sCol
                            
                            While (Not .EOF)
                                sRow = Val("" & .Fields("SAVEROW").Value)
                                sCol = Val("" & .Fields("SAVECOL").Value)
                                
                                spList.SetText sCol, sRow, .Fields("SPCBARCD").Value & vbNewLine & .Fields("PTNM").Value & vbNewLine & .Fields("SPCNM").Value
                                .MoveNext
                            Wend
                        End If
                        .Close
                    End If
                End With
            End If
            .Close
        End If
    End With
    
    pfRackExistsCheck = sReturn

End Function

Private Sub cboDepot_Click()
Dim sDepot() As String, cPis092 As clsPis092, sRow As Integer

    If cboDepot.ListIndex >= 0 Then
        sDepot = Split(cboDepot.Text, gCboSplitStr)
    
        Set cPis092 = New clsPis092
        With cPis092
            If .cfSeek(Trim(sDepot(1))) Then
                Call gsSpreadClear(spDepot, .floor, False)
                
                spDepot.Row = 1:        spDepot.Row2 = spDepot.MaxRows
                spDepot.Col = 1:        spDepot.Col2 = 1
                spDepot.BlockMode = True
                spDepot.Text = .rackcnt
                spDepot.BlockMode = False
                
                gSql = "SELECT SAVEFLOOR, COUNT(*) AS CNT FROM S2PIS901 " & vbNewLine & _
                       " WHERE DEPOTCD='" & Trim(sDepot(1)) & "' GROUP BY SAVEFLOOR"
                With cDb.cfRecordSet(gSql)
                    If .State = adStateOpen Then
                        While (Not .EOF)
                            sRow = Val("" & .Fields("SAVEFLOOR").Value)
                            spDepot.SetText 2, sRow, Val("" & .Fields("CNT").Value)
                            
                            .MoveNext
                        Wend
                        .Close
                    End If
                End With
            End If
        End With
    End If

End Sub

Private Sub cboSpc_Click()
Dim sSpc() As String

    If cboSpc.ListIndex >= 0 Then
        sSpc = Split(cboSpc.Text, gCboSplitStr)
        txtDay.Text = Val(sSpc(2))
    End If

End Sub

Private Sub cmdApply_Click()
Dim cPis091 As clsPis091
    
    Call psSpreadDisplay
    
    Set cPis091 = New clsPis091
    With cPis091
        .rackno = Trim(txtRackno.Text)
        .rowcnt = Val(txtRow.Text)
        .colcnt = Val(txtCol.Text)
        .usefg = "0"
        .empid = gUserId
        If .cfSave Then
            txtRow.BackColor = gLockColor
            txtRow.Enabled = False
            txtCol.BackColor = gLockColor
            txtCol.Enabled = False
        
            Call gsButtonEnable(cmdApply, False)
        End If
    End With

End Sub

Private Sub cmdClear_Click()
Dim sCtl As Control

    For Each sCtl In Me.Controls
        If TypeOf sCtl Is XTextBox Then
            sCtl.Text = ""
        ElseIf TypeOf sCtl Is XComboBox Then
            sCtl.ListIndex = -1
        End If
    Next
    
    With spList
        .Row = 1:               .Col = 1
        .Row2 = .MaxRows:       .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .CellTag = ""
        .BackColor = fNoneBackColor
        .BlockMode = False
    End With
    
    Call gsSpreadClear(spDepot, 0, False)
    
    txtRow.BackColor = gLockColor
    txtRow.Enabled = False
    txtCol.BackColor = gLockColor
    txtCol.Enabled = False
    
    dtpSavedt.Value = gfSystemDate
    
    Call gsButtonEnable(cmdApply, False)
    Call gsButtonEnable(cmdSave, False)
    
    cboSpc.ListIndex = 0
    Call cboSpc_Click
    
    spList.Enabled = False
    txtBarcd.Enabled = True
    txtBarcd.SetFocus
   
End Sub

Private Sub cmdClose_Click()

    If spList.Enabled Then
        If MsgBox("검체등록중입니다. 종료하시겠습니까 ?", vbYesNo + vbQuestion) <> vbYes Then
            Exit Sub
        End If
    End If

    Unload Me

End Sub

Private Sub cmdSave_Click()
Dim cPis902 As clsPis902, cPis901 As clsPis901
Dim sRow As Integer, sCol As Integer, sDepot() As String, sSpc() As String
Dim sReturn As Boolean

    If cboSpc.ListIndex < 0 Then
        MsgBox "검체종류를 선택하세요.!", vbCritical
        cboSpc.SetFocus
        Exit Sub
    End If
    If cboDepot.ListIndex < 0 Then
        MsgBox "저장고와 저장층수를 입력하세요.!", vbCritical
        cboDepot.SetFocus
        Exit Sub
    End If

    MousePointer = vbHourglass
    sReturn = True
    
    Call cDb.csBegin
    ' RACK 정보저장
    sDepot = Split(cboDepot.Text, gCboSplitStr)
    sSpc = Split(cboSpc.Text, gCboSplitStr)
    
    Set cPis901 = New clsPis901
    With cPis901
        .rackno = Trim(txtRackno.Text)
        .savedt = Format(dtpSavedt.Value, "yyyyMMdd")
        .expirydt = Format(txtScrapdt.Text, "yyyyMMdd")
        .depotcd = Trim(sDepot(1))
        .savefloor = Val(txtFloor.Text)
        .empid = gUserId
        
        sReturn = .cfSave
    End With
    
    If sReturn = False Then GoTo errSave
    
    ' 검체저장
    Set cPis902 = New clsPis902
    With cPis902
        .savedt = Format(dtpSavedt.Value, "yyyyMMdd")
        .expirydt = Format(txtScrapdt.Text, "yyyyMMdd")
        .depotcd = Trim(sDepot(1))
        .savefloor = Val(txtFloor.Text)
        .rackno = Trim(txtRackno.Text)
        .saveempid = gUserId
        .spccd = Trim(sSpc(1))
        .status = "0"
        
        For sRow = 1 To spList.MaxRows
            For sCol = 1 To spList.MaxCols
                spList.Row = sRow
                spList.Col = sCol
                
                If Len(spList.CellTag) > 0 Then
                    .saverow = sRow
                    .savecol = sCol
                    .spcbarcd = Trim(spList.CellTag)
                    
                    sReturn = .cfSave
                End If
                If sReturn = False Then GoTo errSave
            Next sCol
            If sReturn = False Then GoTo errSave
        Next sRow
    End With
    
    Call cDb.csCommit
    Call cboDepot_Click
    MsgBox "검체저장이 완료되었습니다.!", vbInformation
    
    Call cmdClear_Click
    MousePointer = vbDefault
    Exit Sub
    
errSave:
    Call cDb.csRollback
    MousePointer = vbDefault
End Sub

Private Sub dtpSavedt_Change()

    Call txtDay_Change

End Sub

Private Sub Form_Activate()

    frmMain.lblTitle.Text = Me.Caption

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'    If Me.ActiveControl.Name <> txtBarcd.Name Then
'        txtBarcd.SetFocus
'        txtBarcd.Text = txtBarcd.Text & Chr(KeyAscii)
'        SendKeys "{End}"
'    End If
    
End Sub

Private Sub Form_Load()

    grpMain.BorderColor = gGrpLineColor
    grpMain.BackColor = gGrpBackColor

    Me.Show
    
    Call gsSpcCombo(cboSpc)
    cboSpc.ListIndex = 0
    
    Call gsDepotCombo(cboDepot)
    
    spList.SetText SpreadHeader, SpreadHeader, "열/칸"
    spDepot.SetText SpreadHeader, SpreadHeader, "층수"
    
    Call cmdClear_Click
    
    Me.KeyPreview = True

End Sub

Private Sub Form_Resize()
On Error Resume Next

    grpMain.Top = (Me.ScaleHeight - grpMain.Height) / 2
    grpMain.Left = (Me.ScaleWidth - grpMain.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

Private Sub spList_Click(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 And Col > 0 Then
        With spList
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = fNoneBackColor
            .BlockMode = False
            
            .Row = Row:          .Col = Col
            .Action = ActionActiveCell
            .BackColor = fSelectBackColor
        End With
    End If

End Sub

Private Sub txtBarcd_Change()

    txtBarcd.Text = UCase(txtBarcd.Text)
    SendKeys "{End}"

End Sub

Private Sub txtBarcd_KeyPress(KeyAscii As Integer)
Dim cSpc As clsSpcInfo

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(txtBarcd.Text) Then
            Call pfPamTray(Trim(txtBarcd.Text))
        ElseIf Left(txtBarcd.Text, 2) = "RK" Then
            Call psRackInfo(txtBarcd.Text)
            
            If pfRackExistsCheck(txtBarcd.Text) Then
                MsgBox "이미 저장 등록된 RACK 입니다.!", vbCritical
            End If
        Else
            If cmdSave.Enabled = False Then
                MsgBox "RACK 정보를 먼저 입력하세요.!", vbCritical
            Else
                ' 기존등록 검체확인
                If pfSpcExistsCheck(Trim(txtBarcd.Text)) = False Then
                    Set cSpc = New clsSpcInfo
                    With cSpc
                        If .cfSpcInfo(Mid(txtBarcd.Text, 1, 10)) Then
                            .spcno = Trim(txtBarcd.Text)
                            
                            txtSpcno.Text = .spcno
                            txtSpcnm.Text = .ptnm
                            txtSpcdt.Text = .rcvdt
                            txtSpccst.Text = .hosnm
                            
                            spList.Row = spList.ActiveRow
                            spList.Col = spList.ActiveCol
                            spList.Text = .spcno & vbNewLine & .ptnm & vbNewLine & .hosnm
                            
                            spList.CellTag = .spcno
                            
                            Call psSpreadMove(spList.ActiveRow, spList.ActiveCol + 1)
                        Else
                            MsgBox "접수되지 않은 검체입니다.!", vbCritical
                        End If
                    End With
                End If
            End If
            txtBarcd.SetFocus
        End If
        
        txtBarcd.Text = ""
    End If

End Sub

Private Function pfPamTray(ByVal brNo As String) As Boolean
Dim cSpc As clsSpcInfo, sReturn As Boolean
Dim sRow As Integer, sCol As Integer, sSpcNo As String

    On Error GoTo errPamTray
    sReturn = Not pfRackExistsCheck(brNo)
    
    If sReturn Then
        gSql = "SELECT * FROM PAM_RACKINFO WHERE SYS_DT='" & Format(dtpSavedt.Value, "yyyyMMdd") & "' AND MODULE_NM='OBS'" & vbNewLine & _
               "   AND RACK_NO='" & Trim(txtBarcd.Text) & "' AND USE_FG='Y'" & vbNewLine & _
               " ORDER BY RACK_POS"
        With cDb.cfRecordSet(gSql)
            If .State = adStateOpen Then
                If Not .EOF Then
                    sSpcNo = Trim(txtBarcd.Text)
                    Call cmdClear_Click
                    txtBarcd.Text = sSpcNo
                    
                    Call psRackInfo(txtBarcd.Text, True)
                    Call cmdApply_Click
                    
                    Set cSpc = New clsSpcInfo
                    
                    spList.Redraw = False
                    
                    While (Not .EOF)
                        sRow = Fix(Val("" & .Fields("RACK_POS").Value) / 10) + 1
                        sCol = (Val("" & .Fields("RACK_POS").Value) Mod 10) + 1
                        
                        sSpcNo = "" & .Fields("SPC_NO").Value
                        
                        ' 기존등록 검체확인
                        If pfSpcExistsCheck(sSpcNo) = False Then
                            With cSpc
                                If .cfSpcInfo(Mid(sSpcNo, 1, 10)) Then
                                    txtSpcno.Text = sSpcNo
                                    txtSpcnm.Text = .ptnm
                                    txtSpcdt.Text = .rcvdt
                                    txtSpccst.Text = .hosnm
                                    
                                    spList.Row = spList.ActiveRow
                                    spList.Col = spList.ActiveCol
                                    spList.Text = sSpcNo & vbNewLine & .ptnm & vbNewLine & .hosnm
                                    
                                    spList.CellTag = sSpcNo
                                    
                                    Call psSpreadMove(spList.ActiveRow, spList.ActiveCol + 1)
                                End If
                            End With
                        End If
                        
                        .MoveNext
                    Wend
                    spList.Redraw = True
                    txtBarcd.Enabled = False
                Else
                    MsgBox "PAM에서 등록된 RACK정보가 없습니다.!", vbCritical
                    txtBarcd.SetFocus
                End If
                .Close
            End If
        End With
    Else
        txtBarcd.Enabled = False
    End If
    pfPamTray = sReturn
    Exit Function
    
errPamTray:
    MsgBox Err.Description, vbCritical
    pfPamTray = False

End Function

Private Function pfSpcExistsCheck(ByVal brNo As String) As Boolean
Dim sReturn As Boolean, sRow As Integer, sCol As Integer
Dim cPis902 As clsPis902, sSpc() As String, sMsg As String

    With spList
        For sRow = 1 To .MaxRows
            For sCol = 1 To .MaxCols
                .Row = sRow
                .Col = sCol
                If brNo = .CellTag Then
                    sReturn = True
                    Exit For
                End If
            Next sCol
            If sReturn = True Then Exit For
        Next sRow
    End With
    
    If sReturn Then
        MsgBox "이미 등록된 검체입니다.!" & vbNewLine & "[" & brNo & "] (" & sRow & " 열 / " & sCol & " 칸)", vbCritical
    Else
        sSpc = Split(cboSpc.Text, gCboSplitStr)
        Set cPis902 = New clsPis902
        With cPis902
            If .cfSeek(brNo, Trim(sSpc(1))) Then
                sReturn = True
                
                Select Case .status
                    Case "0"
                            MsgBox Format(.savedt, "####년##월##일") & "에 Rack NO [" & .rackno & "]에 등록된 검체입니다.!", vbCritical
                    Case "1"
                            MsgBox Format(.rentdt, "####년##월##일") & "에 대출된 검체입니다.!", vbCritical
                    Case "2"
                            MsgBox Format(.scrapdt, "####년##월##일") & "에 폐기된 검체입니다.!", vbCritical
                End Select
            End If
        End With
    End If
    pfSpcExistsCheck = sReturn

End Function

Private Sub psSpreadMove(ByVal brRow As Long, ByVal brCol As Long)
Dim sRow As Integer, sCol As Integer, sGetVal As Variant

    With spList
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
        .BackColor = fNoneBackColor
        .BlockMode = False
    
        sRow = brRow:    sCol = brCol
moveAgain:
        If sCol > .MaxCols Then
            sCol = 1:   sRow = sRow + 1
        End If
        
        If sRow > .MaxRows Then
            frmMain.stsBar.Panels(2).Text = "RACK 장착이 완료되었습니다.!"
            Exit Sub
        End If
        
        .GetText sCol, sRow, sGetVal
        If Len(sGetVal) > 0 Then
            sCol = sCol + 1
            GoTo moveAgain
        End If
        
        .Row = sRow:          .Col = sCol
        .Action = ActionActiveCell
        .BackColor = fSelectBackColor
    End With
    
End Sub

Private Sub txtDay_Change()

    txtScrapdt.Text = Format(DateAdd("d", Val(txtDay.Text), dtpSavedt.Value), "yyyy-MM-dd")

End Sub

Private Sub txtDay_GotFocus()

    Me.KeyPreview = False

End Sub

Private Sub txtDay_LostFocus()

    Me.KeyPreview = True

End Sub

Private Sub txtFloor_GotFocus()

    Me.KeyPreview = False
    
End Sub

Private Sub txtFloor_LostFocus()

    Me.KeyPreview = True
    
End Sub
