VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{B88C4DC8-B707-435E-8B13-08058839823E}#2.0#0"; "XLabel.ocx"
Object = "{94265C54-C72D-40E9-95C4-FBB6BF532C26}#2.0#0"; "XGroupBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XCombobox.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form PIS209 
   BackColor       =   &H00FFFFFF&
   Caption         =   "½Ã¾à À¯È¿±âÇÑ º¯°æ"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
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
   WindowState     =   2  'ÃÖ´ëÈ­
   Begin XLibrary_XGroupBox.XGroupBox grpChulgo 
      Height          =   3285
      Left            =   4530
      Top             =   3480
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
         Name            =   "±¼¸²"
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
      Begin TDBDate6Ctl.TDBDate dtpDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Top             =   1620
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   556
         Calendar        =   "PIS209.frx":0000
         Caption         =   "PIS209.frx":00E7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PIS209.frx":014A
         Keys            =   "PIS209.frx":0168
         Spin            =   "PIS209.frx":01C6
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
      Begin BHButton.BHImageButton cmdSave 
         Height          =   375
         Left            =   3510
         TabIndex        =   26
         Top             =   2790
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Àú Àå"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS209.frx":01EE
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdCancel 
         Height          =   375
         Left            =   4770
         TabIndex        =   30
         Top             =   2790
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Ãë ¼Ò"
         CaptionChecked  =   "BHImageButton3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS209.frx":19B0
         BackColor       =   14737632
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XComboBox.XComboBox cboReason 
         Height          =   315
         Left            =   1320
         TabIndex        =   23
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
            Name            =   "±¼¸²Ã¼"
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
      Begin XLibrary_XLabel.XLabel XLabel12 
         Height          =   315
         Left            =   150
         TabIndex        =   22
         Top             =   2340
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "ºñ°í»çÇ×"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         TabIndex        =   21
         Top             =   2340
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         BackColor       =   16777215
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
      Begin XLibrary_XLabel.XLabel XLabel11 
         Height          =   315
         Left            =   150
         TabIndex        =   20
         Top             =   1980
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "º¯°æ»çÀ¯"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         TabIndex        =   19
         Top             =   1620
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "º¯°æÀÏÀÚ"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         Left            =   3120
         TabIndex        =   18
         Top             =   1260
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "Àç°í·®"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         TabIndex        =   16
         Top             =   900
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "ÀÔ°í·®"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   900
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
      Begin XLibrary_XLabel.XLabel XLabel5 
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   1260
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "À¯È¿±âÇÑ"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         TabIndex        =   12
         Top             =   900
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "LOT NO"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
      Begin XLibrary_XTextBox.XTextBox txtLot 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   900
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
      Begin XLibrary_XTextBox.XTextBox txtNm 
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   540
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
      Begin XLibrary_XTextBox.XTextBox XTextBox2 
         Height          =   405
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   714
         BackColor       =   16249839
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "À¯È¿±âÇÑ º¯°æ"
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
      Begin XLibrary_XLabel.XLabel XLabel2 
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   540
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColor       =   16311512
         Text            =   "Ç°¸ñÄÚµå"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
         Left            =   1320
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   540
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
      Begin XLibrary_XTextBox.XTextBox txtEntdt 
         Height          =   315
         Left            =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2790
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
      Begin XLibrary_XTextBox.XTextBox txtEntseq 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2790
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BackColor       =   14737632
         BorderColor     =   16744576
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
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
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00808080&
      BorderStyle     =   0  '¾øÀ½
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   4650
      ScaleHeight     =   3285
      ScaleWidth      =   6180
      TabIndex        =   24
      Top             =   3630
      Visible         =   0   'False
      Width           =   6180
   End
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
         Name            =   "±¼¸²"
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
      Begin BHButton.BHImageButton cmdChange 
         Height          =   375
         Left            =   12300
         TabIndex        =   29
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "±âÇÑº¯°æ"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS209.frx":3172
         BackColor       =   16777215
         AlphaColor      =   16711680
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdFind 
         Height          =   375
         Left            =   11070
         TabIndex        =   28
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Á¶ È¸"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS209.frx":4934
         BackColor       =   14737632
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClear 
         Height          =   375
         Left            =   9840
         TabIndex        =   27
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "È­¸éÁö¿ò"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS209.frx":60F6
         BackColor       =   16777215
         AlphaColor      =   16711680
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdClose 
         Height          =   375
         Left            =   13530
         TabIndex        =   25
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "´Ý ±â"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentPicture=   "PIS209.frx":78B8
         BackColor       =   14737632
         ImgOutLineSize  =   3
      End
      Begin XLibrary_XGroupBox.XGroupBox XGroupBox3 
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
            Name            =   "±¼¸²"
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
            Left            =   2880
            TabIndex        =   33
            Top             =   180
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "PIS209.frx":907A
            Caption         =   "PIS209.frx":9161
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS209.frx":91C4
            Keys            =   "PIS209.frx":91E2
            Spin            =   "PIS209.frx":9240
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
            Left            =   1470
            TabIndex        =   32
            Top             =   180
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   556
            Calendar        =   "PIS209.frx":9268
            Caption         =   "PIS209.frx":934F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "PIS209.frx":93B2
            Keys            =   "PIS209.frx":93D0
            Spin            =   "PIS209.frx":942E
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
            Left            =   6900
            TabIndex        =   31
            Top             =   180
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            Caption         =   "..."
            CaptionChecked  =   "BHImageButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TransparentPicture=   "PIS209.frx":9456
            BackColor       =   14737632
            AlphaColor      =   16777215
            ImgOutLineSize  =   3
         End
         Begin XLibrary_XTextBox.XTextBox txtStkcd 
            Height          =   315
            Left            =   5820
            TabIndex        =   4
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            BackColor       =   16777215
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
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
         Begin XLibrary_XLabel.XLabel XLabel4 
            Height          =   315
            Left            =   4410
            TabIndex        =   3
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "ÀÔ°íÇ°¸ñ"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
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
         Begin XLibrary_XTextBox.XTextBox txtStkNm 
            Height          =   315
            Left            =   7260
            TabIndex        =   2
            Top             =   180
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   556
            BackColor       =   14737632
            BorderColor     =   16744576
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²"
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
         Begin XLibrary_XLabel.XLabel XLabel6 
            Height          =   315
            Left            =   210
            TabIndex        =   1
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   16311512
            Text            =   "ÀÔ°í±â°£"
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
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
            Name            =   "±¼¸²Ã¼"
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
         SpreadDesigner  =   "PIS209.frx":AC18
         VisibleCols     =   3
         VisibleRows     =   10
         CellNoteIndicatorColor=   16576
      End
   End
End
Attribute VB_Name = "PIS209"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    grpMain.Enabled = True
    picShadow.Visible = False
    grpChulgo.Visible = False
    
End Sub

Private Sub cmdChange_Click()
Dim sGetVal As Variant

    With spList
        If .ActiveCol > 0 And .ActiveRow > 0 Then
            .GetText 1, .ActiveRow, sGetVal:        txtEntdt.Text = Trim(sGetVal)
            .GetText 4, .ActiveRow, sGetVal:        txtCd.Text = Trim(sGetVal)
            .GetText 5, .ActiveRow, sGetVal:        txtNm.Text = Trim(sGetVal)
            .GetText 6, .ActiveRow, sGetVal:        txtLot.Text = Trim(sGetVal)
            .GetText 8, .ActiveRow, sGetVal:        txtExpirydt.Text = Trim(sGetVal)
            .GetText 9, .ActiveRow, sGetVal:        txtEntqty.Text = Trim(sGetVal):     txtEntqty.Tag = Val(Str(txtEntqty.Text))
            .GetText 10, .ActiveRow, sGetVal:       txtUseqty.Text = Trim(sGetVal):     txtUseqty.Tag = Val(Str(txtUseqty.Text))
            .GetText 11, .ActiveRow, sGetVal:       txtEntseq.Text = Val(sGetVal)
            
            dtpDate.Value = gfSystemDate
            
            grpMain.Enabled = False
            
            Call gsReasonCombo(cboReason, gReasonExpirydt)
            
            cboReason.ListIndex = -1
            picShadow.Visible = True:       picShadow.ZOrder 0
            grpChulgo.Visible = True:       grpChulgo.ZOrder 0
        End If
    End With
    
End Sub

Private Sub cmdClear_Click()

    dtpFrdt.Value = Format(gfSystemDate, "yyyy-MM") & "-01"
    dtpTodt.Value = gfSystemDate
    
    Call gsSpreadClear(spList, 0, True)
    
    Call gsButtonEnable(cmdFind, True)
    Call gsButtonEnable(cmdChange, False)
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
        
End Sub

Private Sub cmdFind_Click()
Dim sRow As Long, sFrDt As String, sToDt As String

    MousePointer = vbHourglass
    sFrDt = Format(dtpFrdt.Value, "yyyyMMdd")
    sToDt = Format(dtpTodt.Value, "yyyyMMdd")
    
    gSql = "SELECT A.*, X.NM_ITEM AS STKNM, C.CSTNM FROM S2PIS201 A                             " & vbNewLine & _
           "       LEFT JOIN " & gTBLstk & " X ON A.STKCD=X.CD_ITEM " & gERPStkCondition & "    " & vbNewLine & _
           "       LEFT JOIN S2PIS002 C ON A.CSTCD=C.CSTCD                                      " & vbNewLine & _
           " WHERE A.ENTDT BETWEEN '" & sFrDt & "' AND '" & sToDt & "'                          " & vbNewLine & _
           "   AND (A.OQTY_SO IS NULL OR A.OQTY_SO < A.IQTY_SO)                                 " & vbNewLine
    If Len(txtStkcd.Text) > 0 Then
        gSql = gSql & "  AND A.STKCD='" & Trim(txtStkcd.Text) & "'                              " & vbNewLine
    End If
    gSql = gSql & " ORDER BY A.ENTDT, A.ENTSEQ"
    With cDb.cfRecordSet(gSql)
        If .State = adStateOpen Then
            If Not .EOF Then
                gPrgBar.Max = .RecordCount:     gPrgBar.Value = 0
                gPrgBar.Visible = True:         gPrgBar.Refresh
                Call gsSpreadClear(spList, .RecordCount, True)
                While (Not .EOF)
                    sRow = sRow + 1

                    spList.SetText 1, sRow, Format("" & .Fields("ENTDT").Value, "####-##-##")
                    spList.SetText 2, sRow, "" & .Fields("CSTCD").Value
                    spList.SetText 3, sRow, "" & .Fields("CSTNM").Value
                    spList.SetText 4, sRow, "" & .Fields("STKCD").Value
                    spList.SetText 5, sRow, "" & .Fields("STKNM").Value
                    spList.SetText 6, sRow, "" & .Fields("LOTNO").Value
                    spList.SetText 7, sRow, Format("" & .Fields("ORG_EXPIRYDT").Value, "####-##-##")
                    spList.SetText 8, sRow, Format("" & .Fields("EXPIRYDT").Value, "####-##-##")
                    spList.SetText 9, sRow, gfQtyOutputStr(Val("" & .Fields("IQTY_SO").Value))
                    spList.SetText 10, sRow, gfQtyOutputStr(Val("" & .Fields("IQTY_SO").Value) - Val("" & .Fields("OQTY_SO").Value))
                    spList.SetText 11, sRow, "" & .Fields("ENTSEQ").Value
                    
                    .MoveNext
                Wend
                gPrgBar.Visible = False
                
                Call gsButtonEnable(cmdFind, False)
                Call gsButtonEnable(cmdChange, True)
            Else
                Call gsSpreadClear(spList, 0, True)
                MsgBox "ÇØ´ç±â°£¿¡ ÀÔ°íÀÚ·á°¡ ¾ø½À´Ï´Ù.!", vbCritical
            End If
            .Close
        End If
    End With
    MousePointer = vbDefault
                    
End Sub

Private Sub cmdSave_Click()
Dim cPis202 As clsPis202, sReason() As String

    MousePointer = vbHourglass
    If cboReason.ListIndex < 0 Then
        MsgBox "º¯°æ»çÀ¯¸¦ ¼±ÅÃÇÏ¼¼¿ä.!", vbCritical
        Exit Sub
    End If
    
    Set cPis202 = New clsPis202
    With cPis202
        .entdt = Format(txtEntdt.Text, "yyyyMMdd")
        .entseq = Val(txtEntseq.Text)
        .seq = 0
        .prevdt = Format(txtExpirydt.Text, "yyyyMMdd")
        .changedt = Format(dtpDate.Value, "yyyyMMdd")
        If cboReason.ListIndex >= 0 Then
            sReason = Split(cboReason.Text, gCboSplitStr)
            .reasoncd = Trim(sReason(1))
        End If
        .remark = Trim(txtRemark.Text)
        .empid = gUserId
        
        Call cDb.csBegin
        If .cfSave Then
            Call cDb.csCommit
            
            Call cmdCancel_Click
            Call cmdFind_Click
        Else
            Call cDb.csRollback
        End If
    End With
    MousePointer = vbDefault
    
End Sub

Private Sub cmdStk_Click()

    hlpStkList.Tag = "one"
    hlpStkList.Show vbModal
    
    If Len(gHelpCode) > 0 Then
        txtStkcd.Text = gHelpCode
        txtStkNm.Text = gfStkName(gHelpCode)
    Else
        txtStkcd.Text = ""
        txtStkNm.Text = ""
    End If
    
End Sub

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
    spList.Height = (grpMain.Height - spList.Top) - 50
    
    grpChulgo.Left = (Me.ScaleWidth - grpChulgo.Width) / 2
    grpChulgo.Top = (Me.ScaleHeight - grpChulgo.Height) / 2
    
    picShadow.Left = grpChulgo.Left + 80
    picShadow.Top = grpChulgo.Top + 80
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.lblTitle.Text = ""
    
End Sub

Private Sub txtStkcd_LostFocus()

    txtStkNm.Text = gfStkName(Trim(txtStkcd.Text))
    If Len(txtStkNm.Text) = 0 Then txtStkcd.Text = ""

End Sub
