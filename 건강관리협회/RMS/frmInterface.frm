VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{BD146989-F30B-4134-B202-680CC90638EF}#2.0#0"; "XTextBox.ocx"
Object = "{1A3A9E7F-34C1-4F5C-BD80-63FA100EC4A0}#2.0#0"; "XComboBox.ocx"
Object = "{38B18A4D-67F2-4F9B-B495-7ABA033953BB}#2.0#0"; "XProgressBar.ocx"
Object = "{3B930683-5AF1-4F07-9CE8-CA8063E1F3DD}#2.0#0"; "XButton.ocx"
Begin VB.Form frmInterface 
   BackColor       =   &H00FFFFFF&
   Caption         =   " RMS [Relation Management System] "
   ClientHeight    =   13095
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   18285
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   13095
   ScaleWidth      =   18285
   Begin XLibrary_XProgressBar.XProgress XProgress1 
      Height          =   345
      Left            =   2220
      TabIndex        =   60
      Top             =   12060
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   609
      BackColor       =   16777215
      BorderColor     =   14737632
      BorderWidth     =   3
      ProgressColor1  =   8454016
      ProgressColor2  =   8454016
      ProgressStyle   =   0
      Min             =   0
      Max             =   100
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextFontColor   =   0
      TextFontBackColor=   0
      TextFontBackStyle=   1
      TextAlign       =   2
      TextAlignMargin =   0
      GradientStyle   =   4
      GradientPosition=   0
      BevelStyle      =   0
      BevelHeight     =   1
      PictureStyle    =   0
      BoxWidth        =   6
      BoxWidthMargin  =   1
      BoxHeightMargin =   1
      Text            =   ""
      BorderStyle     =   2
      MouseCursor     =   0
      Enabled         =   -1  'True
      rImgWidth       =   0
      rImgHeight      =   0
   End
   Begin Threed.SSPanel SSPanel9 
      Height          =   885
      Left            =   180
      TabIndex        =   48
      Top             =   30
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   1561
      _Version        =   262144
      BackColor       =   16777215
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSRibbon ssMenu 
         Height          =   495
         Index           =   0
         Left            =   180
         TabIndex        =   49
         Top             =   210
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   262144
         BackColor       =   -2147483629
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "  �߾Ӱ˻缾�� �˻�����"
         PictureAlignment=   1
         Value           =   -1  'True
      End
      Begin Threed.SSRibbon ssMenu 
         Height          =   495
         Index           =   1
         Left            =   3030
         TabIndex        =   50
         Top             =   210
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   262144
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�߾Ӱ˻缾�� �˻���"
      End
      Begin Threed.SSRibbon ssMenu 
         Height          =   495
         Index           =   2
         Left            =   5850
         TabIndex        =   51
         Top             =   210
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   262144
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�߾Ӱ˻缾�� ó������"
      End
      Begin Threed.SSRibbon ssMenu 
         Height          =   495
         Index           =   3
         Left            =   8670
         TabIndex        =   52
         Top             =   210
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   262144
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�߾Ӱ˻缾�� �˻����"
      End
      Begin XLibrary_XButton.XButton cmdClose 
         Height          =   435
         Left            =   15750
         TabIndex        =   104
         Top             =   240
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         BackColor1      =   16777215
         BackColor2      =   16777215
         BackColorEx     =   14737632
         BackGradientStyle=   2
         BackStyle       =   4
         BevelHeight     =   5
         BackGradientExPercent=   80
         BackGlassColorStyle=   1
         BackGradientAutoValue=   40
         BackGlassAutoValue=   70
         BackLightShadowShadowValue=   -30
         BackLightShadowLightValue=   30
         BorderStyle     =   0
         BorderWidth     =   1
         BorderColor     =   16744576
         MaskColor       =   13828096
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "����"
         TextWidthPos    =   2
         TextHeightPos   =   2
         TextWidthMargin =   5
         TextHeightMargin=   5
         TextColor       =   128
         IconPosition    =   2
         IconAndTextMargin=   0
         IconMaskColor   =   13828096
         MouseOverMargin =   2
         MouseOverEffectAutoValue=   -20
         MouseDownBorderEffectValue=   -40
         MouseDownDefaultValue=   20
         FocusDefaultMargin=   3
         FocusColor1     =   16777152
         FocusColor2     =   16777088
         FocusColorStyle =   1
         FocusColorMargin=   2
         FocusEffectAutoValue=   -20
         ToolTipBodyText =   "XBUTTON 2"
         ToolTipTitleText=   ""
         ToolTipCentered =   -1  'True
         ToolTipBackColor=   12648447
         ToolTipExBackColor1=   12648447
         ToolTipExHoverTime=   1000
         ToolTipExPopupTime=   10000
         ToolTipExPopupPos=   0
         ToolTipExArrowWidth=   10
         ToolTipExArrowHeight=   15
         ToolTipExBorderRoundNum=   0
         ToolTipExPopupPosWMargin=   5
         ToolTipExPopupPosHMargin=   5
         ToolTipExBackColor2=   16777215
         ToolTipExBorderColor=   4210752
         ToolTipExTitleText=   "Title"
         ToolTipExIconAndTitleMargin=   5
         ToolTipExTitleAlign=   2
         BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTopMargin=   5
         ToolTipExBottomMargin=   5
         ToolTipExLeftMargin=   5
         ToolTipExRightMargin=   5
         ToolTipExBodyText=   "Body Text"
         ToolTipExBodyTextColor=   4210752
         BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTitleLineColor=   4210752
         ToolTipExTitleAndLineMargin=   5
         ToolTipExPostScriptText=   "PostScript"
         ToolTipExIconAndPostScriptMargin=   5
         ToolTipExPostScriptLineColor=   4210752
         BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ToolTipExTitleLineShadow=   -1  'True
         ToolTipExTitleLine=   -1  'True
         ToolTipExTitleLineLeftMargin=   5
         ToolTipExTitleLineRightMargin=   5
         ToolTipExPostScriptLineShadow=   -1  'True
         ToolTipExPostScriptLine=   -1  'True
         ToolTipExPostScriptLineLeftMargin=   5
         ToolTipExPostScriptLineRightMargin=   5
         ToolTipExTitleAndBodyMargin=   5
         ToolTipExBodyAndPostScriptMargin=   5
         ToolTipExTitleTextBackColor=   16777215
         ToolTipExTitleIconMaskColor=   13828096
         ToolTipExTitleIconAndTextAlign=   2
         ToolTipExTitleIconAndTextMargin=   5
         ToolTipExPopupAutoPos=   -1  'True
         ToolTipExPostScriptAndLineMargin=   5
         ToolTipExPostScriptIconPos=   1
         ToolTipExPostScriptIconAndTextMargin=   5
         ToolTipExPostScriptIconAndTextAlign=   2
         ToolTipExPostScriptIconMaskColor=   13828096
         ToolTipExBodyTextBackColor=   16761024
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   0
         X2              =   17190
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   -30
         X2              =   17190
         Y1              =   90
         Y2              =   90
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   8235
      Left            =   17760
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   15015
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10260
         Picture         =   "frmInterface.frx":14F5
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   6270
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   8910
         TabIndex        =   7
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8910
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   6
         Top             =   5580
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   12480
         TabIndex        =   5
         Top             =   5580
         Width           =   1125
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13680
         TabIndex        =   4
         Top             =   5610
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   10530
         MultiLine       =   -1  'True
         ScrollBars      =   3  '�����
         TabIndex        =   3
         Top             =   5580
         Width           =   1875
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   1050
         Left            =   10230
         TabIndex        =   2
         Top             =   6720
         Width           =   1425
         Begin MSCommLib.MSComm comEqp 
            Left            =   135
            Top             =   300
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            RThreshold      =   1
            RTSEnable       =   -1  'True
            EOFEnable       =   -1  'True
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   720
            Top             =   270
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin FPSpreadADO.fpSpread vasPrint 
         Height          =   1545
         Left            =   90
         TabIndex        =   9
         Top             =   1590
         Width           =   8160
         _Version        =   524288
         _ExtentX        =   14393
         _ExtentY        =   2725
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
         MaxCols         =   9
         SpreadDesigner  =   "frmInterface.frx":1A7F
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread vasPrintBuf 
         Height          =   1245
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   8115
         _Version        =   524288
         _ExtentX        =   14314
         _ExtentY        =   2196
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
         SpreadDesigner  =   "frmInterface.frx":2119
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread vasExcel 
         Height          =   1875
         Left            =   90
         TabIndex        =   11
         Top             =   3180
         Width           =   8205
         _Version        =   524288
         _ExtentX        =   14473
         _ExtentY        =   3307
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
         SpreadDesigner  =   "frmInterface.frx":256F
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread spdTot_Print 
         Height          =   1335
         Left            =   120
         TabIndex        =   12
         Top             =   5220
         Width           =   8235
         _Version        =   524288
         _ExtentX        =   14526
         _ExtentY        =   2355
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   41
         OperationMode   =   2
         SelectBlockOptions=   0
         ShadowColor     =   14548991
         SpreadDesigner  =   "frmInterface.frx":29C5
         UserResize      =   2
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread vasCode 
         Height          =   1125
         Left            =   8910
         TabIndex        =   13
         Top             =   180
         Width           =   5685
         _Version        =   524288
         _ExtentX        =   10028
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":58F0
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread vasTemp1 
         Height          =   1305
         Left            =   8970
         TabIndex        =   14
         Top             =   4020
         Width           =   5685
         _Version        =   524288
         _ExtentX        =   10028
         _ExtentY        =   2302
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
         SpreadDesigner  =   "frmInterface.frx":5D46
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread vasResTemp 
         Height          =   1425
         Left            =   8910
         TabIndex        =   15
         Top             =   1350
         Width           =   5715
         _Version        =   524288
         _ExtentX        =   10081
         _ExtentY        =   2514
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
         SpreadDesigner  =   "frmInterface.frx":619C
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread vasTemp 
         Height          =   1125
         Left            =   8910
         TabIndex        =   16
         Top             =   2820
         Width           =   5715
         _Version        =   524288
         _ExtentX        =   10081
         _ExtentY        =   1984
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
         SpreadDesigner  =   "frmInterface.frx":65F2
         AppearanceStyle =   0
      End
      Begin BHButton.BHImageButton cmdMode 
         Height          =   405
         Left            =   8970
         TabIndex        =   17
         Top             =   6960
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   714
         Caption         =   "���ո��"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label8 
         Caption         =   "�˻縶����"
         Height          =   1095
         Left            =   8610
         TabIndex        =   20
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   12510
         TabIndex        =   19
         Top             =   6060
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   13710
         TabIndex        =   18
         Top             =   6060
         Width           =   915
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   12660
      Width           =   18285
      _ExtentX        =   32253
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3810
            MinWidth        =   3809
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   14994
            MinWidth        =   14994
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "2015-04-29"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "���� 11:29"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "�ѱ��ǰ�������ȸ"
            TextSave        =   "�ѱ��ǰ�������ȸ"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSFrame ssfMst 
      Height          =   9645
      Left            =   240
      TabIndex        =   40
      Top             =   3960
      Visible         =   0   'False
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   17013
      _Version        =   262144
      BackColor       =   16777215
      Begin VB.Timer Timer1 
         Left            =   3750
         Top             =   3900
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   8805
         Left            =   90
         TabIndex        =   41
         Top             =   780
         Width           =   16875
         _ExtentX        =   29766
         _ExtentY        =   15531
         _Version        =   262144
         BackColor       =   16777215
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   8625
            Left            =   5190
            Picture         =   "frmInterface.frx":6A48
            ScaleHeight     =   8565
            ScaleWidth      =   11535
            TabIndex        =   46
            Top             =   90
            Width           =   11595
         End
         Begin FPSpreadADO.fpSpread spdMst 
            CausesValidation=   0   'False
            Height          =   8685
            Left            =   60
            TabIndex        =   47
            Tag             =   "20001"
            Top             =   30
            Width           =   5085
            _Version        =   524288
            _ExtentX        =   8969
            _ExtentY        =   15319
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            MaxCols         =   5
            MaxRows         =   10
            Protect         =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmInterface.frx":23669
            VisibleCols     =   5
            VisibleRows     =   10
            TextTip         =   2
            CellNoteIndicatorColor=   16576
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   30
         TabIndex        =   42
         Top             =   30
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   1191
         _Version        =   262144
         BackColor       =   -2147483629
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin XLibrary_XButton.XButton XButton2 
            Height          =   405
            Left            =   3000
            TabIndex        =   62
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "��ȸ"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XComboBox.XComboBox XComboBox2 
            Height          =   315
            Left            =   1500
            TabIndex        =   63
            Top             =   180
            Width           =   1425
            _ExtentX        =   2514
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
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
            ToolTipTitleText=   "Title"
            ToolTipBodyText =   "XComboBox"
            TextColor       =   0
            ListBgColor     =   16777215
            ListTextColor   =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XButton.XButton XButton3 
            Height          =   405
            Left            =   13080
            TabIndex        =   64
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "���"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton4 
            Height          =   405
            Left            =   14160
            TabIndex        =   65
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Excel"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton5 
            Height          =   405
            Left            =   15690
            TabIndex        =   66
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "ȭ������"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            X1              =   15450
            X2              =   15450
            Y1              =   180
            Y2              =   510
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�˻�μ�"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   420
            TabIndex        =   45
            Top             =   240
            Width           =   720
         End
      End
   End
   Begin Threed.SSFrame ssfRpt 
      Height          =   9645
      Left            =   240
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   17013
      _Version        =   262144
      BackColor       =   16777215
      Begin Threed.SSPanel SSPanel5 
         Height          =   8805
         Left            =   90
         TabIndex        =   33
         Top             =   780
         Width           =   16875
         _ExtentX        =   29766
         _ExtentY        =   15531
         _Version        =   262144
         BackColor       =   16777215
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox Check2 
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   34
            Top             =   60
            Width           =   225
         End
         Begin FPSpreadADO.fpSpread spdRpt 
            CausesValidation=   0   'False
            Height          =   8745
            Left            =   30
            TabIndex        =   35
            Tag             =   "20001"
            Top             =   30
            Width           =   16815
            _Version        =   524288
            _ExtentX        =   29660
            _ExtentY        =   15425
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            MaxCols         =   15
            MaxRows         =   10
            Protect         =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmInterface.frx":23DC7
            VisibleCols     =   10
            VisibleRows     =   10
            TextTip         =   2
            CellNoteIndicatorColor=   16576
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   36
         Top             =   30
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   1191
         _Version        =   262144
         BackColor       =   -2147483629
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSOption SSOption1 
            Height          =   255
            Left            =   180
            TabIndex        =   38
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   450
            _Version        =   262144
            BackColor       =   -2147483629
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "��������"
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   315
            Left            =   1500
            TabIndex        =   37
            Top             =   210
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   65208321
            CurrentDate     =   40248
         End
         Begin Threed.SSOption SSOption2 
            Height          =   255
            Left            =   180
            TabIndex        =   39
            Top             =   390
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   450
            _Version        =   262144
            BackColor       =   -2147483629
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�Ƿ�����"
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   495
            Left            =   7350
            TabIndex        =   53
            Top             =   90
            Width           =   5565
            _ExtentX        =   9816
            _ExtentY        =   873
            _Version        =   262144
            BackColor       =   -2147483629
            Begin XLibrary_XTextBox.XTextBox XTextBox4 
               Height          =   285
               Left            =   900
               TabIndex        =   54
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin XLibrary_XTextBox.XTextBox XTextBox5 
               Height          =   285
               Left            =   2670
               TabIndex        =   55
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin XLibrary_XTextBox.XTextBox XTextBox6 
               Height          =   285
               Left            =   4470
               TabIndex        =   56
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "�˻�Ǽ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   3690
               TabIndex        =   59
               Top             =   180
               Width           =   720
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "��ü�Ǽ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   1875
               TabIndex        =   58
               Top             =   180
               Width           =   720
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "�ǷڰǼ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   57
               Top             =   180
               Width           =   720
            End
         End
         Begin XLibrary_XButton.XButton XButton6 
            Height          =   405
            Left            =   4500
            TabIndex        =   67
            Top             =   150
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "��ȸ"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton7 
            Height          =   405
            Left            =   13080
            TabIndex        =   68
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "���"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton8 
            Height          =   405
            Left            =   14160
            TabIndex        =   69
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Excel"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton9 
            Height          =   405
            Left            =   15690
            TabIndex        =   70
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "ȭ������"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XComboBox.XComboBox XComboBox3 
            Height          =   315
            Left            =   3000
            TabIndex        =   71
            Top             =   210
            Width           =   1425
            _ExtentX        =   2514
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
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
            ToolTipTitleText=   "Title"
            ToolTipBodyText =   "XComboBox"
            TextColor       =   0
            ListBgColor     =   16777215
            ListTextColor   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            X1              =   15450
            X2              =   15450
            Y1              =   180
            Y2              =   510
         End
      End
   End
   Begin Threed.SSFrame ssfRst 
      Height          =   9675
      Left            =   240
      TabIndex        =   23
      Top             =   1740
      Visible         =   0   'False
      Width           =   17025
      _ExtentX        =   30030
      _ExtentY        =   17066
      _Version        =   262144
      BackColor       =   16777215
      Begin Threed.SSPanel SSPanel4 
         Height          =   8775
         Left            =   90
         TabIndex        =   28
         Top             =   780
         Width           =   16845
         _ExtentX        =   29713
         _ExtentY        =   15478
         _Version        =   262144
         BackColor       =   16777215
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox Check1 
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   210
            TabIndex        =   29
            Top             =   90
            Width           =   225
         End
         Begin FPSpreadADO.fpSpread spdRstD 
            CausesValidation=   0   'False
            Height          =   8715
            Left            =   3990
            TabIndex        =   30
            Tag             =   "20001"
            Top             =   30
            Width           =   12825
            _Version        =   524288
            _ExtentX        =   22622
            _ExtentY        =   15372
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            MaxCols         =   11
            MaxRows         =   10
            Protect         =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   16761087
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmInterface.frx":245DE
            VisibleCols     =   9
            VisibleRows     =   10
            TextTip         =   2
            CellNoteIndicatorColor=   16576
         End
         Begin FPSpreadADO.fpSpread spdRstH 
            CausesValidation=   0   'False
            Height          =   8715
            Left            =   30
            TabIndex        =   31
            Tag             =   "20001"
            Top             =   30
            Width           =   3945
            _Version        =   524288
            _ExtentX        =   6959
            _ExtentY        =   15372
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            MaxCols         =   13
            MaxRows         =   10
            Protect         =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmInterface.frx":24D78
            VisibleCols     =   10
            VisibleRows     =   10
            TextTip         =   2
            CellNoteIndicatorColor=   16576
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   24
         Top             =   30
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   1191
         _Version        =   262144
         BackColor       =   -2147483629
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1500
            TabIndex        =   72
            Top             =   180
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   65208321
            CurrentDate     =   40248
         End
         Begin XLibrary_XButton.XButton XButton10 
            Height          =   405
            Left            =   4500
            TabIndex        =   73
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "��ȸ"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XComboBox.XComboBox XComboBox1 
            Height          =   315
            Left            =   3000
            TabIndex        =   74
            Top             =   180
            Width           =   1425
            _ExtentX        =   2514
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
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
            ToolTipTitleText=   "Title"
            ToolTipBodyText =   "XComboBox"
            TextColor       =   0
            ListBgColor     =   16777215
            ListTextColor   =   0
            Enabled         =   -1  'True
         End
         Begin XLibrary_XButton.XButton XButton11 
            Height          =   405
            Left            =   5760
            TabIndex        =   76
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "���Ȯ��"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            TextColor       =   16711680
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   495
            Left            =   7350
            TabIndex        =   77
            Top             =   90
            Width           =   5565
            _ExtentX        =   9816
            _ExtentY        =   873
            _Version        =   262144
            BackColor       =   -2147483629
            Begin XLibrary_XTextBox.XTextBox XTextBox1 
               Height          =   285
               Left            =   900
               TabIndex        =   78
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin XLibrary_XTextBox.XTextBox XTextBox2 
               Height          =   285
               Left            =   2670
               TabIndex        =   79
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin XLibrary_XTextBox.XTextBox XTextBox3 
               Height          =   285
               Left            =   4470
               TabIndex        =   80
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "�ǷڰǼ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   83
               Top             =   180
               Width           =   720
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "��ü�Ǽ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   1875
               TabIndex        =   82
               Top             =   180
               Width           =   720
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "�˻�Ǽ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   3690
               TabIndex        =   81
               Top             =   180
               Width           =   720
            End
         End
         Begin XLibrary_XButton.XButton XButton12 
            Height          =   405
            Left            =   13080
            TabIndex        =   84
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "���"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton13 
            Height          =   405
            Left            =   14160
            TabIndex        =   85
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Excel"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton14 
            Height          =   405
            Left            =   15690
            TabIndex        =   86
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "ȭ������"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            X1              =   15450
            X2              =   15450
            Y1              =   180
            Y2              =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�Ƿ�����"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   75
            Top             =   240
            Width           =   720
         End
      End
   End
   Begin XLibrary_XButton.XButton XButton1 
      Height          =   525
      Left            =   11460
      TabIndex        =   61
      Top             =   11880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   926
      BackColor1      =   16777215
      BackColor2      =   16777215
      BackColorEx     =   14737632
      BackGradientStyle=   2
      BackStyle       =   4
      BevelHeight     =   5
      BackGradientExPercent=   80
      BackGlassColorStyle=   1
      BackGradientAutoValue=   40
      BackGlassAutoValue=   70
      BackLightShadowShadowValue=   -30
      BackLightShadowLightValue=   30
      BorderStyle     =   0
      BorderWidth     =   1
      BorderColor     =   16744576
      MaskColor       =   13828096
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "����"
      TextWidthPos    =   2
      TextHeightPos   =   2
      TextWidthMargin =   5
      TextHeightMargin=   5
      IconPosition    =   2
      IconAndTextMargin=   0
      IconMaskColor   =   13828096
      MouseOverMargin =   2
      MouseOverEffectAutoValue=   -20
      MouseDownBorderEffectValue=   -40
      MouseDownDefaultValue=   20
      FocusDefaultMargin=   3
      FocusColor1     =   16777152
      FocusColor2     =   16777088
      FocusColorStyle =   1
      FocusColorMargin=   2
      FocusEffectAutoValue=   -20
      ToolTipBodyText =   "XBUTTON 2"
      ToolTipTitleText=   ""
      ToolTipCentered =   -1  'True
      ToolTipBackColor=   12648447
      ToolTipExBackColor1=   12648447
      ToolTipExHoverTime=   1000
      ToolTipExPopupTime=   10000
      ToolTipExPopupPos=   0
      ToolTipExArrowWidth=   10
      ToolTipExArrowHeight=   15
      ToolTipExBorderRoundNum=   0
      ToolTipExPopupPosWMargin=   5
      ToolTipExPopupPosHMargin=   5
      ToolTipExBackColor2=   16777215
      ToolTipExBorderColor=   4210752
      ToolTipExTitleText=   "Title"
      ToolTipExIconAndTitleMargin=   5
      ToolTipExTitleAlign=   2
      BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipExTopMargin=   5
      ToolTipExBottomMargin=   5
      ToolTipExLeftMargin=   5
      ToolTipExRightMargin=   5
      ToolTipExBodyText=   "Body Text"
      ToolTipExBodyTextColor=   4210752
      BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipExTitleLineColor=   4210752
      ToolTipExTitleAndLineMargin=   5
      ToolTipExPostScriptText=   "PostScript"
      ToolTipExIconAndPostScriptMargin=   5
      ToolTipExPostScriptLineColor=   4210752
      BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTipExTitleLineShadow=   -1  'True
      ToolTipExTitleLine=   -1  'True
      ToolTipExTitleLineLeftMargin=   5
      ToolTipExTitleLineRightMargin=   5
      ToolTipExPostScriptLineShadow=   -1  'True
      ToolTipExPostScriptLine=   -1  'True
      ToolTipExPostScriptLineLeftMargin=   5
      ToolTipExPostScriptLineRightMargin=   5
      ToolTipExTitleAndBodyMargin=   5
      ToolTipExBodyAndPostScriptMargin=   5
      ToolTipExTitleTextBackColor=   16777215
      ToolTipExTitleIconMaskColor=   13828096
      ToolTipExTitleIconAndTextAlign=   2
      ToolTipExTitleIconAndTextMargin=   5
      ToolTipExPopupAutoPos=   -1  'True
      ToolTipExPostScriptAndLineMargin=   5
      ToolTipExPostScriptIconPos=   1
      ToolTipExPostScriptIconAndTextMargin=   5
      ToolTipExPostScriptIconAndTextAlign=   2
      ToolTipExPostScriptIconMaskColor=   13828096
      ToolTipExBodyTextBackColor=   16761024
   End
   Begin Threed.SSFrame ssfReg 
      Height          =   9645
      Left            =   240
      TabIndex        =   21
      Top             =   900
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   17013
      _Version        =   262144
      BackColor       =   16777215
      Begin Threed.SSPanel SSPanel3 
         Height          =   8775
         Left            =   90
         TabIndex        =   25
         Top             =   780
         Width           =   16875
         _ExtentX        =   29766
         _ExtentY        =   15478
         _Version        =   262144
         BackColor       =   16777215
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkOrder 
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   26
            Top             =   60
            Width           =   225
         End
         Begin FPSpreadADO.fpSpread spdRcp 
            CausesValidation=   0   'False
            Height          =   8685
            Left            =   30
            TabIndex        =   27
            Tag             =   "20001"
            Top             =   30
            Width           =   16815
            _Version        =   524288
            _ExtentX        =   29660
            _ExtentY        =   15319
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            ColHeaderDisplay=   0
            DisplayRowHeaders=   0   'False
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            MaxCols         =   15
            MaxRows         =   10
            Protect         =   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmInterface.frx":25523
            VisibleCols     =   10
            VisibleRows     =   10
            TextTip         =   2
            CellNoteIndicatorColor=   16576
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   30
         TabIndex        =   87
         Top             =   30
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   1191
         _Version        =   262144
         BackColor       =   -2147483629
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   1500
            TabIndex        =   88
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   65208321
            CurrentDate     =   40248
         End
         Begin XLibrary_XButton.XButton XButton15 
            Height          =   405
            Left            =   4500
            TabIndex        =   89
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "��ȸ"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton16 
            Height          =   405
            Left            =   5760
            TabIndex        =   90
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "�Ƿڵ��"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            TextColor       =   16711680
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   495
            Left            =   7350
            TabIndex        =   91
            Top             =   90
            Width           =   5565
            _ExtentX        =   9816
            _ExtentY        =   873
            _Version        =   262144
            BackColor       =   -2147483629
            Begin XLibrary_XTextBox.XTextBox XTextBox7 
               Height          =   285
               Left            =   900
               TabIndex        =   92
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin XLibrary_XTextBox.XTextBox XTextBox8 
               Height          =   285
               Left            =   2670
               TabIndex        =   93
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin XLibrary_XTextBox.XTextBox XTextBox9 
               Height          =   285
               Left            =   4470
               TabIndex        =   94
               Top             =   120
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               BackColor       =   16777215
               BorderColor     =   16744576
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����ü"
                  Size            =   9.75
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
               ToolTipIcon     =   0
               ToolTipPopupTime=   -1
               ToolTipHoverTime=   -1
               ToolTipBackColor=   16777215
               ToolTipForeColor=   0
               ToolTipStyle    =   0
               ToolTipCentered =   0   'False
               ToolTipTitleText=   "Title"
               ToolTipBodyText =   "XTextBox"
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
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "�˻�Ǽ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   3690
               TabIndex        =   97
               Top             =   180
               Width           =   720
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "��ü�Ǽ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   1875
               TabIndex        =   96
               Top             =   180
               Width           =   720
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               Caption         =   "�ǷڰǼ�"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   95
               Top             =   180
               Width           =   720
            End
         End
         Begin XLibrary_XButton.XButton XButton17 
            Height          =   405
            Left            =   13080
            TabIndex        =   98
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "���"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton18 
            Height          =   405
            Left            =   14160
            TabIndex        =   99
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Excel"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin XLibrary_XButton.XButton XButton19 
            Height          =   405
            Left            =   15690
            TabIndex        =   100
            Top             =   120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BackColor1      =   16777215
            BackColor2      =   16777215
            BackColorEx     =   14737632
            BackGradientStyle=   2
            BackStyle       =   4
            BevelHeight     =   5
            BackGradientExPercent=   80
            BackGlassColorStyle=   1
            BackGradientAutoValue=   40
            BackGlassAutoValue=   70
            BackLightShadowShadowValue=   -30
            BackLightShadowLightValue=   30
            BorderStyle     =   0
            BorderWidth     =   1
            BorderColor     =   16744576
            MaskColor       =   13828096
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "ȭ������"
            TextWidthPos    =   2
            TextHeightPos   =   2
            TextWidthMargin =   5
            TextHeightMargin=   5
            IconPosition    =   2
            IconAndTextMargin=   0
            IconMaskColor   =   13828096
            MouseOverMargin =   2
            MouseOverEffectAutoValue=   -20
            MouseDownBorderEffectValue=   -40
            MouseDownDefaultValue=   20
            FocusDefaultMargin=   3
            FocusColor1     =   16777152
            FocusColor2     =   16777088
            FocusColorStyle =   1
            FocusColorMargin=   2
            FocusEffectAutoValue=   -20
            ToolTipBodyText =   "XBUTTON 2"
            ToolTipTitleText=   ""
            ToolTipCentered =   -1  'True
            ToolTipBackColor=   12648447
            ToolTipExBackColor1=   12648447
            ToolTipExHoverTime=   1000
            ToolTipExPopupTime=   10000
            ToolTipExPopupPos=   0
            ToolTipExArrowWidth=   10
            ToolTipExArrowHeight=   15
            ToolTipExBorderRoundNum=   0
            ToolTipExPopupPosWMargin=   5
            ToolTipExPopupPosHMargin=   5
            ToolTipExBackColor2=   16777215
            ToolTipExBorderColor=   4210752
            ToolTipExTitleText=   "Title"
            ToolTipExIconAndTitleMargin=   5
            ToolTipExTitleAlign=   2
            BeginProperty ToolTipExTitleTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTopMargin=   5
            ToolTipExBottomMargin=   5
            ToolTipExLeftMargin=   5
            ToolTipExRightMargin=   5
            ToolTipExBodyText=   "Body Text"
            ToolTipExBodyTextColor=   4210752
            BeginProperty ToolTipExBodyTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineColor=   4210752
            ToolTipExTitleAndLineMargin=   5
            ToolTipExPostScriptText=   "PostScript"
            ToolTipExIconAndPostScriptMargin=   5
            ToolTipExPostScriptLineColor=   4210752
            BeginProperty ToolTipExPostScriptTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ToolTipExTitleLineShadow=   -1  'True
            ToolTipExTitleLine=   -1  'True
            ToolTipExTitleLineLeftMargin=   5
            ToolTipExTitleLineRightMargin=   5
            ToolTipExPostScriptLineShadow=   -1  'True
            ToolTipExPostScriptLine=   -1  'True
            ToolTipExPostScriptLineLeftMargin=   5
            ToolTipExPostScriptLineRightMargin=   5
            ToolTipExTitleAndBodyMargin=   5
            ToolTipExBodyAndPostScriptMargin=   5
            ToolTipExTitleTextBackColor=   16777215
            ToolTipExTitleIconMaskColor=   13828096
            ToolTipExTitleIconAndTextAlign=   2
            ToolTipExTitleIconAndTextMargin=   5
            ToolTipExPopupAutoPos=   -1  'True
            ToolTipExPostScriptAndLineMargin=   5
            ToolTipExPostScriptIconPos=   1
            ToolTipExPostScriptIconAndTextMargin=   5
            ToolTipExPostScriptIconAndTextAlign=   2
            ToolTipExPostScriptIconMaskColor=   13828096
            ToolTipExBodyTextBackColor=   16761024
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   315
            Left            =   3090
            TabIndex        =   102
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   65208321
            CurrentDate     =   40248
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2940
            TabIndex        =   103
            Top             =   210
            Width           =   60
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   101
            Top             =   240
            Width           =   720
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            X1              =   15450
            X2              =   15450
            Y1              =   180
            Y2              =   510
         End
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "2015�� 4�� 24�� 110�� ���� 115�� ���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   11790
      TabIndex        =   44
      Top             =   360
      Width           =   4170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "2015�� 4�� 25�� 120�� ���� 125�� ���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   11790
      TabIndex        =   43
      Top             =   540
      Width           =   4155
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȫ�浿���� �α����Ͽ����ϴ� - 2015.04.25 12:45"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   180
      Left            =   11790
      TabIndex        =   22
      Top             =   180
      Width           =   4155
   End
   Begin VB.Menu MnMain 
      Caption         =   "����"
      Begin VB.Menu MnMode 
         Caption         =   "�и�ȭ��"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MnMode 
         Caption         =   "����ȭ��"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MnSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnLog 
         Caption         =   "�α׺���"
         Visible         =   0   'False
      End
      Begin VB.Menu MnSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "ȯ�漳��"
      Begin VB.Menu MnTConfig 
         Caption         =   "��ż���"
      End
      Begin VB.Menu MnSep11 
         Caption         =   "-"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "�ڵ弳��"
      End
      Begin VB.Menu MnSep12 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnDBConfig 
         Caption         =   "DB����"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "LIS����"
      Begin VB.Menu MnSave 
         Caption         =   "�ڵ�����"
         Index           =   0
      End
      Begin VB.Menu MnSave 
         Caption         =   "��������"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- �˻����� Spread
Const colChkBox = 1
Const colHospCd = 2
Const colOrdDt = 3
Const colBarCd = 4
Const colPNm = 5
Const colDOB = 6
Const colSex = 7
Const colSpcCd = 8
Const colSpcNm = 9
Const colTstCd = 10
Const colTstNm = 11
Const colReqDt = 12
Const colReqTm = 13
Const colReqID = 14
Const colReqNm = 15


Private Sub BHImageButton8_Click()
    
    Timer1.Interval = 50
    Timer1.Enabled = True
    XProgress1.Value = 1
    
    XProgress1.ProgressStyle = Text1.Text
    
    XProgress1.Visible = True

    
End Sub

Private Sub cmdClose_Click()
    
    If MsgBox("RMS ���α׷��� �����Ͻüý��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
        
    
'    Dim Retval As Long
'
''    Retval = SetParent(XProgress1.hwnd, StatusBar1.hwnd)
'    Retval = SetParent(XProgress1.hwnd, StatusBar1.hwnd)
    
    Call SetFrmInitial

End Sub


Private Sub SetFrmInitial()

    ssfReg.Left = 240
    ssfRst.Left = 240
    ssfRpt.Left = 240
    ssfMst.Left = 240
    
    ssfReg.Top = 840
    ssfRst.Top = 840
    ssfRpt.Top = 840
    ssfMst.Top = 840
    
    
    spdRcp.MaxRows = 0
    spdRpt.MaxRows = 0
    spdRstH.MaxRows = 0
    spdRstD.MaxRows = 0
    spdMst.MaxRows = 0
    

    XProgress1.Min = 1
    XProgress1.Max = 100
    
    
    XProgress1.Left = StatusBar1.Left + StatusBar1.Panels(1).Width + 50
    XProgress1.Top = StatusBar1.Top + 50
    XProgress1.Width = StatusBar1.Panels(2).Width - 20
    XProgress1.Height = StatusBar1.Height - 40

    XProgress1.Visible = False

End Sub


Private Sub ssMenu_Click(Index As Integer, Value As Integer)
    
    If Index = 0 Then
        ssfReg.Visible = True
        ssfRst.Visible = False
        ssfRpt.Visible = False
        ssfMst.Visible = False
        
        ssMenu(0).BackColor = &H80000013
        ssMenu(1).BackColor = &H80000018
        ssMenu(2).BackColor = &H80000018
        ssMenu(3).BackColor = &H80000018
    ElseIf Index = 1 Then
        ssfReg.Visible = False
        ssfRst.Visible = True
        ssfRpt.Visible = False
        ssfMst.Visible = False
    
        ssMenu(0).BackColor = &H80000018
        ssMenu(1).BackColor = &H80000013
        ssMenu(2).BackColor = &H80000018
        ssMenu(3).BackColor = &H80000018
    ElseIf Index = 2 Then
        ssfReg.Visible = False
        ssfRst.Visible = False
        ssfRpt.Visible = True
        ssfMst.Visible = False
    
        ssMenu(0).BackColor = &H80000018
        ssMenu(1).BackColor = &H80000018
        ssMenu(2).BackColor = &H80000013
        ssMenu(3).BackColor = &H80000018
    ElseIf Index = 3 Then
        ssfReg.Visible = False
        ssfRst.Visible = False
        ssfRpt.Visible = False
        ssfMst.Visible = True
    
        ssMenu(0).BackColor = &H80000018
        ssMenu(1).BackColor = &H80000018
        ssMenu(2).BackColor = &H80000018
        ssMenu(3).BackColor = &H80000013
    End If
End Sub

'Private Type OPENFILENAME
'    lStructSize As Long
'    hwndOwner As Long
'    hInstance As Long
'    lpstrFilter As String
'    lpstrCustomFilter As String
'    nMaxCustFilter As Long
'    nFilterIndex As Long
'    lpstrFile As String
'    nMaxFile As Long
'    lpstrFileTitle As String
'    nMaxFileTitle As Long
'    lpstrInitialDir As String
'    lpstrTitle As String
'    FLAGS As Long
'    nFileOffset As Integer
'    nFileExtension As Integer
'    lpstrDefExt As String
'    lCustData As Long
'    lpfnHook As Long
'    lpTemplateName As String
'End Type
'
'
'Dim OFName As OPENFILENAME
'
'
'
'Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
'Dim sCnt As Long
'
'    SeqSearch = 0
'    If brspread.MaxRows <= 0 Then
'        Exit Function
'    End If
'
'    With brspread
'        For sCnt = 1 To .MaxRows
'            .Row = sCnt
'            .Col = brCol
'            If Trim(.Text) = brSeq Then
'                SeqSearch = sCnt 'brSeq
'                .Action = ActionActiveCell
'                .Refresh
'                Exit For
'            End If
'        Next sCnt
'    End With
'
'End Function
'
''spdorder, vasrid colum
''Const colSpecNo = 0 '�̻��
''Const colCheckBox = 1
''Const colBarcode = 2
''Const colRack = 3
''Const colDISK = 3
''Const colPos = 4
''Const colPID = 5
''Const colPName = 6
''Const colSex = 7
''Const colAge = 8
''
''Const colOCnt = 9
''Const colRCnt = 10
''Const colState = 11
''
''Const colA1c = 12
''Const colIFCC = 13
''Const coleAg = 14
'
'
'
'
''sendflag
''0: Order
''1: Result
''2: Trans
'
''vasres, vasrres colum
''Const colEquipCode = 1
''Const colExamCode = 2
''Const colExamName = 3
'''Const colMachResult = 4
''Const colResult = 4
''Const colSeq = 5
''Const colFLAG = 6
'
''Dim gRow As Long
''Dim gsBarCode As String
''Dim gsSampleType As String
''Dim gsPID As String
''Dim gsRackNo As String
''Dim gsPosNo As String
''Dim gsResDateTime As String
''Dim gsSeqNo As String
''Dim gsExamCode As String
''Dim gsExamName As String
''Dim gsOrder As String
''Dim gsResult As String
''Dim gsFlag As String
''
''Dim gMT As String
''Dim gComState As Long
''Dim gErrState As Long
''
''Dim gIFCC1 As String
''Dim gIFCC2 As String
''Dim geAg1 As String
''Dim geAg2 As String
''Dim gADD_IFCC As String
''Dim gADD_eAg As String
''
''Dim strBuffer As String
''
''Public gENQFlag As Integer
''Public gNAKFlag As Integer
'
''===============================
''Const STX As String = ""
''Const ETX As String = ""
''Const ENQ As String = ""
''Const ACK As String = ""
''Const NAK As String = ""
''Const EOT As String = ""
''Const ETB As String = ""
''Const FS  As String = ""
''Const Rst As String = ""
''Const GS  As String = ""
''
''
''Dim strRecvData()   As String
''Dim intPhase        As Integer
''Dim strState        As String
''Dim intBufCnt       As Integer
''Dim blnIsETB        As Boolean
''Dim intSndPhase     As Integer
''Dim intFrameNo      As Integer
''===============================
'
'
''Private Sub chkAll_Click()
''    Dim iRow As Long
''
''    If chkAll.Value = 1 Then
''        For iRow = 1 To spdorder.DataRowCnt
''            spdorder.Row = iRow
''            spdorder.Col = 1
''
''            spdorder.Value = 1
''        Next iRow
''    ElseIf chkAll.Value = 0 Then
''        For iRow = 1 To spdorder.DataRowCnt
''            spdorder.Row = iRow
''            spdorder.Col = 1
''
''            spdorder.Value = 0
''        Next iRow
''    End If
''End Sub
'
'Private Sub chkMode_Click()
'    If chkMode.Value = 1 Then
'        chkMode.Caption = "�ڵ�����"
'        Call MnSave_Click(0)
'    Else
'        chkMode.Caption = "��������"
'        Call MnSave_Click(1)
'    End If
'
'
'End Sub
'
'
''Private Sub chkRAll_Click()
''    Dim iRow As Long
''
''    If chkRAll.Value = 1 Then
''        For iRow = 1 To vasRID.DataRowCnt
''            vasRID.Row = iRow
''            vasRID.Col = 1
''
''            vasRID.Value = 1
''        Next iRow
''    ElseIf chkRAll.Value = 0 Then
''        For iRow = 1 To vasRID.DataRowCnt
''            vasRID.Row = iRow
''            vasRID.Col = 1
''
''            vasRID.Value = 0
''        Next iRow
''    End If
''End Sub
'
''Private Sub cmdExcel_Click()
''    Dim iRow As Integer
''    Dim j As Integer
''
''    Dim sCurDate As String
''    Dim sSerDate As String
''    Dim sHead As String
''    Dim sFoot As String
''    Dim sFileName As String
''
''    Dim sA1c As String
''    Dim sIFCC As String
''    Dim seAg As String
''
''
''
''    ClearSpread vasPrint
''
''    j = 1
''
''    For iRow = 1 To vasRID.DataRowCnt
''        vasRID.Row = iRow
''        vasRID.Col = 1
''
''        If vasRID.Value = 1 Then
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colBarcode)), j, 1
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colPID)), j, 2
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colPName)), j, 3
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colSex)), j, 4
''
''            SQL = "SELECT RESULT " & vbCrLf & _
''                  "FROM PAT_RES " & vbCrLf & _
''                  "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
''                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''                  "  AND BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf & _
''                  "  AND PID = '" & Trim(GetText(vasPrint, iRow, 3)) & "' " & vbCrLf & _
''                  "ORDER BY SEQNO"
''            Res = GetDBSelectVas(gLocal, SQL, vasPrintBuf)
''
''            sA1c = GetText(vasPrintBuf, 1, 1)
''            sIFCC = GetText(vasPrintBuf, 2, 1)
''            seAg = GetText(vasPrintBuf, 3, 1)
''
''            ClearSpread vasPrintBuf, 1, 1
''
''            SetText vasPrint, sA1c, j, 7
''            SetText vasPrint, sIFCC, j, 8
''            SetText vasPrint, seAg, j, 9
''
''            '"GROUP BY BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, JUMIN, Hospital, SENDFLAG"
''
'''            SetText vasprint, Trim(GetText(vasrid, iRow, vasrid.MaxCols)), j, 8
'''            SetText vasprint, Trim(GetText(vasrid, iRow, 10)), j, 9
''
''            j = j + 1
''        End If
''    Next iRow
''
''    If vasPrint.DataRowCnt < 1 Then
''        MsgBox "������ �ڷᰡ �����ϴ�.", , "�� ��"
''        Exit Sub
''    Else
''        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
''        CommonDialog1.ShowSave
''        sFileName = CommonDialog1.Filename
''        SaveExcel sFileName, vasPrint
''
''    End If
''End Sub
'Sub SaveExcel(FileName As String, argSpread As fpSpread)
'
'On Error Resume Next
'
'' Excel Object Library �� �����մϴ�.
'Dim xlApp As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet
'
'Dim iRow As Integer
'Dim iCol As Integer
'Dim i As Integer
'
'    Set xlApp = CreateObject("Excel.Application")
'
'    xlApp.DisplayAlerts = False
'
'    Set xlBook = xlApp.Workbooks.Add
'
'    Set xlSheet = xlBook.Worksheets(1)
'
'    For iRow = 0 To argSpread.DataRowCnt
'        For iCol = 1 To argSpread.DataColCnt
'            argSpread.Row = iRow
'            argSpread.Col = iCol
'            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
'        Next iCol
'    Next iRow
'
'    xlBook.SaveAs (FileName)
'    xlApp.Quit
'
'
'End Sub
'
''Private Sub cmdIFClear_Click()
''    Dim i As Integer
''
''    Var_Clear
''
''    txtData.Text = ""
''    txtErr.Text = ""
''
''    SetForeColor spdSeparationOrder(0), 1, spdSeparationOrder(0).MaxRows, 1, spdSeparationOrder(0).MaxCols, 0, 0, 0
''    SetForeColor spdSeparationResult(0), 1, spdSeparationResult(0).MaxRows, 1, spdSeparationResult(0).MaxCols, 0, 0, 0
''
''    spdSeparationOrder(0).MaxRows = 0
''    spdSeparationResult(0).MaxRows = 0
''
'''    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
''
''    gRow = 0
''
''End Sub
'
''Private Sub cmdIFTrans_Click()
''    Dim lRow As Long
''
''    For lRow = 1 To spdorder.DataRowCnt
''        spdorder.Row = lRow
''        spdorder.Col = 1
''        If spdorder.Value = 1 Then
''
''            'If Mid(Trim(GetText(spdorder, lRow, 3)), 1, 2) = "99" Then
''            '    res = SaveTransDataW_QC(gRow)
''            'Else
''                Res = SaveTransDataW(gRow)
''            'End If
''
''            If Res = -1 Then
''                SetForeColor spdorder, lRow, lRow, 1, colState, 255, 0, 0
''                SetText spdorder, "Failed", lRow, colState
''            Else
''                spdorder.Row = lRow
''                spdorder.Col = 1
''                spdorder.Value = 1
''
''                SetBackColor spdorder, lRow, lRow, 1, colState, 202, 255, 112
''                SetText spdorder, "Trans", lRow, colState
''
''                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
''                      " SENDFLAG = '2' " & vbCrLf & _
''                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''                      " AND BARCODE = '" & Trim(GetText(spdorder, lRow, colBarcode)) & "' "
''                Res = SendQuery(gLocal, SQL)
''                If Res = -1 Then
''                    SaveQuery SQL
''                    Exit Sub
''                End If
''
''            End If
''            spdorder.Row = lRow
''            spdorder.Col = 1
''            spdorder.Value = 0
''        End If
''    Next lRow
''End Sub
'
''Private Sub cmdRClear_Click()
''    Dim i As Integer
''
'''    Var_Clear
''
''    txtData.Text = ""
''    txtErr.Text = ""
''
''    SetForeColor vasRID, 1, vasRID.MaxRows, 1, vasRID.MaxCols, 0, 0, 0
''    SetForeColor vasRRes, 1, vasRRes.MaxRows, 1, vasRRes.MaxCols, 0, 0, 0
''
''    vasRID.MaxRows = 0
''    vasRRes.MaxRows = 0
''
''    dtpExamDate = Date
''
''End Sub
'
''Private Sub cmdRSch_Click()
''    Dim iRow As Long
''
''    ClearSpread vasRID
''    ClearSpread vasRRes
''    Call chkRAll_Click
''
''    SQL = "SELECT '', BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, COUNT(*), COUNT(*), SENDFLAG " & vbCrLf & _
''          "FROM PAT_RES " & vbCrLf & _
''          "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
''          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''          "  AND SENDFLAG IN ('0','1', '2') " & vbCrLf & _
''          "GROUP BY BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG"
''    Res = GetDBSelectVas(gLocal, SQL, vasRID)
''
''          '"  AND SENDFLAG IN ('1', '2') "
''    If Res = -1 Then
''        SaveQuery SQL
''        Exit Sub
''    End If
''
''    For iRow = 1 To vasRID.DataRowCnt
''        Select Case Trim(GetText(vasRID, iRow, colState))
''        Case "2"
''            SetBackColor vasRID, iRow, iRow, 1, colState, 202, 255, 112
''            SetText vasRID, "�Ϸ�", iRow, colState
''        Case "0"
''            'SetText spdorder, "����", iRow, colState
''            SetText spdorder, "����", iRow, colState
''        Case "1"
''            SetText vasRID, "���", iRow, colState
''        End Select
''    Next iRow
''End Sub
'
''Private Sub cmdRTrans_Click()
''    Dim lRow As Long
''
''    For lRow = 1 To vasRID.DataRowCnt
''        vasRID.Row = lRow
''        vasRID.Col = 1
''        If vasRID.Value = 1 Then
''            Res = SaveTransDataR(lRow)
''
''            If Res = -1 Then
''                SetForeColor vasRID, lRow, lRow, 1, colState, 255, 0, 0
''                SetText vasRID, "Failed", lRow, colState
''            ElseIf Res = 0 Then
''
''            Else
''                vasRID.Row = lRow
''                vasRID.Col = 1
''                vasRID.Value = 1
''
''                SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
''                SetText vasRID, "Trans", lRow, colState
''
''                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
''                      " SENDFLAG = '2' " & vbCrLf & _
''                      " WHERE EQUIPNO = '" & gEquipCode & "' " & vbCrLf & _
''                      " AND BARCODE = '" & Trim(GetText(vasRID, lRow, colBarcode)) & "' "
''                Res = SendQuery(gLocal, SQL)
''                If Res = -1 Then
''                    SaveQuery SQL
''                    Exit Sub
''                End If
''
''            End If
''            vasRID.Row = lRow
''            vasRID.Col = 1
''            vasRID.Value = 0
''        End If
''    Next lRow
''End Sub
'
'Private Sub chkOrder_Click()
'    Dim iRow As Long
'
'    If gMode = "0" Then
'        If chkOrder.Value = 1 Then
'            For iRow = 1 To spdOrder.DataRowCnt
'                spdOrder.Row = iRow
'                spdOrder.Col = 1
'
'                spdOrder.Value = 1
'            Next iRow
'        ElseIf chkOrder.Value = 0 Then
'            For iRow = 1 To spdOrder.DataRowCnt
'                spdOrder.Row = iRow
'                spdOrder.Col = 1
'
'                spdOrder.Value = 0
'            Next iRow
'        End If
'    Else
'        If chkOrder.Value = 1 Then
'            For iRow = 1 To spdTot.DataRowCnt
'                spdTot.Row = iRow
'                spdTot.Col = 1
'
'                spdTot.Value = 1
'            Next iRow
'        ElseIf chkOrder.Value = 0 Then
'            For iRow = 1 To spdTot.DataRowCnt
'                spdTot.Row = iRow
'                spdTot.Col = 1
'
'                spdTot.Value = 0
'            Next iRow
'        End If
'    End If
'
'End Sub
'
'Private Sub cmdAdd_Click()
'
'    spdOrder.MaxRows = spdOrder.MaxRows + 1
'    spdOrder.RowHeight(-1) = 12.27
'
'End Sub
'
'Private Sub cmdClear_Click()
'
'    Call CtlInitializing
'
'    Call SpdInitializing
'
'End Sub
'
'
''Private Function ShowOpen(Ufilter As String, Upath As String) As String
''
''    OFName.lStructSize = Len(OFName)
''    OFName.hwndOwner = Me.hwnd
''    OFName.hInstance = App.hInstance
''    OFName.lpstrFilter = Ufilter
''    OFName.lpstrFile = Space$(254)
''    OFName.nMaxFile = 255
''    OFName.lpstrFileTitle = Space$(254)
''    OFName.nMaxFileTitle = 255
''    OFName.lpstrInitialDir = Upath
''    OFName.lpstrTitle = "Open File"
''    OFName.FLAGS = 0
''
''    If GetOpenFilename(OFName) Then
''        ShowOpen = Trim$(OFName.lpstrFile)
''        'ShowOpen = Mid(ShowOpen, 1, Len(ShowOpen) - 1)
''    Else
''        ShowOpen = ""
''    End If
''
''End Function
'
'Public Function Get_ExamCode(asExamName As String) As String
'    Dim strExamName As String
'    Get_ExamCode = ""
'    strExamName = Trim(asExamName)
'
'    SQL = "SELECT EXAMCODE FROM EQUIPEXAM WHERE EXAMNAME = '" & asExamName & "'"
'    Res = db_select_Col(gLocal, SQL)
'
'    If Res < 0 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    Get_ExamCode = Trim(gReadBuf(0))
'
'End Function
'
'Public Function Get_ExamName(asExamCode As String) As String
'    Dim strExamCode As String
'    Get_ExamName = ""
'    strExamCode = Trim(asExamCode)
'
'    SQL = "SELECT EXAMNAME FROM EQUIPEXAM WHERE EXAMCODE = '" & strExamCode & "'"
'    Res = db_select_Col(gLocal, SQL)
'
'    If Res < 0 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    Get_ExamName = Trim(gReadBuf(0))
'
'End Function
'
'
'Public Function Get_EquipCode(asExamCode As String) As String
'    Dim strExamCode As String
'    Get_EquipCode = ""
'    strExamCode = Trim(asExamCode)
'
'    SQL = "SELECT EQUIPCODE FROM EQUIPEXAM WHERE EXAMCODE = '" & strExamCode & "'"
'    Res = db_select_Col(gLocal, SQL)
'
'    If Res < 0 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    Get_EquipCode = Trim(gReadBuf(0))
'
'End Function
'
''
''
''
'''======================================================================================
''' Function Name : GetXhwnd
''' DateTime      : 2007-09-06 21:32
''' Author        : �����ƺ� (http://cafe.naver.com/xlsvba/489)
''' Purpose       : �ش� Excel������ �ڵ鰪�� ���Ѵ�.
''' Param         : strFN - �ش������� Path�� ������ �̸�
''' Return        : GetXhwnd - �ش� Excel������ �ڵ鰪
'''======================================================================================
''Public Function GetXhwnd(ByVal strFN As String) As Long
''
''   On Error GoTo GetXhwnd_Err
''        strFN = Replace(strFN, ".xlsx", "", , , vbTextCompare)
''        strFN = Replace(strFN, ".xls", "", , , vbTextCompare)
''        GetXhwnd = FindWindow("XLMAIN", "Microsoft Excel - " & strFN)
''
''GetXhwnd_Err:
''
''   Call Err_Message("GetXhwnd", "modExcel")
''
''End Function
''
''
''
'''=================================================================================
''' Procedure : XlOpen
''' DateTime  : 2007-09-07 08:42
''' Author    : �����ƺ� (http://cafe.naver.com/xlsvba/489)
''' Purpose   : ���� ��ü�� �����ϰ� ���ο� ��ũ���� ����� ��Ʈ�� ������ �����Ѵ�.
''' Param     : bState - ������ ���¼��� True - xlMaximized, False - xlMinimized
'''=================================================================================
''Public Sub XlOpen(ByVal XFile As String, Optional bState As Boolean = False)
''  Dim lngState As Long
''  Dim hxls     As Long
''  Dim strBuf   As String
''  Dim strFN    As String
''
''    On Error GoTo XlOpen_Err
''        '## ���ڿ� ���� Ȯ��
''        strBuf = String(128, 0)
''
''        '## Ȯ���ڸ� ������ ���ϸ� ���
''        Call GetFileTitle(XFile, strBuf, Len(strBuf))
''
''        '## Chr(0)���ڿ� ����
''        strFN = Replace(strBuf, Chr(0), "", , , 1)
''
''        '## �̹� ����Ʈ�� ���µǾ����� Ȯ���ϱ�
''        hxls = GetXhwnd(strFN)
''
''        '## ���� ���¶�� ���μ����� ����
''        If hxls > 0 Then Call Process_Kill(hxls)
''
''        '## ���� ��ü �ν��Ͻ� ����
''        Set xlapp = CreateObject("Excel.Application")
''
''        '## ������ ���� �μ� ���
''        lngState = -4140                            '## IIf(bState = True, -4137, -4140)
''
''        With xlapp
''            .WindowState = lngState                 '## ������ ����
''            .Visible = False                        '## VIsible
''            .EnableAnimations = False               '## Animationȿ�� ����
''            If bState Then
''                .Workbooks.Open XFile
''            Else
''                .Workbooks.Open "�������ϸ�"        '## ���� ȣ���ϴ� ���� ���� ����
''            End If
''        End With
''
''XlOpen_Err:
''
''   Call Err_Message("XlOpen", "modExcel")
''
''End Sub
''
''
'''=================================================================================
''' Procedure : SetXlsPicture
''' DateTime  : 2007-09-07 09:24
''' Author    : �����ƺ� (http://cafe.naver.com/xlsvba/489)
''' Purpose   : �ش� ��ũ��Ʈ�� ������ �Է��Ѵ�.
''' Param     : strShtName - ������ �Էµ� ��Ʈ��
'''=================================================================================
''Public Sub SetXlsPicture(ByVal strShtName As String, strImg As String)
''  Dim picInsa  As Variant
''  Dim rngX     As Variant
''  Dim Pic      As Variant
''  Dim strImage As String
''
''
''    On Error GoTo SetXlsPicture_Err
''
''    With xlapp.Workbooks(XName).Worksheets(strShtName)
''        xlapp.Visible = True
''        '## ��ũ��Ʈ���� Ÿ��Ʋ Shape�� ������ ��� Picture ��ü�� �����Ѵ�.
''        For Each Pic In .Shapes
''            If Left(Pic.Name, 4) <> "picImg" Then
''                Pic.Delete
''            End If
''        Next Pic
''        .Pictures.Delete
''
''        '## ������ ���� Range���� ����(�ѹ����� ���Ǵ°��̶� Range�κ��� �ϵ��ڵ��Ͽ���)
''        Set rngX = .Range("D11:J18")
''
''        '## �ҷ��� ������ FullName�� ���
''        strImage = strImg
''
''        '## �ش� ������ ��ũ��Ʈ�� �ø���.
''        Set picInsa = .Pictures.Insert(strImg)
''
''    End With
''
''    '## ������ ��ġ�� ������ ����
''    With picInsa
''        .Left = rngX.Left + 2
''        .Top = rngX.Top + 2
''        .Width = rngX.Width - 4
''        .Height = (rngX.Height) - 4
''    End With
''
''SetXlsPicture_Err:
''
''   Call Err_Message("SetXlsPicture", "modExcel")
''
''End Sub
'
'
''
'''=======================================================================================
''' Procedure     : Process_Kill
''' Description   : API�� �ش� �������μ��� �ְܹ�����
''' DateTime      : 2007-09-06 20:56
''' Author        : �����ƺ� (http://cafe.naver.com/xlsvba/489)
''' Parameter     : strFN   : ������ ������������ ĸ�Ǹ�(Ȯ���ڸ� ������ ���ϸ�)
'''=======================================================================================
''Public Sub Process_Kill(ByVal hxls As Long)
''  Dim hProcess As Long
''  Dim pID      As Long
''  Dim tID      As Long
''
''    On Error GoTo Process_Kill_Err
''
''        '## ����� �ڵ鰪���� ProcessId ���
''        tID = GetWindowThreadProcessId(hxls, pID)
''
''        '## �ش� ���μ��� ����
''        hProcess = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, 0&, pID)
''
''        '## ���μ��� ����
''        Call TerminateProcess(hProcess, 0&)
''        '## OpenProcess�� �ݵ�� ��������...
''        Call CloseHandle(hProcess)
''
''Process_Kill_Err:
''
''    Call Err_Message("Process_Kill", "modExcel")
''
''End Sub
''
''
'''=======================================================================================
''' Procedure   : SetXicon
''' Description : �������� ���� ��Ű
''' Author      : �����ƺ�(http://cafe.naver.com/xlsvba/489)
''' DateTime    : 2007-09-07 11:58
''' Parameter   : strIconPath - ���������� FullName
'''=======================================================================================
''Public Sub SetXicon(ByVal hxls As Long, Optional strIconPath As String = "C:\Program Files\Microsoft Office\OFFICE11\FORMS\1042\CONTACTS.ICO")
''   Dim SHinfo    As SHFILEINFO
''   Dim iconHwnd  As Long
''
''     If hxls = 0 Then Exit Sub
''
''     '## ���������� ���������� ����ü ���·� ���
''     Call SHGetFileInfo(strIconPath, 0&, SHinfo, Len(SHinfo), SHGFI_ICON)
''
''     '## ������ ������ �ڵ� ���
''     iconHwnd = SHinfo.hIcon
''
''     If iconHwnd <= 1 Then Exit Sub
''
''     '## ����
''     Call SendMessage(hxls, WM_SETICON, True, iconHwnd)
''
''     Call SendMessage(hxls, WM_SETICON, False, iconHwnd)
''
''End Sub
'''=======================================================================================
''' Procedure   : SetCaption
''' Description : ���� �����츦 ĸ�ǹ� ���� ������� ����
''' Author      : �����ƺ�(http://cafe.naver.com/xlsvba/489)
''' DateTime    : 2007-09-07 12:52
''' Parameter   : hxls - �ش������� �ڵ�
'''=======================================================================================
''Public Sub SetCaption(ByVal hxls As Long)
''  Dim Ret As Long
''
''    Ret = GetWindowLong(hxls, GWL_STYLE)                    '## ���� ������ ������ �Ӽ� ���
''
''    Ret = Ret And Not (WS_CAPTION)                          '## ĸ�ǹ� �������·� ����
''
''    Call SetWindowLong(hxls, GWL_STYLE, Ret)                '## �������·� �Ӽ�����
''
''End Sub
''
'''=======================================================================
''' Function : TrimString
''' Author   : S.J Lee(http://cafe.naver.com/xlsvba/489)
''' LA Time  : 2007-09-13 10:28
''' Purpose  : ���鹮�ڿ� ����
''' Return   : ���ŵ� ���ڿ�
''' Param    : ������ ���ڿ�
'''=======================================================================
''Public Function TrimString(strString As String) As String
''    TrimString = Left$(strString, lstrlenW(StrPtr(strString)))
''End Function
''
''
'''================================================================================
''' Procedure   : Err_Message
''' DateTime    : 2007-09-07 08:05
''' Author      : �� ����(http://cafe.naver.com/xlsvba/489)
''' Purpose     : �����޼��� ó��(�����ġ, ���ν���(�Լ�)�� ����
''' Param       : strFuncName - �ش� �Լ��� ���ν�����
'''             : strModuleName - ���ν����� �����ϰ� �ִ� ����
'''================================================================================
''Public Sub Err_Message(ByVal strFuncName As String, ByVal strModuleName As String)
''
''    If Err.Number <> 0 Then
''        MsgBox "������ �߻��Ͽ����ϴ�." & vbCrLf & _
''               "������ ������ " & Err.Description & vbCrLf & _
''               "������ ��ġ�� Function(Procedure) : " & strFuncName & "Module : " & strModuleName, vbCritical
''    End If
''
''    On Error GoTo 0
''End Sub
''
''
'''=============================================================================
''' Function : SetHiddenCtl
''' Author   : �̼���(http://cafe.naver.com/xlsvba/489)
''' LA Time  : 2007-11-02 09:50
''' Purpose  : ���� ��ũ��Ʈ���� ��Ʈ�Ѱ� �޴����� Hidden
''' Param    : objXl - ������ ����App
'''=============================================================================
''Public Sub SetHiddenCtl(objXl As Object)
''  Dim DelBar As Variant
''
''    For Each DelBar In objXl.CommandBars
''        If DelBar.BuiltIn Then
''            DelBar.Enabled = False
''        Else
''            DelBar.Delete
''        End If
''    Next DelBar
''
''    With objXl
''        .ShowStartupDialog = False
''        .DisplayFormulaBar = False
''        .DisplayStatusBar = False
''        .ShowWindowsInTaskbar = False
''    End With
''End Sub
''
'''=============================================================================
''' Function : ConnectADO
''' Author   : �̼���(http://cafe.naver.com/xlsvba/489)
''' LA Time  : 2007-10-17 14:05
''' Purpose  : ADO�� �̿��Ͽ� DB Connect
''' Return   : ��������
''' Param    : strDB - DB��
'''            intDB (Excel - 1, MDB - 2 , MSSQL - 3, oracle - 4)
'''            strServer - ������
'''            strId - ID
'''            strPw - PassWord
'''=============================================================================
''Public Function ConnectADO(ByVal strDB As String, _
''                           Optional intDB As Integer = 1, Optional strServer As String = "", _
''                           Optional strId As String = "", Optional strPw As String = "") As Boolean
''
''  On Error GoTo ConnectADO_Err
''
''    ConnectADO = True
''
''    Select Case intDB
''
''        Case 1: strCn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
''                        "Data Source=" & strDB & ";" & "Extended Properties=Excel 8.0;"
''        Case 2: strCn = "Provider=Microsoft.JET.OLEDB.4.0;data source= " & strDB & _
''                        ";Jet OLEDB:Database Password=" & strPw & ""
''        Case 3: strCn = "Driver={SQL Server};server=" & strServer & _
''                        ";Database=" & strDB & " & "", " & strId & ", " & strPw & ""
''        Case 4: strCn = "Provider=MSDAORA;Data Source=" & strDB & _
''                        ";User ID=" & strId & ";Password=" & strPw & ""
''        Case 5: strCn = "DSN=HDSEM;User ID=**********;password=********;database=********;"
''
''    End Select
''
''
''    Set xCn = New ADODB.Connection
''
''    xCn.ConnectionString = strCn
''
''    xCn.Open strCn
''
''ConnectADO_Err:
''
''    If Err.Number <> 0 Then
''        ConnectADO = False
''        Call Err_Message("ConnectADO", "modExcel")
''    End If
''
''End Function
''
'''=======================================================================
''' Function : RemoveXlb
''' Author   : �̼���(http://cafe.naver.com/xlsvba/489)
''' LA Time  : 2007-10-24 16:50
''' Purpose  : �������μ����� �ν��Ͻ� ������ ���� �а� �����ϴ� ���Ϸ�
'''            ������������ ����ִ� .xlb ������ �����Ѵ�.
'''            Excel Report�� ���������Ŀ� ��â ���½� �����αװ� �������� �ʴ´�.
'''            ����ڰ� "�ɼ�"���� Path�� �����Ҽ��� �ִ�.
'''            SearchTreeForFile()���� �Լ��� ã���� ������ ĳ�û���������
'''            ���α׷��� �ӵ����ϸ� �����ϹǷ� �̴� �������� ������~~
'''=======================================================================
''Public Sub RemoveXlb()
''  Dim strSP As String
''  Dim strDF As String
''  Dim FSO   As Object
''
''  On Error Resume Next
''
''    Call GetSPfolder(strSP, 26, True)               '## ~~~~\UserName\Application Data ��� ���
''
''    strSP = strSP & "\Microsoft\Excel\excel*.xlb"   '## xlb������ FullName�� ������ �Ҵ�
''
''    Set FSO = CreateObject("Scripting.FileSystemObject")
''
''    FSO.DeleteFile strSP                            '## ����
''
''    Set FSO = Nothing                               '## �޸� ����
''
''  On Error GoTo 0
''
'' End Sub
''
'''=======================================================================
''' Function : GetSPfolder
''' Author   : S.J Lee(http://cafe.naver.com/xlsvba/489)
''' LA Time  : 2007-09-13 10:31
''' Purpose  : Ư������� Path��� Ȥ�� �ش� �����츦 ����
''' Return   : ��ΰ� ����(Call by Reference)
''' Param    : strSP - ��ΰ�, CSIDL - Ư����� �μ�,
'''            BLK - True(��ΰ� ���ڿ��� ����), False(�ش� ������ ����)
'''=======================================================================
''Public Function GetSPfolder(ByRef strSP As String, ByVal CSIDL As Long, hxls As Long, Optional BLK As Boolean = True) As Boolean
''    Dim IDL    As ITEMIDLIST
''    Dim SEI    As SHELLEXECUTEINFO
''    Dim lngRet As Long
''    Dim strBuf As String
''
''    '## Ư����ΰ� ���(IDL)
''    lngRet = SHGetSpecialFolderLocation(hxls, CSIDL, IDL)
''
''    If lngRet = 0 Then              '## ��ΰ��� ���ڿ��� ����Ѵٸ�
''
''        '## ���ڿ� ���� buffer ����
''        strBuf = String(512, Chr(0))
''
''        '## IDList�� ���� ��ΰ� ���
''        Call SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strBuf)
''
''        '## Chr(0) ���ڿ� ����
''        strSP = TrimString(strBuf)
''
''    ElseIf lngRet = 0 And BLK = False Then   '## �ش� ��θ� ������� �����Ѵٸ�
''        '## �ش� ����ü ������ ���
''        With SEI
''            .cbSize = Len(SEI)
''            .hwnd = hxls
''            .lpVerb = "Open"
''            .lpFile = vbNullString
''            .lpParameters = vbNullString
''            .lpDirectory = vbNullString
''            .nShow = 1
''            .fMask = SEE_MASK_IDLIST
''            .lpIDList = IDL.mkid.cb
''        End With
''
''        '## ����
''        Call ShellExecuteEX(SEI)
''
''    End If
''
''End Function
'
'
''���� ������ �׸��忡 �ֱ�
'Private Sub Excel_Open()
'    Dim xlApp   As New Excel.Application
'    Dim XLappWS As Worksheet
'    Dim lngSCnt As Long
'    Dim lngSColCnt(100) As Long
'    Dim dummy       As String
'    Dim strRowData  As Variant
'    Dim lngRowCnt   As Long
'    Dim chk_str     As String
'    Dim dummy_max   As Long
'    Dim lngTotColCnt   As Long
'    Dim lngTotRowCnt   As Long
'    Dim i, j, k     As Long
'
'
'
''Dim xlapp As New Excel.Application
''Dim xlapp_worksheet As Worksheet
''Dim sheet_count As Long
''Dim sheet_col_count(100) As Long
''Dim i, j, k As Long
''Dim dummy As String
''Dim row_data As Variant
''Dim row_cnt As Long
''Dim chk_str As String
''Dim dummy_max As Long
''Dim tot_col_count As Long
''Dim tot_row_count As Long
'
'    lngTotColCnt = 0
'    lngTotRowCnt = 0
'
'
'    '���� ����
'    CommonDialog1.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls"
'    CommonDialog1.Action = 1
'
'
'    If CommonDialog1.FileTitle = "" Then
'        Exit Sub
'    End If
'
'    xlApp.Workbooks.Open (Trim(CommonDialog1.FileName))
'
'    lngSCnt = xlApp.Worksheets.Count
'
'    '-- ��ü ��ũ��Ʈ �ҷ�����ͼ� '�ӽ�.txt' ���Ϸ� ����
'    For i = 1 To lngSCnt
'        Set XLappWS = xlApp.Worksheets(i)
'        XLappWS.Activate
'        lngSColCnt(i) = XLappWS.UsedRange.Columns.Count
'        xlApp.DisplayAlerts = False
'
'        '''xlApp.ActiveWorkbook.SaveAs App.Path & "\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 ����
'        xlApp.ActiveWorkbook.SaveAs "C:\CFX_EXCEL\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 ����
'
'
'        'XLappWS.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False ==>���� 2000��
'        'ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False  ===>���� 2003��
'    Next i
'
'    xlApp.Quit
'    Set XLappWS = Nothing
'    Set xlApp = Nothing
'
'    '-- ��ü ������ MAX cols�� ����
'    dummy_max = 0
'    For i = 1 To lngSCnt
'        If lngSColCnt(i) >= dummy_max Then
'            dummy_max = lngSColCnt(i)
'        End If
'    Next i
'    lngTotColCnt = dummy_max
'
'    '��ü row�� ����
'    For i = 1 To lngSCnt
''''        Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
'        Open ("C:\CFX_EXCEL\" & Trim(i) & ".txt") For Input As #1
'        While Not EOF(1)
'            Line Input #1, dummy
'            strRowData = Split(Trim(dummy), Chr(9))
'            chk_str = ""
'            For j = 0 To UBound(strRowData)
'                chk_str = chk_str & strRowData(j)
'            Next j
'            If Len(Trim(dummy)) > 0 Then
'                lngTotRowCnt = lngTotRowCnt + 1
'            End If
'        Wend
'        Close #1
'    Next i
'
'    '-- �׸��� �ʱ�ȭ
'    vasExcel.MaxRows = 0
'    vasExcel.MaxRows = lngTotRowCnt
'    vasExcel.MaxCols = lngTotColCnt
'
'    '-- �׸��忡 ���
'    For i = 1 To lngSCnt
'        '''Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
'        Open ("C:\CFX_EXCEL\" & Trim(i) & ".txt") For Input As #1
'        While Not EOF(1)
'            Line Input #1, dummy
'            strRowData = Split(Trim(dummy), Chr(9))
'            chk_str = ""
'            For j = 0 To UBound(strRowData)
'                chk_str = chk_str & strRowData(j)
'            Next j
'            If Len(chk_str) > 0 Then
'                lngRowCnt = lngRowCnt + 1
'                For j = 0 To UBound(strRowData)
'                    Call vasExcel.SetText(j + 1, lngRowCnt, CStr(strRowData(j)))
'                Next j
'            End If
'        Wend
'        Close #1
'    Next i
'
''    Call SpreadSheetSort(vasExcel, 6, 2)
'    With vasExcel
'        .Col = 1: .Col2 = .MaxCols
'        .Row = 2: .Row2 = .DataRowCnt
'        .SortBy = 0
'        .SortKey(1) = 6       '����Ű ����ȣ
'        .SortKey(2) = 2       '����Ű ����ȣ
'
'        .SortKeyOrder(1) = SortKeyOrderAscending
'        .SortKeyOrder(2) = SortKeyOrderAscending
'
'        .Action = ActionSort
'    End With
'
'
''Dim SortKeys, SortKeyOrder As Variant
''
''    SortKeys = Array(6, 2)
''    SortKeyOrder = Array(6, 2)
''    ' Sort data in first five columns and rows by column 1 and 3
''    vasExcel.Sort 6, 2, 2, vasExcel.MaxRows, SS_SORT_BY_ROW, SortKeys, SortKeyOrder
'
'End Sub
'
'
''Private Sub cmdExcelFind_Click()
''    Dim sSeq As String
''    Dim sBarcode As String
''    Dim strEqpResult As String
''    Dim strLisResult As String
''    Dim strIntBase As String
''    Dim lsExamCode As String
''    Dim lsExamName As String
''    Dim lsSeqNo As String
''    Dim lsResRow    As String
''    Dim lsEquipRes As String
''    Dim lsResult_Buff As String
''
''    Dim lRow As Integer
''    Dim lRow1 As Integer
''    Dim intRow As Integer
''    Dim sWellOld As String
''    Dim sWell As String
''    Dim sExamCode As String
''    Dim sExamName As String
''    Dim sEquipCode As String
''    Dim sItemCode As String
''    Dim strAge As String
''    Dim strSex As String
''    Dim strPtno As String
''    Dim strPtname As String
''    Dim varTmp As Variant
''    Dim intTstCnt As Integer
''
''    Call Excel_Open
''
''    intTstCnt = 0
''
''    With vasExcel
''        For intRow = 2 To .DataRowCnt
''            .GetText 6, intRow, varTmp: sSeq = varTmp
''            If sSeq <> "" Then
''                '-- ���� Seq ã��
''                If gMode = "0" Then
''                    lRow = SeqSearch(spdOrder, sSeq, colSeqNo)
''                Else
''                    lRow = SeqSearch(spdTot, sSeq, colSeqNo)
''                End If
''
''                If lRow >= 1 Then
''                    '-- ȯ������
''                    If gMode = "0" Then
''                        spdOrder.GetText colBarcode, lRow, varTmp: sBarcode = varTmp
''                        Call SetPatInfo(sBarcode)
''                        SetText spdOrder, "Result", lRow, colState
''                    Else
''                        spdTot.GetText colBarcode, lRow, varTmp: sBarcode = varTmp
''                        Call SetPatInfo(sBarcode)
''                        SetText spdTot, "Result", lRow, colState
''                    End If
''
''                    '-- Well (Position)
'''                    .GetText 2, intRow, varTmp: sWell = varTmp
'''
'''                    If gMode = "0" Then
'''                        Call spdOrder.SetText(colWell, lRow, sWell)
'''                    Else
'''                        spdTot.GetText colWell, lRow, varTmp: sWellOld = varTmp
'''                        If sWellOld <> "" And sWellOld <> sWell Then
'''                            sWell = sWellOld & "/" & sWell
'''                        End If
'''                        If Len(sWellOld) < 6 Then
'''                            Call spdTot.SetText(colWell, lRow, sWell)
'''                        Else
'''                            Call spdTot.SetText(colWell, lRow, sWellOld)
'''                        End If
'''                    End If
''
''                    '-- ä��
''                    .GetText 3, intRow, varTmp: strIntBase = varTmp
''
''                    If intTstCnt < 3 Then
''                        strIntBase = strIntBase + "L"
''                    Else
''                        strIntBase = strIntBase + "H"
''                    End If
''
''                    intTstCnt = intTstCnt + 1
''                    If intTstCnt = 6 Then intTstCnt = 0
''                    '-- ���
''                    .GetText 7, intRow, varTmp: strEqpResult = varTmp
''                    If UCase(Mid(strIntBase, 1, 3)) = "CY5" Then
''                        strLisResult = strEqpResult
''                    Else
''                        If strEqpResult = "" Then
''                            strLisResult = "Negative" & strEqpResult
''                        ElseIf strEqpResult > 38 Then
''                            strLisResult = "Negative" & strEqpResult
''                        ElseIf strEqpResult < 38 Then
''                            strLisResult = "Positive" & " (" & strEqpResult & ")"
''                        Else
''                            strLisResult = "Negative" & strEqpResult
''                        End If
''                    End If
''
''                    If strLisResult <> "" Then
''                              SQL = "Select examcode, examname, seqno "
''                        SQL = SQL & "  From equipexam"
''                        SQL = SQL & " Where equipno = '" & gEquip & "' "
''                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
''                        SQL = SQL & "   and examcode in (" & gOrderExam & ") "      '"'36721','36722','36723','36724'"
''                        Res = db_select_Col(gLocal, SQL)
''
''                        If Res > 0 Then
''                            lsExamCode = Trim(gReadBuf(0))
''                            lsExamName = Trim(gReadBuf(1))
''                            lsSeqNo = Trim(gReadBuf(2))
''
''                            lsResRow = spdResult.DataRowCnt + 1
''                            If spdResult.MaxRows < lsResRow Then
''                                spdResult.MaxRows = lsResRow
''                            End If
''
''                            '�Ҽ��� ó��, ��� ���� ó��
''                            lsEquipRes = strLisResult
''                            strLisResult = SetResult(strLisResult, strIntBase)
''                            lsResult_Buff = strLisResult
''
''                            If gMode = "0" Then
''                                SetText spdResult, strIntBase, lsResRow, colChannel       '����ڵ�
''                                SetText spdResult, lsExamCode, lsResRow, colTestCd        '�˻��ڵ�
''                                SetText spdResult, lsExamName, lsResRow, colTestNm        '�˻��
''                                SetText spdResult, strEqpResult, lsResRow, colEqpResult           '�����
''                                SetText spdResult, strLisResult, lsResRow, colLisResult           'LIS���
''                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
''                            Else
''                                SetText spdTot, strLisResult, lRow, colState + lsSeqNo         'LIS���
''
''                                SetLocalDBTot lRow, strIntBase, lsExamCode, lsExamName, strEqpResult, strLisResult, lsSeqNo
''
''                            End If
''
''                            lsResult_Buff = ""
''
''                        Else
''                            '-- ���� ���� ���
''                                  SQL = "Select examcode, examname, seqno "
''                            SQL = SQL & "  From equipexam"
''                            SQL = SQL & " Where equipno = '" & gEquip & "' "
''                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
''                            Res = db_select_Col(gLocal, SQL)
''
''                            If Res > 0 Then
''                                lsExamCode = Trim(gReadBuf(0))
''                                lsExamName = Trim(gReadBuf(1))
''                                lsSeqNo = Trim(gReadBuf(2))
''
''                                lsResRow = spdResult.DataRowCnt + 1
''                                If spdResult.MaxRows < lsResRow Then
''                                    spdResult.MaxRows = lsResRow
''                                End If
''
''                                '�Ҽ��� ó��, ��� ���� ó��
''                                lsEquipRes = strLisResult
''                                strLisResult = SetResult(strLisResult, strIntBase)
''                                lsResult_Buff = strLisResult
''
''                                If gMode = "0" Then
''                                    SetText spdResult, strIntBase, lsResRow, colChannel       '����ڵ�
''                                    SetText spdResult, lsExamCode, lsResRow, colTestCd        '�˻��ڵ�
''                                    SetText spdResult, lsExamName, lsResRow, colTestNm        '�˻��
''                                    SetText spdResult, strEqpResult, lsResRow, colEqpResult           '�����
''                                    SetText spdResult, strLisResult, lsResRow, colLisResult           'LIS���
''                                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
''                                Else
'''                                    SetText spdTot, strIntBase, lsResRow, colChannel       '����ڵ�
'''                                    SetText spdTot, lsExamCode, lsResRow, colTestCd        '�˻��ڵ�
'''                                    SetText spdTot, lsExamName, lsResRow, colTestNm        '�˻��
'''                                    SetText spdTot, strEqpResult, lsResRow, colEqpResult           '�����
''                                    SetText spdTot, strLisResult, lRow, colState + lsSeqNo         'LIS���
''
''                                    SetLocalDBTot lRow, strIntBase, lsExamCode, lsExamName, strEqpResult, strLisResult, lsSeqNo
''                                End If
''                                lsResult_Buff = ""
''                                strState = ""
''                            End If
''                        End If
''                    End If
''
''                    strState = "R"
''                Else
'''                    If gMode = "0" Then
'''                        spdOrder.MaxRows = spdOrder.MaxRows + 1
'''                        spdOrder.RowHeight(-1) = 12
'''                        lRow = spdOrder.MaxRows
'''
'''                        SetText spdOrder, sSeq, lRow, colBarcode  'colSeqNo
'''                        SetText spdOrder, "Result", lRow, colState
'''                    Else
'''                        spdTot.MaxRows = spdTot.MaxRows + 1
'''                        spdTot.RowHeight(-1) = 12
'''                        lRow = spdTot.MaxRows
'''
'''                        SetText spdTot, sSeq, lRow, colSeqNo
'''                        SetText spdTot, "Result", lRow, colState
'''                    End If
''                End If
''            End If
''        Next
''    End With
'''    CommonDialog1.InitDir = App.Path & "\"
'''    CommonDialog1.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls"
'''    CommonDialog1.Action = 1
'''
'''
'''    strExcel = CommonDialog1.FileName
'''
'''
'''   ' Call XlOpen(strExcel, False)
'''
'''    FileName = strExcel
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''vasExcel.ScriptEnhanced = True
'''
'''    If Right(FileName, 3) = "xls" Then   '// Excel 2003
'''        x = vasExcel.IsExcelFile(FileName)
'''        strType = "2003��"
'''    ElseIf Right(FileName, 4) = "xlsx" Then
'''        x = vasExcel.IsExcel2007File(FileName)
'''
'''        'X = vasExcel.IsExcel2007File(Mid(FileName, 1, Len(FileName) - 1))
'''        strType = "2007��"
'''    End If
'''
'''
'''  '//Check if file is an Excel file and set result to x
'''
'''         '//If file is Excel file, tell user, import sheet
'''             '//list, and set result to y
'''            If x = 1 Then
'''
'''                If strType = "2007" Then
'''                    y = vasExcel.OpenExcel2007File(FileName, "", -1, 0, App.Path & "\" & "ILOGFILE22.TXT")
'''                Else
'''                    y = vasExcel.GetExcelSheetList(FileName, List, listcount, App.Path & "\" & "Report.txt", handle, True)
'''                End If
'''
'''                Debug.Print y
'''                '//If received sheet list, tell user, import file,
'''                '//and set result to z
'''                If y = True Then
'''
'''
'''                     z = vasExcel.ImportExcelSheet(handle, 0)
'''
'''
'''                    '//Tell user result based on T/F value of z
'''                    If z = True Then
'''
'''                       Dim rowcount, colcount As Integer
'''
'''                        '//Return the last row that contains data
'''                        rowcount = vasExcel.DataRowCnt
'''
'''                        '//// Import Cell Row Count Check
'''                        Dim nSpreadInputCnt As Integer
'''
'''
'''                       ''''' nSpreadInputCnt = 10000 ? fpSpread1.DataRowCnt
'''                        If nSpreadInputCnt < rowcount Then
'''                              MsgBox "????? ??? ????? ???????.", , "Result"
'''                        Else
'''                             '//Return the last column that contains data
'''                             colcount = vasExcel.DataColCnt
'''
'''                             vasExcel.AllowMultiBlocks = True
'''                             vasExcel.SetSelection 1, 1, colcount, rowcount
'''
'''                             vasExcel.ClipboardCopy
'''
''''                             fpSpread1.SetFocus
'''
''''                             fpSpread1.ClipboardPaste
'''                             '//MsgBox "Import complete.", , "Result"
'''                        End If
'''
'''                    Else
'''                        MsgBox "?? ???? ???? ?? ? ??? ??????.", , "Result"
'''                    End If
'''                Else
'''                    '//Tell user cannot obtain sheet list
'''                    MsgBox "?? ???? ???? ?? ? ????.", , "Result"
'''                End If
'''            Else
'''                '//Tell user file is not Excel file or is locked
'''                MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
'''            End If
'''
'''
''''    Pic_LoadingBar.Visible = False
'''
''
''
''
''Exit Sub
'
''    Dim X As Integer, Y As Boolean, z As Boolean
''    Dim ListCount As Integer, handle As Integer
''    Dim List(10) As String
''    Dim intRow, intCol As Long
''    Dim varTmp As Variant
''    Dim strExcel As String
''    Dim idates1$, idates2$, iexamcode$
''    Dim pt_no$(), patname$(), Sex$(), Age$()
''    Dim spc_no$(), gnl_item_cd$(), bl_gth_dte$()
''    Dim dept$(), wd_no$(), tst_cd$()
''    Dim rv As Integer
''    Dim lRow As Integer
''    Dim lRow1 As Integer
''    Dim sWell As String
''    Dim sExamCode As String
''    Dim sExamName As String
''    Dim sEquipCode As String
''    Dim sItemCode As String
''    Dim strAge As String
''    Dim strSex As String
''    Dim strPtno As String
''    Dim strPtname As String
''
''    Dim ispcno$
''    Dim strTmp As String
''
''    Dim sSeq As String
''    Dim sBarcode As String
''    Dim strEqpResult As String
''    Dim strLisResult As String
''    Dim strIntBase As String
''    Dim lsExamCode As String
''    Dim lsExamName As String
''    Dim lsSeqNo As String
''    Dim lsResRow    As String
''    Dim lsEquipRes As String
''    Dim lsResult_Buff As String
''
''    Dim sFile As String
''    sFile = ShowOpen("Excel Files (*.xls)|*.xls|All Files (*.*)|*.*", App.Path)
''    If sFile <> "" Then
''        'spdOrder.MaxRows = 0
''
''        strExcel = sFile
''
''        vasExcel.ScriptEnhanced = True
''        X = vasExcel.IsExcelFile(strExcel)
''        If X = 1 Then
''            'Try
''                Y = vasExcel.GetExcelSheetList(strExcel, List, ListCount, "Report.txt", handle, True)
''
''
''
''                If Y = True Then
''                    z = vasExcel.ImportExcelSheet(handle, 0)
''                    If z = True Then
''                        'MsgBox "�������� ����"
''                    Else
''                        'MsgBox "�������� ����"
''                    End If
''                Else
''                    'MsgBox "���� ���Ͽ��� �����͸� ���� �� �����ϴ�."
''                End If
''            'Catch ex As Exception
''            '    MessageBox.Show(ex.Message, "��ȯ����", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
''            'End Try
''        Else
''            'MessageBox.Show ("���� ���Ͽ��� �����͸� ���� �� �����ϴ�.")
''        End If
''
''        With vasExcel
''            For intRow = 2 To .DataRowCnt
''                For intCol = 1 To 5
''                    .GetText intCol, intRow, varTmp: sSeq = varTmp
''                    sSeq = "1"
''                    If varTmp <> "" Then
''                        Select Case intCol
''                        Case 1
''                            'ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
''                            'ispcno$ , pt_no$(), patname$(), Sex$(), Age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$()
'''                            rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), Sex$(), Age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
''
''                            '-- ���� Seq ã��
''                            lRow = SeqSearch(spdOrder, sSeq, colSeqNo)
''
''                            If lRow >= 1 Then
''                                '-- ȯ������
''                                spdOrder.GetText colBarcode, lRow, varTmp: sBarcode = varTmp
''                                Call SetPatInfo(sBarcode)
''
''
''                                SetText spdOrder, "Result", lRow, colState
''
''                            Else
''                                spdOrder.MaxRows = spdOrder.MaxRows + 1
''                                spdOrder.RowHeight(-1) = 12
''                                lRow = spdOrder.MaxRows
''
''                                SetText spdOrder, sSeq, lRow, colBarcode  'colSeqNo
''                                SetText spdOrder, "Result", lRow, colState
''
''                            End If
''
''                        Case "4"
''                                '-- Well (Position)
''                                .GetText 4, intRow, varTmp: sWell = varTmp
''                                SetText spdOrder, sWell, lRow, colWell
''
''                        Case "5"
''                                '-- ä��
''                                .GetText 5, intRow, varTmp: strIntBase = varTmp
''
''                                '-- ���
''                                .GetText 6, intRow, varTmp: strEqpResult = varTmp
''                                If strEqpResult = "-" Then
''                                    strLisResult = "Negative"
''                                Else
''                                    strLisResult = "Positive"
''                                End If
''
''                                If strLisResult <> "" Then
''                                          SQL = "Select examcode, examname, seqno "
''                                    SQL = SQL & "  From equipexam"
''                                    SQL = SQL & " Where equipno = '" & gEquip & "' "
''                                    SQL = SQL & "   and equipcode = '" & strIntBase & "' "
''                                    SQL = SQL & "   and examcode in (" & gOrderExam & ") "
''                                    Res = db_select_Col(gLocal, SQL)
''
''                                    If Res > 0 Then
''                                        lsExamCode = Trim(gReadBuf(0))
''                                        lsExamName = Trim(gReadBuf(1))
''                                        lsSeqNo = Trim(gReadBuf(2))
''
''                                        lsResRow = spdResult.DataRowCnt + 1
''                                        If spdResult.MaxRows < lsResRow Then
''                                            spdResult.MaxRows = lsResRow
''                                        End If
''
''                                        '�Ҽ��� ó��, ��� ���� ó��
''                                        lsEquipRes = strLisResult
''                                        strLisResult = SetResult(strLisResult, strIntBase)
''                                        lsResult_Buff = strLisResult
''
''                                        SetText spdResult, strIntBase, lsResRow, colChannel       '����ڵ�
''                                        SetText spdResult, lsExamCode, lsResRow, colTestCd        '�˻��ڵ�
''                                        SetText spdResult, lsExamName, lsResRow, colTestNm        '�˻��
''                                        SetText spdResult, strEqpResult, lsResRow, colEqpResult           '�����
''                                        SetText spdResult, strLisResult, lsResRow, colLisResult           'LIS���
''                                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
''
''                                        lsResult_Buff = ""
''
'''Public Const colRstCheck = 1
'''Public Const colTestNm = 2
'''Public Const colTestCd = 3
'''Public Const colChannel = 4
'''Public Const colEqpResult = 5
'''Public Const colLisResult = 6
'''Public Const colComment = 7
'''Public Const colFlag = 8
'''Public Const colN = 9
'''Public Const colD = 10
'''Public Const colP = 11
'''Public Const colC = 12
'''Public Const colPrevResult = 13
'''Public Const colPrevTestDt = 14
'''Public Const colPrevBarcode = 15
'''Public Const colReference = 16
'''Public Const colOther = 17
''
''
''                                    Else
''                                        '-- ���� ���� ���
''                                              SQL = "Select examcode, examname, seqno "
''                                        SQL = SQL & "  From equipexam"
''                                        SQL = SQL & " Where equipno = '" & gEquip & "' "
''                                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
''                                        Res = db_select_Col(gLocal, SQL)
''
''                                        If Res > 0 Then
''                                            lsExamCode = Trim(gReadBuf(0))
''                                            lsExamName = Trim(gReadBuf(1))
''                                            lsSeqNo = Trim(gReadBuf(2))
''
''                                            lsResRow = spdResult.DataRowCnt + 1
''                                            If spdResult.MaxRows < lsResRow Then
''                                                spdResult.MaxRows = lsResRow
''                                            End If
''
''                                            '�Ҽ��� ó��, ��� ���� ó��
''                                            lsEquipRes = strLisResult
''                                            strLisResult = SetResult(strLisResult, strIntBase)
''                                            lsResult_Buff = strLisResult
''
''                                            SetText spdResult, strIntBase, lsResRow, colChannel       '����ڵ�
''                                            SetText spdResult, lsExamCode, lsResRow, colTestCd        '�˻��ڵ�
''                                            SetText spdResult, lsExamName, lsResRow, colTestNm        '�˻��
''                                            SetText spdResult, strEqpResult, lsResRow, colEqpResult           '�����
''                                            SetText spdResult, strLisResult, lsResRow, colLisResult           'LIS���
''                                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
''
''                                            lsResult_Buff = ""
''                                            strState = ""
''                                        End If
''                                    End If
''                                End If
''
''                                strState = "R"
''                        End Select
''                    End If
''                Next
''            Next
''        End With
''    Else
'''        MsgBox "You pressed cancel"
''    End If
'End Sub
'
'
'Function SetResult(asResult As String, asEquipCode As String)
'    Dim i As Integer
'    Dim sLVal As String
'    Dim sHVal As String
'    Dim sEquipCode As String
'    Dim sEquipRes As String
'    Dim sResult As String
'    Dim sPoint As Integer
'    Dim sResType As String
'    Dim sResFlag As String
'
'
'    sEquipRes = Trim(asResult)
'    sEquipCode = Trim(asEquipCode)
'    sResFlag = ""
'
'    If sEquipCode = "" Then
'        Exit Function
'    End If
'
''    If IsNumeric(sEquipRes) = False Then
''        Exit Function
''    End If
'
'    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
'    Res = db_select_Col(gLocal, SQL)
'
'    If IsNumeric(gReadBuf(0)) = True Then
'        sPoint = CInt(gReadBuf(0))
'        sResType = ""
'        For i = 0 To sPoint
'            If i = 0 Then
'                sResType = "#0"
'            ElseIf i = 1 Then
'                sResType = sResType & ".0"
'            Else
'                sResType = sResType & "0"
'            End If
'        Next
'
'        sResult = Format(sEquipRes, sResType)
'    Else
'        sResult = sEquipRes
'    End If
'
'    SetResult = sResult
'
'End Function
'
'
''Private Sub cmdMode_Click()
''
''    '-- �и���� Ŭ��
''    If cmdMode.Tag = "0" Then
''        spdTot.Visible = False
''        spdOrder.Visible = True
''        spdResult.Visible = True
''        chkResult.Visible = True
''        cmdMode.Caption = "���ո��"
''        cmdMode.Tag = 1
''
''    '-- ���ո�� Ŭ��
''    Else
''        spdTot.Visible = True
''        spdOrder.Visible = False
''        spdResult.Visible = False
''        chkResult.Visible = False
''        cmdMode.Caption = "�и����"
''        cmdMode.Tag = 0
''    End If
''
''
''End Sub
'
'
'
'Private Sub cmdResultSearch_Click()
'    Dim iRow As Long
'    Dim RS As ADODB.Recordset
''    ClearSpread vasRID
''    ClearSpread vasRRes
'
'    SQL = "SELECT '',PID, EXAMDATE, BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, COUNT(*), COUNT(*), '',SENDFLAG " & vbCrLf & _
'          "FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf
'    If cboWhere.ListIndex = 0 Then
'          SQL = SQL & "  AND SENDFLAG IN ('0','1') "
'    ElseIf cboWhere.ListIndex = 0 Then
'          SQL = SQL & "  AND SENDFLAG IN ('1') "
'    ElseIf cboWhere.ListIndex = 0 Then
'          SQL = SQL & "  AND SENDFLAG IN ('0') "
'    End If
'
'    SQL = SQL & "GROUP BY PID, EXAMDATE, BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG "
'    SQL = SQL & " ORDER BY PID * 10"
'
'
'    If gMode = "0" Then
'        Res = GetDBSelectVas(gLocal, SQL, spdOrder)
'    Else
'        Res = GetDBSelectVas(gLocal, SQL, spdTot)
'    End If
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    ElseIf Res = 0 Then
'        Exit Sub
'    End If
'
'
'    For iRow = 1 To spdTot.DataRowCnt
'        Select Case Trim(GetText(spdTot, iRow, colState))
'        Case "2"
'            SetBackColor spdTot, iRow, iRow, 1, colState, 202, 255, 112
'            SetText spdTot, "�Ϸ�", iRow, colState
'        Case "0"
'            SetText spdTot, "���", iRow, colState
'        Case "1"
'            SetText spdTot, "���", iRow, colState
'        End Select
'
'        SQL = "SELECT SEQNO, RESULT " & _
'              "  FROM PAT_RES " & vbCrLf & _
'              " WHERE EXAMDATE = '" & Trim(GetText(spdTot, iRow, colOrdDate)) & "' " & vbCrLf & _
'              "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              "   AND BARCODE =  '" & Trim(GetText(spdTot, iRow, colBarcode)) & "'"
'        SQL = SQL & " Order By SEQNO "
'
'        cn.CursorLocation = adUseClient
'        Set RS = cn.Execute(SQL, , 1)
'
'        Do Until RS.EOF
'            'GetOrderExamCode_New = GetOrderExamCode_New & "'" & Trim(rs_svr.Fields(0)) & "',"
'            SetText spdTot, Trim(RS.Fields("RESULT").Value) & "", iRow, colState + Trim(RS.Fields("SEQNO").Value) & ""
'            RS.MoveNext
'        Loop
'        RS.Close
'
'
'
''    '-- Record Count ������
''    cn.CursorLocation = adUseClient
''    Set RS = cn.Execute(SQL, , 1)
''
''    If RS.RecordCount > 0 Then
''        spdTot.MaxCols = 14 + RS.RecordCount
''        spdResult.MaxRows = RS.RecordCount
''    Else
''        SaveQuery SQL
''        Exit Sub
''    End If
''    i = 14
''    j = 0
''
''    Do Until RS.EOF
''        i = i + 1
''        j = j + 1
''        '-- ������
''        Call spdTot.SetText(i, 0, Trim(RS.Fields("examname").Value) & "")
''        spdTot.ColWidth(i) = 15
''        '-- �и���
''        Call spdResult.SetText(colTestNm, j, Trim(RS.Fields("examname").Value) & "")
''        Call spdResult.SetText(colTestCd, j, Trim(RS.Fields("examcode").Value) & "")
''        Call spdResult.SetText(colReference, j, Trim(RS.Fields("reflow").Value) & "-" & Trim(RS.Fields("refhigh").Value))
''        RS.MoveNext
''    Loop
'
'
'
'
'
'    Next iRow
'
'    spdTot.RowHeight(-1) = 13
'
'End Sub
'
'Private Sub cmdRsltPrint_Click()
'    Dim intRow  As Integer
'    Dim intCol  As Integer
'    Dim varTmp  As Variant
'    Dim strInfA As String
'    Dim strInfB As String
'    Dim strH1N1 As String
'    Dim strH3N2 As String
'
'    spdTot_Print.MaxRows = 0
'    spdTot_Print.MaxRows = spdTot.MaxRows
'    spdTot_Print.MaxCols = spdTot.MaxCols + 1
'    spdTot_Print.ColWidth(spdTot_Print.MaxCols) = 12
'
'    spdTot_Print.SetText spdTot_Print.MaxCols, 0, "confirm"
'
'    With spdTot
'        For intRow = 1 To .MaxRows
'            For intCol = 1 To .MaxCols
'                .GetText intCol, intRow, varTmp
'
'                Select Case intCol
'                    Case colState + 1: .GetText intCol, intRow, varTmp: strInfA = varTmp
'                    Case colState + 2: .GetText intCol, intRow, varTmp: strInfB = varTmp
'                    Case colState + 3: .GetText intCol, intRow, varTmp: strH1N1 = varTmp
'                    Case colState + 4: .GetText intCol, intRow, varTmp: strH3N2 = varTmp
'                End Select
'                spdTot_Print.SetText intCol, intRow, CStr(varTmp)
'            Next
'            '-- �߰�
'            If Mid(strInfA, 1, 1) = "P" And Mid(strInfB, 1, 1) = "P" Then
'                spdTot_Print.SetText spdTot_Print.MaxCols, intRow, "confirm"
'            '-- �߰�
'            ElseIf Mid(strH1N1, 1, 1) = "P" And Mid(strH3N2, 1, 1) = "P" Then
'                spdTot_Print.SetText spdTot_Print.MaxCols, intRow, "confirm"
'            ElseIf Mid(strInfA, 1, 1) = "P" And Mid(strH1N1, 1, 1) = "P" Then
'                spdTot_Print.SetText spdTot_Print.MaxCols, intRow, "Seasonal H1N1"
'            ElseIf Mid(strInfA, 1, 1) = "P" And Mid(strH3N2, 1, 1) = "P" Then
'                spdTot_Print.SetText spdTot_Print.MaxCols, intRow, "Seasonal H3N2"
'            ElseIf Mid(strInfB, 1, 1) = "P" Then
'                spdTot_Print.SetText spdTot_Print.MaxCols, intRow, "Infulenza B"
'            ElseIf Mid(strInfA, 1, 1) = "P" Then
'                spdTot_Print.SetText spdTot_Print.MaxCols, intRow, "SW H1N1"
'            '-- �߰�
'            ElseIf Mid(strInfA, 1, 1) = "N" And Mid(strInfB, 1, 1) = "N" And Mid(strH1N1, 1, 1) = "N" And Mid(strH3N2, 1, 1) = "N" Then
'                spdTot_Print.SetText spdTot_Print.MaxCols, intRow, "Negative"
'            Else
'                spdTot_Print.SetText spdTot_Print.MaxCols, intRow, "confirm"
'            End If
'            strInfA = ""
'            strInfB = ""
'            strH1N1 = ""
'            strH3N2 = ""
'        Next
'    End With
'
'    spdTot_Print.PrintOrientation = PrintOrientationLandscape '�������
'    spdTot_Print.Action = 13
'
'
'
'End Sub
'
'Private Sub cmdSave_Click()
'    Dim lRow As Long
'
'    If gMode = "0" Then
'        For lRow = 1 To spdOrder.DataRowCnt
'            spdOrder.Row = lRow
'            spdOrder.Col = 1
'            If spdOrder.Value = 1 Then
'
'                Res = SaveTransDataW(gRow)
'
'                If Res = -1 Then
'                    SetForeColor spdOrder, lRow, lRow, 1, colState, 255, 0, 0
'                    SetText spdOrder, "Failed", lRow, colState
'                Else
'                    spdOrder.Row = lRow
'                    spdOrder.Col = 1
'                    spdOrder.Value = 1
'
'                    SetBackColor spdOrder, lRow, lRow, 1, colState, 202, 255, 112
'                    SetText spdOrder, "Trans", lRow, colState
'
'                    SQL = " UPDATE PAT_RES SET " & vbCrLf & _
'                          " SENDFLAG = '1' " & vbCrLf & _
'                          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'                          " AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBarcode)) & "' "
'                    Res = SendQuery(gLocal, SQL)
'                    If Res = -1 Then
'                        SaveQuery SQL
'                        Exit Sub
'                    End If
'
'                End If
'                spdOrder.Row = lRow
'                spdOrder.Col = 1
'                spdOrder.Value = 0
'            End If
'        Next lRow
'    Else
'        For lRow = 1 To spdTot.DataRowCnt
'            spdTot.Row = lRow
'            spdTot.Col = 1
'            If spdTot.Value = 1 Then
'
'                Res = SaveTransDataW(lRow)
'
'                If Res = -1 Then
'                    SetForeColor spdTot, lRow, lRow, 1, colState, 255, 0, 0
'                    SetText spdTot, "Failed", lRow, colState
'                Else
'                    spdTot.Row = lRow
'                    spdTot.Col = 1
'                    spdTot.Value = 1
'
'                    SetBackColor spdTot, lRow, lRow, 1, colState, 202, 255, 112
'                    SetText spdTot, "Trans", lRow, colState
'
'                    SQL = " UPDATE PAT_RES SET " & vbCrLf & _
'                          " SENDFLAG = '1' " & vbCrLf & _
'                          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'                          " AND BARCODE = '" & Trim(GetText(spdTot, lRow, colBarcode)) & "' "
'                    Res = SendQuery(gLocal, SQL)
'                    If Res = -1 Then
'                        SaveQuery SQL
'                        Exit Sub
'                    End If
'
'                End If
'                spdTot.Row = lRow
'                spdTot.Col = 1
'                spdTot.Value = 0
'            End If
'        Next lRow
'    End If
'End Sub
'
'
'' asRow1 = Work List
'Function SetLocalDBUP()
'    Dim intRow As Integer
'    Dim intCol As Integer
'    Dim varTmp As Variant
'    Dim strExamDate As String
'    Dim strBarCode  As String
'    Dim strExamName As String
'    Dim strExamCode As String
'    Dim strEqipCode As String
'    Dim strResult   As String
'
'    If gMode = "0" Then
'
'    Else
'        With spdTot
'            For intRow = 1 To .DataRowCnt
'                For intCol = colState + 1 To .MaxCols
'                    .Row = intRow
'                    .Col = intCol
'                    If .BackColor = vbGreen Then
'                        Call .GetText(colOrdDate, intRow, varTmp): strExamDate = varTmp
'                        Call .GetText(colBarcode, intRow, varTmp): strBarCode = varTmp
'                        Call .GetText(intCol, intRow, varTmp): strResult = varTmp
'                        Call .GetText(intCol, 0, varTmp): strExamName = varTmp
'
'                        strExamCode = Get_ExamCode(strExamName)
'                        strEqipCode = Get_EquipCode(strExamCode)
'
'                              SQL = ""
'                              SQL = "UPDATE PAT_RES " & vbCrLf
'                        SQL = SQL & "   Set RESULT    = '" & strResult & "'" & vbCrLf
'                        SQL = SQL & " WHERE EXAMDATE  = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf
'                        SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
'                        SQL = SQL & "   AND BARCODE   = '" & strBarCode & "' " & vbCrLf
'                        SQL = SQL & "   AND EQUIPCODE = '" & strEqipCode & "'" & vbCrLf
'                        SQL = SQL & "   AND EXAMCODE  = '" & strExamCode & "'"
'
'                        Res = SendQuery(gLocal, SQL)
'
'                        If Res = -1 Then
'                            SaveQuery SQL
'                            Exit Function
'                        End If
'
'                        .BackColor = vbWhite
'
'                    End If
'                Next
'            Next
'        End With
'    End If
'
'End Function
'
'Private Sub cmdUpdate_Click()
'
'    Call SetLocalDBUP
'
'End Sub
'
'Private Sub cmdWorkSearch_Click()
'    Dim adoRS2      As ADODB.Recordset
'    Dim strKeyno    As String
'    Dim intRow      As Integer
'    Dim intCnt      As Integer
'
''    spdTot.MaxRows = 10
''    spdTot.SetText 2, 1, "1"
''    Exit Sub
''
'
'    Set adoRS2 = New ADODB.Recordset
'    Set adoRS2 = adoExecQuery50P("SLRTRM50P", Format(dtpFromDt.Value, "yyyymmdd"), gEquipCode, Val(txtFromWN.Text), Val(txtToWN.Text), Space$(5), "0", "")
''    Set adoRS2 = adoExecQuery50P("SLRTRM50P", Format(dtpFromDt.Value, "yyyymmdd"), gEquipCode, Val(txtFromWN.Text), Val(txtToWN.Text), Text1.Text, Text2.Text, "")
'
'    spdOrder.MaxRows = 0
'    spdTot.MaxRows = 0
'
'    If adoRS2.RecordCount <= 0 Then
'        adoRS2.Close: Set adoRS2 = Nothing
'        Exit Sub
'    End If
'
'    If Not adoRS2.EOF Then
'        intRow = 0
'        strKeyno = ""
'        chkOrder.Value = "1"
'
'        Do While Not adoRS2.EOF
'            If gMode = "0" Then
'                With spdOrder
'                    'If InStr(gAllExam, Trim(adoRS2("ITEMCODE"))) > 0 Then
'                        '-- ���ÿ� ��ϵ� �˻��ڵ常 ������
'                        If strKeyno <> adoRS2.Fields("BARCODENO") And InStr(gAllExam, Trim(adoRS2("ITEMCODE"))) > 0 Then
'                            intRow = intRow + 1
'                            If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
'
'                            .SetText colCheckBox, intRow, "1"
'                            .SetText colSeqNo, intRow, Trim(adoRS2.Fields("WORKNO")) 'txtSeq.Text
'                            .SetText colOrdDate, intRow, Trim(adoRS2("ORDDATE"))
'                            .SetText colBarcode, intRow, Trim(adoRS2("BARCODENO"))
'                            .SetText colPID, intRow, Trim(adoRS2.Fields("WORKNO"))
'                            .SetText colPName, intRow, Trim(adoRS2("PNAME"))
'                            '-- �˻��ڵ� Ȯ�� �׽�Ʈ��
'                            '.SetText colState, intRow, Trim(adoRS2("ITEMCODE")) & Trim(adoRS2("ITEMNAME"))
'                            strKeyno = adoRS2("BARCODENO")
'
'                        End If
'                    'End If
'                    'strKeyno = adoRS2("BARCODENO")
'                End With
'            Else
'                With spdTot
'                    'If InStr(gAllExam, Trim(adoRS2("ITEMCODE"))) > 0 Then
'                        '-- ���ÿ� ��ϵ� �˻��ڵ常 ������
'                        If strKeyno <> adoRS2.Fields("BARCODENO") And InStr(gAllExam, Trim(adoRS2("ITEMCODE"))) > 0 Then
'                            intRow = intRow + 1
'                            If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
'
'                            .SetText colCheckBox, intRow, "1"
'                            .SetText colSeqNo, intRow, Trim(adoRS2.Fields("WORKNO")) 'txtSeq.Text
'                            .SetText colOrdDate, intRow, Trim(adoRS2("ORDDATE"))
'                            .SetText colBarcode, intRow, Trim(adoRS2("BARCODENO"))
'                            .SetText colPID, intRow, Trim(adoRS2.Fields("WORKNO"))
'                            .SetText colPName, intRow, Trim(adoRS2("PNAME"))
'                            '-- �˻��ڵ� Ȯ�� �׽�Ʈ��
'                            '.SetText colState, intRow, Trim(adoRS2("ITEMCODE")) & Trim(adoRS2("ITEMNAME"))
'                            strKeyno = adoRS2("BARCODENO")
'
'                        End If
'                    'End If
'                    'strKeyno = adoRS2("BARCODENO")
'                End With
'            End If
'            adoRS2.MoveNext
'        Loop
'
'    End If
'
'End Sub
'
'Private Sub lblclear_Click()
'    lblChangePID.Caption = ""
'    lblChangeBar.Caption = ""
''    lblBarcode(0).Caption = ""
''    lblBarcode(1).Caption = ""
''    lblPname(0).Caption = ""
''    lblPname(1).Caption = ""
'End Sub
'
''Private Sub Command16_Click()
''    Dim i As Long
''    Dim lsChar As String
''
''    strBuffer = ""
''    strBuffer = strBuffer & "1H|\^&||||||||||P||05" & vbCrLf
''    strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||||||||||3B" & vbCrLf
''    strBuffer = strBuffer & "3O|1|11208647111|807^00042^3^^SAMPLE^NORMAL|ALL|R|20111205092128|||||X||||||||||||||O|||||38" & vbCrLf
''    strBuffer = strBuffer & "4R|1|^^^321^^0|>100.0|ng/ml|0.000^4.00|>||F|||20111205092406|20111205094226|CF" & vbCrLf
''    strBuffer = strBuffer & "5C|1|I|51^Above measuring range|I04" & vbCrLf
''    strBuffer = strBuffer & "6R|2|^^^391^^0|13.78|ng/ml|^|N||F|||20111205092448|20111205094308|14" & vbCrLf
''    strBuffer = strBuffer & "7L|140" & vbCrLf
''    strBuffer = strBuffer & "" & vbCrLf
''
''    '-- Seq
''    'strBuffer = "D 000701 6826      01201206196826    E001    24  102    15  003    46  104  7.59  005  4.96  106  0.13  007    56  108   427  009   144  110  95.4  011    47  112  8.97  013  1.08  114  3.32  015   178  116   147  017    48  118  9.66  019  2.59  120  1.08   23   101   24     2   25     8   26     1   27     3  "
''    '-- ���ڵ�
''    strBuffer = "D 000201 039903073000126             E01   226H 02    85H 11   9.1  13   141H 18  13.2  21   0.7  24  20.2H "
''
''    strBuffer = "DERERBDB"
''    strBuffer = "R 003201 0018          1013002058"
''
''    strBuffer = "D 003401 0019          1013002058    E      32   1.4  46    26  26  0.81H 01   130  02  3.32L 03  4.29  04   7.3  05   0.5  06   0.1  07   158  09   124H 10   0.7L 11  11.2  12    57  14    39H 15    47H 16    74H 17   259  19   9.1  21   4.7H "
''
''    strBuffer = "R 000101 00011013002042"
''
''    strBuffer = "D 000101 00011013002042    E012    18  017   129  018    26  "
''
''    Call comEqp_OnComm
''
''
''End Sub
'
'
'Public Sub CtlInitializing()
'
'    '-- �˻�����
'    dtpToday = Format(CDate(Date), "yyyy/mm/dd")
'    '-- ��ȸ����
'    cboWhere.ListIndex = 0
'    '-- ���ã�� �ɼ�
''    optGbn(0).Value = True
'    '-- �۾����� From ~ To (��ũ����Ʈ)
'    dtpFromDt = Format(CDate(Date), "yyyy/mm/dd")
'    dtpToDt = Format(CDate(Date), "yyyy/mm/dd")
'    '-- �۾���ȣ From ~ To (��ũ����Ʈ)
''    txtFromWN = "1"
''    txtToWN = "99999"
''    txtSeq = "1"
'
''    txtBarcode = ""
'
'    cmdMode.Tag = 1
'
'    gRow = 0
'
'End Sub
'
'Public Sub SpdInitializing()
'
'
'    spdTot.MaxRows = 0
'    spdTot_Print.MaxRows = 0
''    spdOrder.MaxRows = 0
''    spdResult.MaxRows = 0
'
'    cmdMode.Tag = 1
'
'    gRow = 0
'
'
'End Sub
'
'
'Private Sub Form_Load()
''    Dim sDate As String
''    Dim i As Integer
''
''On Error GoTo Rst
''
''    If App.PrevInstance Then
''        End
''    End If
''
'''1. Read Ini
'''   - ��������
'''   - ����������
'''   - ���������
'''
''    Call GetSetup
''
''    Call MnMode_Click(CInt(gMode))
''
''    Call MnSave_Click(CInt(gSave))
''
'''    txtBarcode.MaxLength = gBarLen
''
''    '-- ����� ID(�������� �ʿ��� LIS��ü�� ����
''    frmInterface.StatusBar1.Panels(1).Text = gUserID
''
'''2. Spread Set
''    Call SpdInitializing
''
'''3. Control Initial
''
''    Call CtlInitializing
''
'''4. Local DB Open
''
''    If Not Connect_Local Then
''        MsgBox "���� MDB�� ������� �ʾҽ��ϴ�." & vbNewLine & "���α׷��� ����˴ϴ�.", vbCritical, Me.Caption
''        cn_Local_Flag = False
''        'Exit Sub
''        End
''    Else
''        cn_Local_Flag = True
''    End If
'
'
''5. Server DB Open(1)
''   Server DB Open(2)
'
'''    If Not Connect_Server Then
'''        MsgBox "LIS DB�� ������� �ʾҽ��ϴ�." & vbNewLine & "���α׷��� ����˴ϴ�.", vbCritical, Me.Caption
'''        cn_Server_Flag = False
'''        'Exit Sub
'''        End
'''    Else
'''        cn_Server_Flag = True
'''    End If
'
''
''------------------------------------------------------------------------------------------------
''
''Dim adoServerName As String
''Dim adoLoginID  As String
''Dim adoLoginPassword  As String
''Dim adodefaultDatabaseName  As String
''
''Dim lngReturn As Long
''Dim strReturn As String
''Dim strAppName As String
''Dim strFileName As String
''
''
''    'strFileName = App.Path & "\DB_Info.ini"
''    strFileName = App.Path & "\didim.ini"
''    strAppName = "DATABASE"
''
''    adoServerName = Space(256)
''    adoLoginID = Space(256)
''    adoLoginPassword = Space(256)
''    adodefaultDatabaseName = Space(256)
''
''    ' ���� �ҷ�����
''    lngReturn = GetPrivateProfileString(strAppName, "SERVER", strReturn, adoServerName, Len(adoServerName), strFileName)
''    adoServerName = Replace(Trim(adoServerName), Chr(0), "")
''
''    ' ID �ҷ�����
''    lngReturn = GetPrivateProfileString(strAppName, "UID", strReturn, adoLoginID, Len(adoLoginID), strFileName)
''    adoLoginID = Replace(Trim(adoLoginID), Chr(0), "")
''
''    ' PW �ҷ�����
''    lngReturn = GetPrivateProfileString(strAppName, "PWD", strReturn, adoLoginPassword, Len(adoLoginPassword), strFileName)
''    adoLoginPassword = Replace(Trim(adoLoginPassword), Chr(0), "")
''
''    ' DB Name
''    lngReturn = GetPrivateProfileString(strAppName, "DATABASE", strReturn, adodefaultDatabaseName, Len(adodefaultDatabaseName), strFileName)
''    adodefaultDatabaseName = Replace(Trim(adodefaultDatabaseName), Chr(0), "")
''
''    adoConnectSQLServer adoServerName, adoLoginID, adoLoginPassword, adodefaultDatabaseName
'
''------------------------------------------------------------------------------------------------
''
''6. Communication Open
'
''    comEqp.CommPort = gSetup.gPort
''    comEqp.RTSEnable = gSetup.gRTSEnable
''    comEqp.DTREnable = gSetup.gDTREnable
''    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
''
''    If comEqp.PortOpen = False Then
''        comEqp.PortOpen = True
''        frmInterface.StatusBar1.Panels(5).Text = "COM" & comEqp.CommPort & " ���Ἲ��"
''    End If
'
'
''6. Test List Get
''   Server(1) Test Get
''   Server(2) Test Get
'
''    Call GetExamCode
''
'''    Call SetExamName
''
'''7. Error Handling
'''
''
''    '-- ���� ������ ����
''    dtpToday = Date
''    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
''
''    SQL = "delete from pat_res where examdate < '" & sDate & "'"
''    Res = SendQuery(gLocal, SQL)
''
'''    lblUser.Caption = gUserID
''
'''    If lblUser.Caption = "" Then
'''        Call picLogin_Click
'''    End If
''
''
''    '==============================
''    ' ASTM ��Ű��� ���� �ʱ�ȭ
''    intPhase = 1
''    intSndPhase = 0
''    intFrameNo = 1
''    intBufCnt = 0
''    strState = ""
''    blnIsETB = False
''    '==============================
''
''    Exit Sub
''
''Rst:
''    If Err.Number = "8002" Then
''        If (MsgBox("��Ʈ ��ȣ�� �߸��Ǿ����ϴ�." & vbNewLine & vbNewLine & "   ��� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
''            frmInterface.StatusBar1.Panels(5).Text = "COM" & comEqp.CommPort & " �������"
''            Resume Next
''        End If
''    Else
''        MsgBox "������ȣ : " & Err.Number & vbNewLine & "�������� : " & Err.Description, vbCritical, Me.Caption
''        End
''    End If
'
'
'End Sub
'
'
'
'Function GetExamCode() As Integer
'    Dim i, j As Long
'
'    ClearSpread vasTemp
'    GetExamCode = -1
'    gAllExam = ""
'         SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf
'    SQL = SQL & "  From equipexam " & vbCrLf
'    SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
'    SQL = SQL & "   And (examcode <> '' or examcode is not null) "
'    SQL = SQL & " Order by  examcode "
'    Res = GetDBSelectVas(gLocal, SQL, vasCode)
'    If Res > 0 Then
'        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 6)
'    Else
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    For i = 1 To vasCode.DataRowCnt
'        If i = 1 Then
'            gAllExam = "'" & Trim(GetText(vasCode, i, 2)) & "'"
'        Else
'            gAllExam = gAllExam & ",'" & Trim(GetText(vasCode, i, 2)) & "'"
'        End If
'
'        gArrEquip(i, 1) = i
'        For j = 1 To 5
'            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
'        Next j
'
'    Next i
'
'    GetExamCode = 1
'
'End Function
'
'Sub SetExamName()
'    Dim i, j As Long
'    Dim RS As ADODB.Recordset
'
'    ClearSpread vasTemp
'    SQL = "Select distinct examname,examcode, seqno,reflow, refhigh " & vbCrLf
'    SQL = SQL & "  From equipexam " & vbCrLf
'    SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
'    SQL = SQL & "   And (examcode <> '' or examcode is not null) "
'    SQL = SQL & " Order by seqno "
'
'    '-- Record Count ������
'    cn.CursorLocation = adUseClient
'    Set RS = cn.Execute(SQL, , 1)
'
'    If RS.RecordCount > 0 Then
'        spdTot.MaxCols = 14 + RS.RecordCount
'        spdTot_Print.MaxCols = 14 + RS.RecordCount
''        spdResult.MaxRows = RS.RecordCount
'    Else
'        SaveQuery SQL
'        Exit Sub
'    End If
'    i = 14
'    j = 0
'
'    Do Until RS.EOF
'        i = i + 1
'        j = j + 1
'        '-- ������
'        Call spdTot.SetText(i, 0, Trim(RS.Fields("examname").Value) & "")
'        spdTot.ColWidth(i) = 12
'        '-- ������(�μ�)
'        Call spdTot_Print.SetText(i, 0, Trim(RS.Fields("examname").Value) & "")
'        spdTot_Print.ColWidth(i) = 12
'        '-- �и���
''        Call spdResult.SetText(colTestNm, j, Trim(RS.Fields("examname").Value) & "")
''        Call spdResult.SetText(colTestCd, j, Trim(RS.Fields("examcode").Value) & "")
''        Call spdResult.SetText(colReference, j, Trim(RS.Fields("reflow").Value) & "-" & Trim(RS.Fields("refhigh").Value))
'        RS.MoveNext
'    Loop
'
'
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If comEqp.PortOpen = True Then
'        comEqp.PortOpen = False
'    End If
'
''    Call dce_close_env      ' Server�� ������ ���� ��
'    DisConnect_Server
'    DisConnect_Local
'    Unload Me
'    End
'End Sub
'
'Private Sub MnDBConfig_Click()
''    frmDBConfig.Show
'End Sub
'
'Private Sub MnExamConfig_Click()
'    frmTestSet.Show
'    GetExamCode
'End Sub
'
'Private Sub MnExit_Click()
'    Unload Me
'End Sub
'
'Private Sub MnMode1_Click()
'
'End Sub
'
'Private Sub MnMode_Click(Index As Integer)
'
'    Dim intCnt As Integer
'
'    '-- �и����
'    If Index = 0 Then
'        spdTot.Visible = False
''        spdOrder.Visible = True
''        spdResult.Visible = True
''        chkResult.Visible = True
'
'        MnMode(0).Checked = True
'        MnMode(1).Checked = False
'
'    '-- ���ո��
'    Else
'        spdTot.Visible = True
''        spdOrder.Visible = False
''        spdResult.Visible = False
''        chkResult.Visible = False
'
'        MnMode(0).Checked = False
'        MnMode(1).Checked = True
'    End If
'
'    Call WritePrivateProfileString("config", "IFMode", CStr(Index), App.Path & "\didim.ini")
'
'End Sub
'
'Private Sub MnSave_Click(Index As Integer)
'
'    If Index = 0 Then
''        chkMode.Caption = "�ڵ�����"
'        MnSave(0).Checked = True
'        MnSave(1).Checked = False
''        chkMode.Value = 1
'    Else
''        chkMode.Caption = "��������"
'        MnSave(0).Checked = False
'        MnSave(1).Checked = True
''        chkMode.Value = 0
'    End If
'
'    Call WritePrivateProfileString("config", "AutoSave", CStr(Index), App.Path & "\didim.ini")
'
'End Sub
'
'Private Sub MnTConfig_Click()
'    frmConfig.Show
'End Sub
'
''Private Sub MnTransAuto_Click()
''    chkMode.Caption = "Auto"
''    MnTransAuto.Checked = True
''    MnTransManual.Checked = False
''    chkMode.Value = 1
''
''End Sub
''
''Private Sub MnTransManual_Click()
''    chkMode.Caption = "Manual"
''    MnTransAuto.Checked = False
''    MnTransManual.Checked = True
''    chkMode.Value = 0
''End Sub
'
''-----------------------------------------------------------------------------'
''   ��� : �������� ����
''-----------------------------------------------------------------------------'
''Private Sub SendOrder()
''    Dim strOutput As String     '�۽��� ������
''
'''''                    Case 0      'Message Header
'''''                        MHead = "1H|\^&||||||||||P"
'''''                        brCom.Output = STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf
'''''                        SendCount = SendCount + 1
'''''                        Debug.Print "[HOST] " & STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf
'''''                        Print #1, "[HOST] " & STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf & Chr(13) + Chr(10);
'''''                        MHead = ""
'''''                    Case 1      'patient information
'''''                        Pinfo = "2P|1||" & PatientID & "|||||||||||||||||||||||||||||||"
'''''                        brCom.Output = STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf
'''''                        SendCount = SendCount + 1
'''''                        Debug.Print "[HOST] " & STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf
'''''                        Print #1, "[HOST] " & STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf & Chr(13) + Chr(10);
'''''                        Pinfo = ""
''''''                        PatientID = ""
'''''                    Case 2      'Test Order
'''''                        SendCount = SendCount + 1
'''''                        Call OrderingTheDataElecsys(brCom, com_sTemp, brSpread, brChannel, brItemdeci)
''
'''''                        Orderoutput = "3O" & "|1|" & PatientID & "|" & PatientSeq & "|" & OutPutData & "|R|" & Format(Now, "YYYYMMDDHHMMSS") & "|||||N||||||||||||||Q"
'''''                        OutPutData = STX & Orderoutput & vbCr & ETX & MakeCS(Orderoutput) & vbCr & vbLf
''
'''''                    Case 3      'Message Terminator
'''''                        SendCount = SendCount + 1
'''''                        brCom.Output = STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf
'''''                        Debug.Print "[HOST] " & STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf
'''''                        Print #1, "[HOST] " & STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf & Chr(13) + Chr(10);
'''''                    Case Else
'''''                        brCom.Output = EOT
'''''                        Debug.Print "[HOST] " & EOT
'''''                        Print #1, "[HOST] " & EOT & Chr(13) + Chr(10);
'''''                        SendCount = 0
'''''                        Flag_HQL = ""
''
''    Select Case intSndPhase
''        Case 1  '## Header
''            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
''            intSndPhase = 2
''            intFrameNo = intFrameNo + 1
''        Case 2  '## Patient
''            strOutput = intFrameNo & "P|1" & vbCr & ETX
''            intSndPhase = 4
''            'strOutput = intFrameNo & "P|1|||||||||||||||||||||||||||||||||" & vbCr & ETX
''            intFrameNo = intFrameNo + 1
''
''        Case 3  '## No Order
''
''        Case 4  '## Order
''            If mOrder.NoOrder = True Then
''                '## ���������� �������
''                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & _
''                            "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|ALL" & _
''                            "|R||||||C||||||||||||||Q" & vbCr & ETX
''                intSndPhase = 5
''
''            Else
''                If mOrder.IsSending = False Then   '## ���� ������
''                    strOutput = "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & "^" & mOrder.TubePos & _
''                                "^^SAMPLE^NORMAL|" & mOrder.Order & "|R||||||N||||||||||||||Q"
''
''                                '3O|1|9905300211|1^00014^1^^SAMPLE^NORMAL|ALL|R|20110613090006|||||X||||||||||||||O|||||
''                                '90
''                    If Len(strOutput) > 230 Then
''                        mOrder.IsSending = True
''                        mOrder.Order = Mid$(strOutput, 231)
''                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
''                        intSndPhase = 4
''                    Else
''                        strOutput = intFrameNo & strOutput & vbCr & ETX
''                        intSndPhase = 5
''                    End If
''                Else                        '## ���� ���ڿ��� ������
''                    strOutput = mOrder.Order
''                    If Len(strOutput) > 230 Then
''                        mOrder.Order = Mid$(strOutput, 231)
''                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
''                        intSndPhase = 4
''                    Else
''                        mOrder.IsSending = False
''                        strOutput = intFrameNo & strOutput & vbCr & ETX
''                        intSndPhase = 5
''                    End If
''                End If
''            End If
''            intFrameNo = intFrameNo + 1
''        Case 5  '## Termianator
''            strOutput = intFrameNo & "L|1" & vbCr & ETX
''            intSndPhase = 6
''            intFrameNo = intFrameNo + 1
''
''        Case 6  '## EOT
''            strState = ""
''            comEqp.Output = EOT
''            SetRawData "[Tx]" & EOT
''            intFrameNo = 1
''
''            Exit Sub
''    End Select
''
''    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
''    comEqp.Output = strOutput
''    Debug.Print strOutput
''    SetRawData "[Tx]" & strOutput
''End Sub
'
''-----------------------------------------------------------------------------'
''   ��� : �ش� ���ڿ��� CheckSum�� ����
''   �μ� :
''       - pMsg : ���ڿ�
''   ��ȯ : CheckSum
''-----------------------------------------------------------------------------'
'Public Function GetChkSum(ByVal pMsg As String) As String
'    Dim lngChkSum   As Long
'    Dim i           As Long
'
'    For i = 1 To Len(pMsg)
'        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
'    Next
'
'    If lngChkSum = 0 Then
'        GetChkSum = "00"
'    Else
'        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
'    End If
'End Function
'
''-- ���ݳ�¥�� �˻����� ���Ѵ�
'Function DateCompare(ByVal FDate As String) As String
'
'    DateCompare = FDate
'    If FDate <> Format(Now, "yyyymmdd") Then
'        DateCompare = Format(Now, "yyyymmdd")
'    End If
'
'End Function
'
'Private Sub comEqp_OnComm()
'    Dim EVMsg As String
'    Dim ERMsg As String
'    Dim Ret   As Long
'    Dim strDate As String
'
'    '-- ��񿡼� �Ѿ�� �ð��� �쿬�� 11:59:59�ʳ� ���Ͽ� ����� �ð��� ���
'    '-- ��� ����� �������� ������ �� �����Ƿ� ��¥�� �ǽð� ������Ʈ �Ѵ�.
'    strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
'    dtpToday.Value = Format(strDate, "####-##-##")
'
'    DoEvents
'
'    Select Case comEqp.CommEvent
'        Case comEvReceive
'            Dim Buffer      As Variant
'            Dim BufChar     As String
'            Dim lngBufLen   As Long
'            Dim i           As Long
'
'            Buffer = comEqp.Input
'
'            ' �αױ��
'            Call SetRawData(CStr(Buffer))
'
'            Call CommDefine(Buffer)
'
''            lngBufLen = Len(Buffer)
''
''            Debug.Print Buffer
''
''
''            For i = 1 To lngBufLen
''                BufChar = Mid$(Buffer, i, 1)
''                Select Case BufChar
''                    Case STX
''                        intBufCnt = 1
''                        Erase strRecvData
''                        ReDim Preserve strRecvData(intBufCnt)
''                    Case ETB
''                    Case ETX
''                        Call EditRcvData
''                    Case Else
''                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
''                End Select
''            Next i
'
'        Case comEvSend
'        Case comEvCTS
'            EVMsg$ = "CTS ���� ����"
'        Case comEvDSR
'            EVMsg$ = "DSR ���� ����"
'        Case comEvCD
'            EVMsg$ = "CD ���� ����"
'        Case comEvRing
'            EVMsg$ = "��ȭ ���� �︮�� ��"
'        Case comEvEOF
'            EVMsg$ = "EOF ����"
'
'        '���� �޽���
'        Case comBreak
'            ERMsg$ = "�ߴ� ��ȣ ����"
'        Case comCDTO
'            ERMsg$ = "�ݼ��� ���� �ð� �ʰ�"
'        Case comCTSTO
'            ERMsg$ = "CTS �ð� �ʰ�"
'        Case comDCB
'            ERMsg$ = "DCB �˻� ����"
'        Case comDSRTO
'            ERMsg$ = "DSR �ð� �ʰ�"
'        Case comFrame
'            ERMsg$ = "�����̹� ����"
'        Case comOverrun
'            ERMsg$ = "�и�Ƽ ����"
'        Case comRxOver
'            ERMsg$ = "���� ���� �ʰ�"
'        Case comRxParity
'            ERMsg$ = "�и�Ƽ ����"
'        Case comTxFull
'            ERMsg$ = "���� ���ۿ� ������ ����"
'        Case Else
'            ERMsg$ = "�� �� ���� ���� �Ǵ� �̺�Ʈ"
'    End Select
'
'
'End Sub
'
''-----------------------------------------------------------------------------'
''   ��� : �ش� ���ڵ��ȣ�� ���� �������� ��ȸ, tblReady, tblResult�� ǥ��
''   �μ� :
''       - pBarNo : ���ڵ��ȣ
''-----------------------------------------------------------------------------'
''Private Sub GetOrder(ByVal pBarNo As String)
''    Dim i           As Integer
''    Dim intRow      As Long
''    Dim strItems    As String
''
''    intRow = -1
''    For i = 1 To spdorder.DataRowCnt
''        If Trim(GetText(spdorder, i, colBarcode)) = pBarNo Then
''            intRow = i
''            Exit For
''        End If
''    Next i
''
''    If intRow < 0 Then
''        intRow = spdorder.DataRowCnt + 1
''        If spdorder.MaxRows < intRow Then
''            spdorder.MaxRows = intRow
''        End If
''    End If
''
''    Call SetText(spdorder, pBarNo, intRow, colBarcode)         '2
''    Call SetText(spdorder, mOrder.RackNo, intRow, colRack)     '3
''    Call SetText(spdorder, mOrder.TubePos, intRow, colPos)     '4
''    Call vasActiveCell(spdorder, intRow, colBarcode)
''    Call ClearSpread(vasRes)
''
''    Call GetSampleInfoW(intRow)                            '5,6,7,8
''
''    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)
''
''    '-- ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.
''    '-- intRow �߰�
''    strItems = GetGetEquipExamCode_AU480(gEquip, pBarNo, intRow)
''
''    If Trim(strItems) = "" Then
''        mOrder.NoOrder = True
''        mOrder.Order = ""
''        'S 003401 0019          1013001918    E
''        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
''        Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
''    Else
''        mOrder.NoOrder = False
''        mOrder.Order = strItems
''        'S 003401 0019          1013001918    E      01020304050607091011121415161719212632
''        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
''        'comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E012" & ETX
''
''
''        Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
''
''    End If
'
'
''End Sub
'
''-----------------------------------------------------------------------------'
''   ��� :
''   �μ� :
''       - pBarNo : ���ڵ��ȣ
''-----------------------------------------------------------------------------'
'Private Sub SetPatInfo(ByVal pBarNo As String)
'    Dim i           As Integer
'    Dim intRow      As Long
'    Dim strItems    As String
'
'    intRow = -1
'
'    If gMode = "0" Then
'        For i = 1 To spdOrder.DataRowCnt
'            If Trim(GetText(spdOrder, i, colBarcode)) = pBarNo Then
'                intRow = i
'                Exit For
'            End If
'        Next i
'
'        If intRow < 0 Then
'            intRow = spdOrder.DataRowCnt + 1
'            If spdOrder.MaxRows < intRow Then
'                spdOrder.MaxRows = intRow
'            End If
'        End If
'
'        Call vasActiveCell(spdOrder, intRow, colBarcode)
'
'    Else
'        For i = 1 To spdTot.DataRowCnt
'            If Trim(GetText(spdTot, i, colBarcode)) = pBarNo Then
'                intRow = i
'                Exit For
'            End If
'        Next i
'
'        If intRow < 0 Then
'            intRow = spdTot.DataRowCnt + 1
'            If spdTot.MaxRows < intRow Then
'                spdTot.MaxRows = intRow
'            End If
'        End If
'
'        Call vasActiveCell(spdTot, intRow, colBarcode)
'
'    End If
'
'
'    Call ClearSpread(spdResult)
'
'    '-- ȯ������ ��������
'    'Call GetSampleInfoW(intRow)                                '5,6,7,8
'
'    gRow = intRow
'
'    '-- 52P�� �¿�� �ȵȴ�.
'    'gOrderExam = GetOrderExamCode(gEquip, pBarNo)
'    gOrderExam = "'36721','36722','36723','36724'"
'
'End Sub
'
'
''-----------------------------------------------------------------------------'
''   ��� : ���κ� ������ ������ ����
''-----------------------------------------------------------------------------'
'Private Sub EditRcvData()
''    Dim strRcvBuf    As String   '������ Data
''    Dim strType      As String   '������ Record Type
''    Dim strBarno     As String   '������ ���ڵ��ȣ
''    Dim strSeq       As String   '������ Sequence
''    Dim strRackNo    As String   '������ Rack Or Disk No
''    Dim strTubePos   As String   '������ Tube Position
''    Dim strIntBase   As String   '������ ������ �˻��
''    Dim strResult    As String   '������ ���
''    Dim strQCResult  As String   '������ ���(QC)
''    Dim strFlag      As String   '������ Abnormal Flag
''    Dim strComm      As String   '������ Comment
''    Dim strTemp1     As String
''    Dim strTemp2     As String
''    Dim intCnt       As Integer
''
''    Dim lsExamCode As String
''    Dim lsExamName As String
''    Dim lsSeqNo As String
''    Dim lsResult_Buff As String
''    Dim lsExamDate As String
''    Dim lsEquipRes As String
''    Dim lsResRow    As String
''    Dim ii As Integer
'''    Dim blnPSA       As Boolean
'''    Dim blnfPSA      As Boolean
'''    Dim strPSA       As String
'''    Dim strfPSA      As String
''    Dim strTmp      As String
''    Dim intIdx      As Integer
''
'''    blnPSA = False
'''    blnfPSA = False
''
''    For intCnt = 1 To UBound(strRecvData)
''        strRcvBuf = strRecvData(intCnt)
''        strType = Mid$(strRcvBuf, 1, 2)
''
''        Select Case strType
''            Case "R "    '## Inquiry Order
''                strBarno = Trim(Mid(strRcvBuf, 14, 20))
''                strRackNo = Mid(strRcvBuf, 3, 4)
''                strTubePos = Mid(strRcvBuf, 7, 2)
''
''                mOrder.BarNo = strBarno
''                mOrder.RackNo = strRackNo
''                mOrder.TubePos = strTubePos
''                mOrder.Seq = Mid(strRcvBuf, 9, 5)
''                'R 003201 0018          1013001917
''                'S 003201 0018          1013001917    E      13
''
''                Call GetOrder(strBarno)
''
''                '===========================================================================
''
''            Case "D "    '## Result
''                strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
''                mResult.BarNo = strBarno
''                mResult.RackNo = Mid(strRcvBuf, 3, 4)
''                mResult.TubePos = Mid(strRcvBuf, 7, 2)
''
''                If strBarno = "" Then Exit Sub
''
''                'intIdx = InStr(strRcvBuf, "E")
''                'intIdx = InStr(strRcvBuf, "E")
''                strTmp = Mid$(strRcvBuf, 29)
''
''                Call SetPatInfo(strBarno)
''
''                Do While Len(strTmp) >= 11
''
''                    strIntBase = Mid$(strTmp, 2, 2)
''                    strResult = Mid$(strTmp, 4, 6)
''                    strComm = Mid$(strTmp, 10, 1)
''
''                    If strResult <> "" Then
''                        SQL = ""
''                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
''                        SQL = SQL & "  FROM EQUIPEXAM"
''                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
''                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
''                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
''
''                        Res = GetDBSelectColumn(gLocal, SQL)
''
''                        '-- ���� ���� ���
''                        If Res > 0 Then
''                            lsExamCode = Trim(gReadBuf(0))
''                            lsExamName = Trim(gReadBuf(1))
''                            lsSeqNo = Trim(gReadBuf(2))
''
''                            lsResRow = vasRes.DataRowCnt + 1
''                            If vasRes.MaxRows < lsResRow Then
''                                vasRes.MaxRows = lsResRow
''                            End If
''
''                            '�Ҽ��� ó��, ��� ���� ó��
''                            lsEquipRes = strResult
''                            strResult = SetResult(strResult, strIntBase)
''                            lsResult_Buff = strResult
''
''                            '-- Work List
''                            SetText spdorder, strResult, gRow, colA1c                  '���
''                            SetText spdorder, strComm, gRow, colA1c + 1                'Flag
''                            SetText spdorder, "Result", gRow, colState                 '�������
''                            '-- ��� List
''                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '����ڵ�
''                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '�˻��ڵ�
''                            SetText vasRes, lsExamName, lsResRow, colExamName       '�˻��
''                            SetText vasRes, strResult, lsResRow, colResult          '���
''                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
''                            SetText vasRes, strComm, lsResRow, 7                    'Flag
''                            '-- ���� ����
''                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
''
''                            lsResult_Buff = ""
''
''                        '-- ���� ���� ���
''                        Else
''
''                                  SQL = "Select examcode, examname, seqno "
''                            SQL = SQL & "  From equipexam"
''                            SQL = SQL & " Where equipno = '" & gEquip & "' "
''                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
''                            Res = GetDBSelectColumn(gLocal, SQL)
''
''                            If Res > 0 Then
''                                lsExamCode = Trim(gReadBuf(0))
''                                lsExamName = Trim(gReadBuf(1))
''                                lsSeqNo = Trim(gReadBuf(2))
''
''                                lsResRow = vasRes.DataRowCnt + 1
''                                If vasRes.MaxRows < lsResRow Then
''                                    vasRes.MaxRows = lsResRow
''                                End If
''
''                                '�Ҽ��� ó��, ��� ���� ó��
''                                lsEquipRes = strResult
''                                strResult = SetResult(strResult, strIntBase)
''                                lsResult_Buff = strResult
''
''                                '-- Work List
''                                SetText spdorder, strResult, gRow, colA1c                  '���
''                                SetText spdorder, strComm, gRow, colA1c + 1                'Flag
''                                SetText spdorder, "Result", gRow, colState                 '�������
''                                '-- ��� List
''                                SetText vasRes, strIntBase, lsResRow, colEquipCode      '����ڵ�
''                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '�˻��ڵ�
''                                SetText vasRes, lsExamName, lsResRow, colExamName       '�˻��
''                                SetText vasRes, strResult, lsResRow, colResult          '���
''                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
''                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
''                                '-- ���� ����
''                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
''
''                                lsResult_Buff = ""
''                                strState = ""
''                            End If
''                        End If
''                    End If
''                    strTmp = Mid$(strTmp, 12)
''                Loop
''                strState = "R"
''
''                If MnTransAuto.Checked = True Then
''
''                    Res = SaveTransDataW(gRow)
''
''                    If Res = -1 Then
''                        '-- ���� ����
''                        SetForeColor spdorder, gRow, gRow, 1, colState, 255, 0, 0
''                        SetText spdorder, "Failed", gRow, colState
''                    Else
''                        '-- ���� ����
''                        SetBackColor spdorder, gRow, gRow, 1, colState, 202, 255, 112
''                        SetText spdorder, "Trans", gRow, colState
''
''                        SQL = " Update pat_res Set " & vbCrLf & _
''                              " sendflag = '2' " & vbCrLf & _
''                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
''                              " And barcode = '" & Trim(GetText(spdorder, gRow, colBarcode)) & "' "
''                        Res = SendQuery(gLocal, SQL)
''                        If Res = -1 Then
''                            SaveQuery SQL
''                            Exit Sub
''                        End If
''                    End If
''                End If
''
''                SetText spdorder, "Result", gRow, colState
''                strState = ""
''
''        End Select
''    Next
'
'End Sub
'
'
''Function SetResult(asResult As String, asEquipCode As String)
''    Dim i As Integer
''    Dim sLVal As String
''    Dim sHVal As String
''    Dim sEquipCode As String
''    Dim sEquipRes As String
''    Dim sResult As String
''    Dim sPoint As Integer
''    Dim sResType As String
''    Dim sResFlag As String
''
''
''    sEquipRes = Trim(asResult)
''    sEquipCode = Trim(asEquipCode)
''    sResFlag = ""
''
''    If sEquipCode = "" Then
''        Exit Function
''    End If
''
'''    If IsNumeric(sEquipRes) = False Then
'''        Exit Function
'''    End If
''
''    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
''    Res = GetDBSelectColumn(gLocal, SQL)
''
''    If IsNumeric(gReadBuf(0)) = True Then
''        sPoint = CInt(gReadBuf(0))
''        sResType = ""
''        For i = 0 To sPoint
''            If i = 0 Then
''                sResType = "#0"
''            ElseIf i = 1 Then
''                sResType = sResType & ".0"
''            Else
''                sResType = sResType & "0"
''            End If
''        Next
''
''        sResult = Format(sEquipRes, sResType)
''    Else
''        sResult = sEquipRes
''    End If
''
''''    If IsNumeric(gReadBuf(1)) = True Then
''''        sLVal = gReadBuf(1)
''''        If CCur(sLVal) > CCur(sEquipRes) Then
''''            sResFlag = "H"
''''        End If
''''    End If
''''
''''    If IsNumeric(gReadBuf(2)) = True Then
''''        sHVal = gReadBuf(2)
''''        If CCur(sHVal) < CCur(sEquipRes) Then
''''            sResFlag = ">"
''''        End If
''''    End If
''
''    If IsNumeric(gReadBuf(1)) = True And IsNumeric(gReadBuf(2)) = True Then
''        sLVal = gReadBuf(1)
''        sHVal = gReadBuf(2)
''        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
''            sResFlag = ""
''        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
''            sResFlag = "H"
''        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
''            sResFlag = "L"
''        End If
''    End If
''
''    gsFlag = sResFlag
''    SetResult = sResult
''
''End Function
'
'' asRow1 = Work List
'' asRow2 = ��� List
'Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
'    Dim sCnt As String
'    Dim sExamDate As String
'
'    sExamDate = Format(dtpToday, "yyyymmdd")
'
'    SQL = ""
'    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(spdOrder, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  AND EQUIPCODE = '" & Trim(GetText(spdResult, asRow2, colChannel)) & "'" & vbCrLf & _
'          "  AND EXAMCODE = '" & Trim(GetText(spdResult, asRow2, colTestCd)) & "'"
'
'    Res = SendQuery(gLocal, SQL)
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    SQL = ""
'    SQL = SQL & "INSERT INTO PAT_RES("
'    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
'                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
'    SQL = SQL & "VALUES("
'    SQL = SQL & "'" & Trim(GetText(spdOrder, asRow1, colOrdDate)) & "', "
'    SQL = SQL & "'" & gEquip & "', "
'    SQL = SQL & "'" & Trim(GetText(spdOrder, asRow1, colBarcode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdOrder, asRow1, colRack)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdOrder, asRow1, colPos)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(spdOrder, asRow1, colPID)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdOrder, asRow1, colPName)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdOrder, asRow1, colWell)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdOrder, asRow1, colWell)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdResult, asRow2, colChannel)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdResult, asRow2, colTestCd)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdResult, asRow2, colOther)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(spdResult, asRow2, colEqpResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdResult, asRow2, colLisResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdResult, asRow2, colTestNm)) & "', "
'    SQL = SQL & "'0', "
'    SQL = SQL & "'" & gIFUser & "')"
'
'
'    Res = SendQuery(gLocal, SQL)
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'End Function
'
'Function SetLocalDBTot(ByVal asRow As Long, ByVal pIntBase As String, ByVal pExamCode As String, ByVal pExamName As String, ByVal pEqpResult As String, ByVal pLisResult, ByVal pSeq)
'    Dim sCnt As String
'    Dim sExamDate As String
'
'    sExamDate = Format(dtpToday, "yyyymmdd")
'
'
'    SQL = ""
'    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Trim(GetText(spdTot, asRow, colOrdDate)) & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(spdTot, asRow, colBarcode)) & "' " & vbCrLf & _
'          "  AND EQUIPCODE = '" & pIntBase & "'" & vbCrLf & _
'          "  AND EXAMCODE = '" & pExamCode & "'"
'
'    Res = SendQuery(gLocal, SQL)
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    SQL = ""
'    SQL = SQL & "INSERT INTO PAT_RES("
'    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
'                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
'    SQL = SQL & "VALUES("
'    SQL = SQL & "'" & Trim(GetText(spdTot, asRow, colOrdDate)) & "', "
'    SQL = SQL & "'" & gEquip & "', "
'    SQL = SQL & "'" & Trim(GetText(spdTot, asRow, colBarcode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdTot, asRow, colRack)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdTot, asRow, colPos)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(spdTot, asRow, colPID)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdTot, asRow, colPName)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdTot, asRow, colWell)) & "', "
'    SQL = SQL & "'" & Trim(GetText(spdTot, asRow, colWell)) & "', "
'    SQL = SQL & "'" & pIntBase & "', "
'    SQL = SQL & "'" & pExamCode & "', "
'    SQL = SQL & "'" & pSeq & "', " & vbCrLf
'    SQL = SQL & "'" & pEqpResult & "', "
'    SQL = SQL & "'" & pLisResult & "', "
'    SQL = SQL & "'" & pExamName & "', "
'    SQL = SQL & "'0', "
'    SQL = SQL & "'" & gIFUser & "')"
'
'
'    Res = SendQuery(gLocal, SQL)
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'End Function
'
'
'
''Sub Var_Clear()
''    gsBarCode = ""
''    gsPID = ""
''    gsRackNo = ""
''    gsPosNo = ""
''    gsResDateTime = ""
''    gsSeqNo = ""
''    gsExamCode = ""
''    gsExamName = ""
''    gsOrder = ""
''    gsResult = ""
''End Sub
'
'
'
'Private Sub picLogin_Click()
'
'    Dim sMsg As String
'    sMsg = "�˻��ڸ� �Է����ּ���."
'    lblUser.Caption = InputBox(sMsg, "�˻��� �Է�")
'
'End Sub
'
'
'
''Private Sub spdIntList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
''    If BlockRow = -1 And BlockRow2 = -1 Then
''        giBSRow = 1
''        giBERow = spdIntList.MaxRows
''    Else
''        giBSRow = CInt(BlockRow)
''        giBERow = CInt(BlockRow2)
''    End If
''End Sub
''
''Private Sub spdorder_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
''    Dim i As Integer
''
''    For i = BlockRow To BlockRow2
''        spdorder.Col = 1
''        spdorder.Row = i
''        If spdorder.Value = 0 Then
''        spdorder.Value = 1
''        Else
''        spdorder.Value = 0
''        End If
''    Next i
''End Sub
''
''
''Private Sub spdorder_Click(ByVal Col As Long, ByVal Row As Long)
''    Dim lsID As String
''
''    If Row < 1 Or Row > spdorder.DataRowCnt Then
''        Exit Sub
''    End If
''
''    lsID = Trim(GetText(spdorder, Row, colBarcode))
''    lblChangeBar.Caption = lsID
''    lblChangePID.Caption = Trim(GetText(spdorder, Row, colPID))
''    lblBarcode(0).Caption = lsID
''    lblPname(0).Caption = Trim(GetText(spdorder, Row, colPName))
''    'Local���� �ҷ�����
''    ClearSpread vasRes
''
''    '����ڵ�, �˻��ڵ�, �˻��, ���, ����
''    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
''          "FROM PAT_RES " & vbCrLf & _
''          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
''          "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf & _
''          " AND DISKNO = '" & Trim(GetText(spdorder, Row, colRack)) & "' " & vbCrLf & _
''          " AND POSNO = '" & Trim(GetText(spdorder, Row, colPos)) & "' " & vbCrLf & _
''          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "
''
''    Res = GetDBSelectVas(gLocal, SQL, vasRes)
''    If Res = -1 Then
''        SaveQuery SQL
''        Exit Sub
''    End If
''
''    vasRes.MaxRows = vasRes.DataRowCnt
''End Sub
'
''Private Sub spdorder_KeyDown(KeyCode As Integer, Shift As Integer)
''    Dim iRow As Long
''    Dim lsID As String
''    Dim lsTime As String
''    Dim lsPid As String
''    Dim i As Integer
''
''    iRow = spdorder.ActiveRow
''    If KeyCode = vbKeyDelete Then
''        If iRow < 1 Or iRow > spdorder.DataRowCnt Then
''            Exit Sub
''        End If
''
''        lsID = Trim(GetText(spdorder, iRow, colBarcode))
''        lsPid = Trim(GetText(spdorder, iRow, colPID))
''
''        If MsgBox("�ش� ȯ�ڰ���� �����Ͻðڽ��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbNo Then
''            Exit Sub
''        End If
''
''        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
''              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
''              " AND PID = '" & lsPid & "' " & vbCrLf & _
''              " AND DISKNO = '" & Trim(GetText(spdorder, iRow, colRack)) & "' " & vbCrLf & _
''              " AND POSNO = '" & Trim(GetText(spdorder, iRow, colPos)) & "' " & vbCrLf & _
''              " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
''        res = SendQuery(gLocal, SQL)
''
''        If res = -1 Then
''            SaveQuery SQL
''            Exit Sub
''        End If
''
''        DeleteRow spdorder, iRow, iRow
''        vasRes.MaxRows = 0
''    ElseIf KeyCode = 13 Then
''
''        GetSampleInfoW (iRow)
''
''        lsID = Trim(GetText(spdorder, iRow, colBarcode))
''
''        'Local���� �ҷ�����
''        ClearSpread vasTemp
''
''        '����ڵ�, �˻��ڵ�, �˻��, ���, ����
''        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, SEQNO " & vbCrLf & _
''              "  FROM EQUIPEXAM " & vbCrLf & _
''              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''              " ORDER BY SEQNO "
''
''        res = GetDBSelectVas(gLocal, SQL, vasTemp)
''        If res = -1 Then
''            SaveQuery SQL
''            Exit Sub
''        End If
''        If lsID <> lblChangeBar.Caption Then
''            For i = 1 To 3
''                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
''                  "POSNO, PID, PNAME, " & vbCrLf & _
''                  "JUMIN, PSEX, PAGE, " & vbCrLf & _
''                  "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
''                  "SEQNO, RESULT, EXAMNAME, " & vbCrLf & _
''                  "SENDFLAG, Hospital, refflag) " & vbCrLf & _
''                  "VALUES('" & gEquip & "', '" & Trim(GetText(spdorder, iRow, colBarcode)) & "', '" & Trim(GetText(spdorder, iRow, colRack)) & "', " & vbCrLf & _
''                  "'" & Trim(GetText(spdorder, iRow, colPos)) & "', '" & Trim(GetText(spdorder, iRow, colPID)) & "', '" & Trim(GetText(spdorder, iRow, colPName)) & "', " & vbCrLf & _
''                  "'" & Trim(GetText(spdorder, iRow, colJumin)) & "', '" & Trim(GetText(spdorder, iRow, colSex)) & "', " & 0 & ", " & vbCrLf & _
''                  "'" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "', '" & Trim(GetText(spdorder, 0, colState + (i * 2) - 1)) & "', '" & Trim(GetText(vasTemp, i, 2)) & "', " & vbCrLf & _
''                  "'" & Trim(GetText(vasTemp, i, 4)) & "', '" & Trim(GetText(spdorder, iRow, colState + (i * 2) - 1)) & "', '" & Trim(GetText(vasTemp, i, 3)) & "', " & vbCrLf & _
''                  "'1', '" & Trim(GetText(spdorder, iRow, colHospital)) & "', '" & Trim(GetText(spdorder, iRow, colState + (i * 2))) & "')"
''                res = SendQuery(gLocal, SQL)
''            Next i
''
''            SQL = " DELETE FROM PAT_RES " & vbCrLf & _
''                  " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''                  " AND BARCODE = '" & lblChangeBar.Caption & "' " & vbCrLf & _
''                  " AND PID = '" & lblChangePID.Caption & "' " & vbCrLf & _
''                  " AND DISKNO = '" & Trim(GetText(spdorder, iRow, colRack)) & "' " & vbCrLf & _
''                  " AND POSNO = '" & Trim(GetText(spdorder, iRow, colPos)) & "' " & vbCrLf & _
''                  " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
''            res = SendQuery(gLocal, SQL)
''
''        ElseIf lsID = lblChangeBar.Caption Then
''            For i = 1 To 3
''                SQL = "UPDATE PAT_RES "
''                SQL = SQL & vbCrLf & "   SET RESULT ='" & Trim(GetText(spdorder, iRow, colState + (i * 2) - 1)) & "' "
''                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(spdorder, iRow, colBarcode)) & "' "
''                SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
''                SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasTemp, i, 2)) & "' "
''                SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(spdorder, 0, colState + (i * 2) - 1)) & "' "
''                SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(spdorder, iRow, colPID)) & "' "
''                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(spdorder, iRow, colRack)) & "' "
''                SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(spdorder, iRow, colPos)) & "' "
''                SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
''                res = SendQuery(gLocal, SQL)
''            Next i
''        End If
''        SetText spdorder, "Result", gRow, colState
''
''    End If
''
''
''End Sub
'
''Private Sub spdorder_KeyUp(KeyCode As Integer, Shift As Integer)
''    Dim lRow As Long
''
''    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
''        lRow = spdorder.ActiveRow
''        If lRow < 1 Or lRow > spdorder.DataRowCnt Then Exit Sub
''
''        spdorder_Click colBarcode, lRow
''    End If
''End Sub
'
''Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
''    Dim sResDateTime As String
''    Dim sControl As String
''    Dim sLotNo As String
''
''    Dim sRefLow As String
''    Dim sRefHigh As String
''    Dim sRefFlag As String
''
''    Dim sCnt As String
''
''    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
''    'sControl = Trim(Left(asBarcode, 2))
''    'sLotNo = Trim(Mid(asBarcode, 3))
''    sControl = asBarcode
''    sRefFlag = ""
''
''    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
''          "where equipno = '" & gEquip & "' " & vbCrLf & _
''          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
''          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
''          "  and levelname = '" & sControl & "' " & vbCrLf & _
''          "  and equipcode = '" & asExamCode & "' "
''    res = GetDBSelectColumn(gLocal, SQL)
''    If res > 0 Then
''        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
''            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
''            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
''            If CCur(sRefHigh) < CCur(asRes2) Then
''                sRefFlag = "H"
''            End If
''            If CCur(sRefLow) > CCur(asRes2) Then
''                sRefFlag = "L"
''            End If
''        End If
''    End If
''
''    sCnt = ""
''    SQL = "Select count(*) from qc_res " & vbCrLf & _
''          "where equipno = '" & gEquip & "' " & vbCrLf & _
''          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
''          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
''          "  and levelname = '" & sControl & "' " & vbCrLf & _
''          "  and equipcode = '" & asExamCode & "' "
''    res = db_select_Var(gLocal, SQL, sCnt)
''    If res <= 0 Then
''        SaveQuery SQL
''        db_RollBack gLocal
''        Exit Function
''    End If
''    res = db_select_Var(gLocal, SQL, sCnt)
''    If res <= 0 Then
''        SaveQuery SQL
''        Exit Function
''    End If
''    If Not IsNumeric(sCnt) Then sCnt = "0"
''
''    If CInt(sCnt) > 0 Then
''        SQL = "delete from qc_res " & vbCrLf & _
''              "where equipno = '" & gEquip & "' " & vbCrLf & _
''              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
''              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
''              "  and levelname = '" & sControl & "' " & vbCrLf & _
''              "  and equipcode = '" & asExamCode & "' "
''        res = SendQuery(gLocal, SQL)
''        If res = -1 Then
''            'db_RollBack gLocal
''            SaveQuery SQL
''            Exit Function
''        End If
''    End If
''    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
''          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
''    res = SendQuery(gLocal, SQL)
''    If res = -1 Then
''        'db_RollBack gLocal
''        SaveQuery SQL
''        Exit Function
''    End If
''
''End Function
'
''Private Sub vasRID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
''    Dim i As Integer
''
''    For i = BlockRow To BlockRow2
''        vasRID.Col = 1
''        vasRID.Row = i
''        If vasRID.Value = 0 Then
''        vasRID.Value = 1
''        Else
''        vasRID.Value = 0
''        End If
''    Next i
''End Sub
''
''Private Sub vasRID_Click(ByVal Col As Long, ByVal Row As Long)
''    Dim lsID As String
''    Dim i As Integer
''
''    If Row < 1 Or Row > vasRID.DataRowCnt Then
''        Exit Sub
''    End If
''
''    lsID = Trim(GetText(vasRID, Row, colBarcode))
''    lblChangeBar.Caption = lsID
''    lblBarcode(1).Caption = lsID
''    lblChangePID.Caption = Trim(GetText(vasRID, Row, colPID))
''    lblPname(1).Caption = Trim(GetText(vasRID, Row, colPName))
''    lblRrow.Caption = Row
''    'Local���� �ҷ�����
''    ClearSpread vasRRes
''
''    '����ڵ�, �˻��ڵ�, �˻��, ���, ����
''    SQL = ""
''    SQL = "SELECT EQUIPCODE,EXAMCODE,EXAMNAME,EQUIPRESULT,RESULT,SEQNO,REFFLAG " & vbCrLf & _
''          "  FROM PAT_RES " & vbCrLf & _
''          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''          "   AND BARCODE = '" & lsID & "' " & vbCrLf & _
''          "   AND DISKNO = '" & Trim(GetText(vasRID, Row, colRack)) & "' " & vbCrLf & _
''          "   AND POSNO = '" & Trim(GetText(vasRID, Row, colPos)) & "' " & vbCrLf & _
''          "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' " & vbCrLf & _
''          " GROUP BY EQUIPCODE,EXAMCODE,EXAMNAME,EQUIPRESULT,RESULT,SEQNO,REFFLAG "
''
''    Res = GetDBSelectVas(gLocal, SQL, vasRRes)
''
''    If Res = -1 Then
''        SaveQuery SQL
''        Exit Sub
''    End If
''
''    vasRRes.MaxRows = vasRRes.DataRowCnt
''
''    For i = 1 To vasRRes.MaxRows
''        If Trim(GetText(vasRRes, i, colFLAG)) = "H" Then
''            SetForeColor vasRRes, i, i, colResult, colResult, 255, 0, 0
''        ElseIf Trim(GetText(vasRRes, i, colFLAG)) = "L" Then
''            SetForeColor vasRRes, i, i, colResult, colResult, 0, 255, 0
''        End If
''    Next i
''End Sub
'
''Private Sub vasRID_KeyDown(KeyCode As Integer, Shift As Integer)
''    Dim iRow As Long
''    Dim lsID As String
''    Dim lsTime As String
''    Dim lsPid As String
''    Dim i As Integer
''
''    iRow = vasRID.ActiveRow
''
''    If KeyCode = 13 Then
''
''        If GetSampleInfoR(iRow) = -1 Then
''            Exit Sub
''        End If
''
''        lsID = Trim(GetText(vasRID, iRow, colBarcode))
''
''        'Local���� �ҷ�����
''        ClearSpread vasTemp
''
''        '����ڵ�, �˻��ڵ�, �˻��, ���, ����
''        SQL = ""
''        SQL = SQL & "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf
''        SQL = SQL & "  FROM PAT_RES " & vbCrLf
''        SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' "
''        SQL = SQL & "   AND BARCODE  = '" & lsID & "' " & vbCrLf
''        SQL = SQL & "   AND EXAMDATE = '" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "' " & vbCrLf
''        SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "
''
''        Res = GetDBSelectVas(gLocal, SQL, vasTemp)
''        If Res = -1 Then
''            SaveQuery SQL
''            Exit Sub
''        End If
''
''        If lsID <> lblChangeBar.Caption Then
''            For i = 1 To vasRRes.DataRowCnt
''                SQL = ""
''                SQL = SQL & "INSERT INTO PAT_RES("
''                SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
''                            "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
''                            "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
''                SQL = SQL & "VALUES("
''                SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
''                SQL = SQL & "'" & gEquip & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colBarcode)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colDISK)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPos)) & "', " & vbCrLf
''                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPID)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPName)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colSex)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colAge)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colEquipCode)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colExamCode)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colSeq)) & "', " & vbCrLf
''                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colMachResult)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colResult)) & "', "
''                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colExamName)) & "', "
''                SQL = SQL & "'0', "
''                SQL = SQL & "'" & gIFUser & "')"
''
''                Res = SendQuery(gLocal, SQL)
''
''                If Res = -1 Then
''                    SaveQuery SQL
''                    Exit Sub
''                End If
''            Next i
''
''            SQL = ""
''            SQL = SQL & "DELETE FROM PAT_RES " & vbCrLf
''            SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' " & vbCrLf
''            SQL = SQL & "   AND BARCODE  = '" & lblChangeBar.Caption & "' " & vbCrLf
''            SQL = SQL & "   AND PID      = '" & lblChangePID.Caption & "' " & vbCrLf
''            SQL = SQL & "   AND DISKNO   = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
''            SQL = SQL & "   AND POSNO    = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
''            SQL = SQL & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
''
''            Res = SendQuery(gLocal, SQL)
''
''            If Res = -1 Then
''                SaveQuery SQL
''                Exit Sub
''            End If
''
''        ElseIf lsID = lblChangeBar.Caption Then
''            For i = 1 To vasRRes.DataRowCnt
''
''                SQL = ""
''                SQL = SQL & "UPDATE PAT_RES " & vbCrLf
''                SQL = SQL & "   SET RESULT    ='" & Trim(GetText(vasRRes, i, colResult)) & "' " & vbCrLf
''                SQL = SQL & " WHERE BARCODE   = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf
''                SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
''                SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRRes, i, colExamCode)) & "' " & vbCrLf
''                SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRRes, i, colEquipCode)) & "' " & vbCrLf
''                SQL = SQL & "   AND PID       = '" & Trim(GetText(vasRID, iRow, colPID)) & "' " & vbCrLf
''                SQL = SQL & "   AND DISKNO    = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
''                SQL = SQL & "   AND POSNO     = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
''                SQL = SQL & "   AND EXAMDATE  = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
''
''                Res = SendQuery(gLocal, SQL)
''
''                If Res = -1 Then
''                    SaveQuery SQL
''                    Exit Sub
''                End If
''            Next i
''        End If
''    ElseIf KeyCode = vbKeyDelete Then
''        If iRow < 1 Or iRow > vasRID.DataRowCnt Then
''            Exit Sub
''        End If
''
''        lsID = Trim(GetText(vasRID, iRow, colBarcode))
''        lsPid = Trim(GetText(vasRID, iRow, colPID))
''
''        If MsgBox("�ش� ȯ�ڰ���� �����Ͻðڽ��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbNo Then
''            Exit Sub
''        End If
''
''        SQL = ""
''        SQL = SQL & "DELETE FROM PAT_RES " & vbCrLf
''        SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' " & vbCrLf
''        SQL = SQL & "   AND BARCODE  = '" & lsID & "' " & vbCrLf
''        SQL = SQL & "   AND PID      = '" & lsPid & "' " & vbCrLf
''        SQL = SQL & "   AND DISKNO   = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
''        SQL = SQL & "   AND POSNO    = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
''        SQL = SQL & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
''
''        Res = SendQuery(gLocal, SQL)
''
''        If Res = -1 Then
''            SaveQuery SQL
''            Exit Sub
''        End If
''
''        DeleteRow vasRID, iRow, iRow
''        vasRRes.MaxRows = 0
''
''    End If
''End Sub
'
''Private Sub vasRID_KeyUp(KeyCode As Integer, Shift As Integer)
''    Dim lRow As Long
''
''    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
''        lRow = vasRID.ActiveRow
''        If lRow < 1 Or lRow > vasRID.DataRowCnt Then Exit Sub
''
''        vasRID_Click colBarcode, lRow
''    End If
''End Sub
'
''Private Sub vasRRes_KeyDown(KeyCode As Integer, Shift As Integer)
''
''    If KeyCode = 13 Then: vasRID_KeyDown KeyCode, 0
''End Sub
'
'Private Sub spdTot_KeyPress(KeyAscii As Integer)
'    Dim varTmp As Variant
'
'    If KeyAscii = 13 And spdTot.ActiveCol > colState Then
'        Call spdTot.GetText(spdTot.ActiveCol, spdTot.ActiveRow, varTmp)
'        If UCase(CStr(varTmp)) = "N" Then
'            Call spdTot.SetText(spdTot.ActiveCol, spdTot.ActiveRow, "Negative")
'            spdTot.Row = spdTot.ActiveRow
'            spdTot.Col = spdTot.ActiveCol
'            spdTot.BackColor = vbGreen
'        ElseIf UCase(CStr(varTmp)) = "P" Then
'            Call spdTot.SetText(spdTot.ActiveCol, spdTot.ActiveRow, "Positive")
'            spdTot.Row = spdTot.ActiveRow
'            spdTot.Col = spdTot.ActiveCol
'            spdTot.BackColor = vbGreen
'        End If
'    End If
'End Sub
'
'
Private Sub Timer1_Timer()
       
    XProgress1.Value = XProgress1.Value + 1
    If XProgress1.Value = 100 Then
        XProgress1.Value = 1
        XProgress1.Visible = False
    End If
    
End Sub

Private Sub XButton2_Click()
    
    Timer1.Interval = 50
    Timer1.Enabled = True
    XProgress1.Value = 1
    
    XProgress1.Visible = True

End Sub

