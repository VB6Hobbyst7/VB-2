VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15735
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15735
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '최대화
   Begin TabDlg.SSTab tabWork 
      Height          =   8490
      Left            =   30
      TabIndex        =   7
      Top             =   510
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   14975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   16711680
      TabCaption(0)   =   " ▒    WorkList     "
      TabPicture(0)   =   "frmComm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAppend(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdMake"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSearch"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtpStopDt"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtpStartDt"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdWorkList"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "pnlCom"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdPrint"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdStartNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "spdWorklist"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtBarCode"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "pnlCom2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdRequist(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtResult"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdRackNo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdWordQuery"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdEot"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "SSPanel2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame3"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdOrder"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdPosNo"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdNext"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdPrevious"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtDump"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "SSPanel1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdSel(2)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdSel(3)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Frame4"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "chkAuto"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdSel(4)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmdSel(5)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "spdResult1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "spdRstview"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   " ▒   받은 결과     "
      TabPicture(1)   =   "frmComm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPanel3"
      Tab(1).Control(1)=   "spdResult2"
      Tab(1).Control(2)=   "cmdResultPrint"
      Tab(1).Control(3)=   "tblexcel"
      Tab(1).Control(4)=   "SSPanel"
      Tab(1).Control(5)=   "cmdExcel"
      Tab(1).Control(6)=   "CommonDialog1"
      Tab(1).Control(7)=   "cmdAppend(1)"
      Tab(1).Control(8)=   "chkExcel"
      Tab(1).Control(9)=   "cmdSel(0)"
      Tab(1).Control(10)=   "cmdSel(1)"
      Tab(1).ControlCount=   11
      Begin FPSpread.vaSpread spdRstview 
         Height          =   7995
         Left            =   12150
         TabIndex        =   98
         Top             =   240
         Visible         =   0   'False
         Width           =   2325
         _Version        =   196608
         _ExtentX        =   4101
         _ExtentY        =   14102
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   6
         MaxRows         =   30
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":0038
         UserResize      =   0
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   7980
         Left            =   60
         TabIndex        =   99
         Top             =   390
         Width           =   15105
         _Version        =   196608
         _ExtentX        =   26644
         _ExtentY        =   14076
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   5
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   9
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm.frx":07FF
         UserResize      =   0
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   465
         Index           =   5
         Left            =   360
         TabIndex        =   96
         Top             =   420
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   820
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":0CC4
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   465
         Index           =   4
         Left            =   90
         TabIndex        =   97
         Top             =   420
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   820
         _StockProps     =   78
         ForeColor       =   14735310
         BevelWidth      =   1
         Picture         =   "frmComm.frx":1146
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   345
         Index           =   1
         Left            =   -74580
         TabIndex        =   94
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   609
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":15B4
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   345
         Index           =   0
         Left            =   -74850
         TabIndex        =   95
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   609
         _StockProps     =   78
         ForeColor       =   14735310
         BevelWidth      =   1
         Picture         =   "frmComm.frx":1A36
      End
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  '평면
         Caption         =   "Auto Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   12120
         TabIndex        =   92
         Top             =   30
         Value           =   1  '확인
         Width           =   1320
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   7065
         Left            =   3870
         TabIndex        =   75
         Top             =   2370
         Visible         =   0   'False
         Width           =   5175
         Begin VB.CommandButton Command1 
            Caption         =   "TEST"
            Height          =   375
            Left            =   1920
            TabIndex        =   91
            Top             =   1140
            Width           =   1230
         End
         Begin VB.TextBox txtURL 
            Height          =   690
            Left            =   390
            MultiLine       =   -1  'True
            TabIndex        =   90
            Top             =   210
            Width           =   4605
         End
         Begin VB.TextBox txtInstCd 
            Height          =   270
            Left            =   390
            TabIndex        =   89
            Top             =   1110
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "테스트"
            Height          =   375
            Left            =   2580
            TabIndex        =   80
            Top             =   2520
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   480
            TabIndex        =   79
            Text            =   "219"
            Top             =   2550
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   480
            TabIndex        =   78
            Text            =   "20120122"
            Top             =   3000
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   480
            TabIndex        =   77
            Text            =   "20120122"
            Top             =   3450
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   2520
            TabIndex        =   76
            Top             =   3420
            Visible         =   0   'False
            Width           =   1605
         End
         Begin MSComctlLib.ListView lvwCuData 
            Height          =   4830
            Left            =   360
            TabIndex        =   81
            Top             =   2040
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   8520
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin InetCtlsObjects.Inet Inet1 
            Left            =   4200
            Top             =   1230
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
         End
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   330
         Index           =   3
         Left            =   1080
         TabIndex        =   68
         Top             =   1920
         Visible         =   0   'False
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   582
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":1EA4
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   330
         Index           =   2
         Left            =   750
         TabIndex        =   69
         Top             =   1920
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   582
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":2326
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1785
         Left            =   1110
         TabIndex        =   51
         Top             =   960
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   3149
         _StockProps     =   15
         BackColor       =   14737632
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
         Begin VB.TextBox txtToNo 
            Alignment       =   2  '가운데 맞춤
            Height          =   285
            Left            =   3900
            MaxLength       =   8
            TabIndex        =   67
            Top             =   90
            Width           =   1185
         End
         Begin VB.TextBox txtFrNo 
            Alignment       =   2  '가운데 맞춤
            Height          =   285
            Left            =   2550
            MaxLength       =   8
            TabIndex        =   66
            Top             =   90
            Width           =   1185
         End
         Begin VB.TextBox txtChart 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   300
            Left            =   4980
            MaxLength       =   12
            TabIndex        =   54
            Top             =   390
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.ComboBox cboChk 
            Height          =   300
            ItemData        =   "frmComm.frx":2794
            Left            =   5160
            List            =   "frmComm.frx":279E
            TabIndex        =   53
            Top             =   90
            Width           =   1095
         End
         Begin VB.ComboBox cboComNm 
            Height          =   300
            ItemData        =   "frmComm.frx":27B2
            Left            =   4590
            List            =   "frmComm.frx":27B4
            TabIndex        =   52
            Top             =   480
            Visible         =   0   'False
            Width           =   1725
         End
         Begin MSMask.MaskEdBox mskOrdtime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Left            =   4560
            TabIndex        =   55
            Top             =   450
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   57
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3750
            TabIndex        =   56
            Top             =   150
            Width           =   315
         End
      End
      Begin VB.TextBox txtDump 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   12840
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   7110
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Timer tmrOrder 
         Left            =   10260
         Top             =   -360
      End
      Begin BHButton.BHImageButton cmdPrevious 
         Height          =   330
         Left            =   90
         TabIndex        =   47
         Top             =   5400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         Caption         =   "◀"
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
         ForeColor       =   16711680
         BackColor       =   16711680
         AlphaColor      =   255
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdNext 
         Height          =   330
         Left            =   330
         TabIndex        =   48
         Top             =   5400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         Caption         =   "▶"
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
         TransparentPicture=   "frmComm.frx":27B6
         ForeColor       =   16711680
         BackColor       =   255
         AlphaColor      =   255
         ImgOutLineSize  =   3
      End
      Begin VB.CommandButton Command2 
         Caption         =   "TEST"
         Height          =   375
         Left            =   8100
         TabIndex        =   46
         Top             =   -30
         Visible         =   0   'False
         Width           =   1230
      End
      Begin BHButton.BHImageButton cmdPosNo 
         Height          =   375
         Left            =   5550
         TabIndex        =   39
         Top             =   -90
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "Pos변경"
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
      Begin BHButton.BHImageButton cmdOrder 
         Height          =   360
         Left            =   6810
         TabIndex        =   35
         Top             =   -30
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   635
         Caption         =   "오더전송"
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
      Begin VB.Frame Frame3 
         Height          =   315
         Left            =   90
         TabIndex        =   44
         Top             =   960
         Width           =   555
      End
      Begin VB.CheckBox chkExcel 
         Appearance      =   0  '평면
         BackColor       =   &H80000004&
         Caption         =   "Excel 생성"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -61080
         TabIndex        =   43
         Top             =   30
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1245
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   465
         Left            =   7260
         TabIndex        =   32
         Top             =   5250
         Visible         =   0   'False
         Width           =   3075
         _Version        =   65536
         _ExtentX        =   5424
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   14737632
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
         Enabled         =   0   'False
         Begin VB.OptionButton optBar 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "병록번호"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1650
            TabIndex        =   34
            Top             =   90
            Width           =   1335
         End
         Begin VB.OptionButton optSeq 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "검사번호"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   210
            TabIndex        =   33
            Top             =   90
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   1
         Left            =   -69570
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   661
         Caption         =   "결과저장"
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
      Begin BHButton.BHImageButton cmdEot 
         Height          =   375
         Left            =   4350
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "초기화"
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
      Begin BHButton.BHImageButton cmdWordQuery 
         Height          =   390
         Left            =   9330
         TabIndex        =   37
         Top             =   5400
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   688
         Caption         =   "조회"
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
      Begin BHButton.BHImageButton cmdRackNo 
         Height          =   375
         Left            =   9270
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "Rack변경"
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
      Begin VB.TextBox txtResult 
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1500
         Left            =   3120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   40
         Top             =   6060
         Visible         =   0   'False
         Width           =   6600
      End
      Begin BHButton.BHImageButton cmdRequist 
         Height          =   390
         Index           =   2
         Left            =   7950
         TabIndex        =   42
         Top             =   5400
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   688
         Caption         =   "Last Order.."
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -63660
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   5385
         Left            =   8460
         TabIndex        =   10
         Top             =   8520
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   9499
         Bevel           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtCOM2 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4395
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   19
            Top             =   300
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   11
            Top             =   4635
            Width           =   5760
            Begin MSComDlg.CommonDialog cdlFile 
               Left            =   5265
               Top             =   60
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin HSCotrol.CButton cmdChksum 
               Height          =   360
               Left            =   2205
               TabIndex        =   12
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "SUM"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMOutput2 
               Height          =   360
               Left            =   1155
               TabIndex        =   13
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMClear2 
               Height          =   360
               Left            =   3600
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMInput2 
               Height          =   360
               Left            =   90
               TabIndex        =   15
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMLoad 
               Height          =   360
               Left            =   4635
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Load"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdACK 
               Height          =   360
               Left            =   3105
               TabIndex        =   17
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ACK"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdENQ 
               Height          =   360
               Left            =   2655
               TabIndex        =   18
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ENQ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
         End
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   12030
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   1500
      End
      Begin FPSpread.vaSpread spdWorklist 
         Height          =   4440
         Left            =   120
         TabIndex        =   45
         Top             =   1020
         Visible         =   0   'False
         Width           =   4755
         _Version        =   196608
         _ExtentX        =   8387
         _ExtentY        =   7832
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   7
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   8
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":2C28
         UserResize      =   2
      End
      Begin BHButton.BHImageButton cmdStartNo 
         Height          =   420
         Left            =   3660
         TabIndex        =   49
         Top             =   -150
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "시작번호변경"
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
      Begin BHButton.BHImageButton cmdPrint 
         Height          =   360
         Left            =   10320
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         Caption         =   "WorkList 출력"
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
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   420
         Left            =   -67620
         TabIndex        =   59
         Top             =   420
         Visible         =   0   'False
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   741
         Caption         =   "Excel 파일 생성 / 출력"
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
      Begin Threed.SSPanel SSPanel 
         Height          =   465
         Left            =   -67140
         TabIndex        =   60
         Top             =   300
         Visible         =   0   'False
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   820
         _StockProps     =   15
         BackColor       =   14737632
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
         Begin VB.ComboBox cboOrdDt 
            Height          =   300
            ItemData        =   "frmComm.frx":3165
            Left            =   4590
            List            =   "frmComm.frx":3167
            TabIndex        =   71
            Top             =   90
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.ComboBox cboChk1 
            Height          =   300
            ItemData        =   "frmComm.frx":3169
            Left            =   2610
            List            =   "frmComm.frx":3173
            TabIndex        =   70
            Top             =   90
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.ComboBox cboRstgbn 
            Height          =   300
            Index           =   1
            ItemData        =   "frmComm.frx":3187
            Left            =   5850
            List            =   "frmComm.frx":3194
            Style           =   2  '드롭다운 목록
            TabIndex        =   62
            Top             =   75
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "frmComm.frx":31BE
            Left            =   4590
            List            =   "frmComm.frx":31C0
            TabIndex        =   61
            Top             =   480
            Visible         =   0   'False
            Width           =   1725
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   4
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   4560
            TabIndex        =   63
            Top             =   450
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "접수일 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3780
            TabIndex        =   72
            Top             =   150
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   64
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin FPSpread.vaSpread tblexcel 
         Height          =   675
         Left            =   -64590
         TabIndex        =   65
         Top             =   300
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
         SpreadDesigner  =   "frmComm.frx":31C2
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   4725
         Left            =   210
         TabIndex        =   25
         Top             =   2430
         Visible         =   0   'False
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   8334
         Bevel           =   1
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
         Begin VB.TextBox txtCom 
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3720
            Left            =   -8430
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   26
            Top             =   360
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   -7800
            TabIndex        =   27
            Top             =   3720
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Save"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMOutput 
               Height          =   360
               Left            =   1155
               TabIndex        =   29
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMClear 
               Height          =   360
               Left            =   9450
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMInput 
               Height          =   360
               Left            =   90
               TabIndex        =   31
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
         End
      End
      Begin BHButton.BHImageButton cmdWorkList 
         Height          =   360
         Left            =   3270
         TabIndex        =   74
         Top             =   960
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   635
         Caption         =   "WorkList 조회"
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
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   315
         Left            =   1170
         TabIndex        =   82
         Top             =   510
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94961664
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   315
         Left            =   4710
         TabIndex        =   83
         Top             =   510
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         _Version        =   393216
         Format          =   94961664
         CurrentDate     =   40248
      End
      Begin BHButton.BHImageButton cmdSearch 
         Height          =   360
         Left            =   8760
         TabIndex        =   84
         Top             =   480
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         Caption         =   "WorkList 조회"
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
      Begin BHButton.BHImageButton cmdResultPrint 
         Height          =   360
         Left            =   -62940
         TabIndex        =   86
         Top             =   480
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   635
         Caption         =   "결과출력"
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
      Begin BHButton.BHImageButton cmdMake 
         Height          =   360
         Left            =   7200
         TabIndex        =   88
         Top             =   480
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   635
         Caption         =   "WorkList 생성"
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
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   300
         Index           =   0
         Left            =   13560
         TabIndex        =   93
         Top             =   0
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         Caption         =   "서버등록"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         ImgOutLineSize  =   3
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7350
         Left            =   -74850
         TabIndex        =   100
         Top             =   900
         Width           =   15015
         _Version        =   196608
         _ExtentX        =   26485
         _ExtentY        =   12965
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   7
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   8
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm.frx":336D
         UserResize      =   0
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   525
         Left            =   -74850
         TabIndex        =   101
         Top             =   330
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   926
         _StockProps     =   15
         BackColor       =   14737632
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
         Begin VB.ComboBox cboRstgbn 
            Height          =   300
            Index           =   0
            ItemData        =   "frmComm.frx":380D
            Left            =   2235
            List            =   "frmComm.frx":381A
            Style           =   2  '드롭다운 목록
            TabIndex        =   102
            Top             =   135
            Width           =   1410
         End
         Begin MSMask.MaskEdBox mskRstDate 
            Height          =   300
            Left            =   1110
            TabIndex        =   103
            Top             =   135
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin BHButton.BHImageButton cmdRstQuery 
            Height          =   375
            Left            =   3735
            TabIndex        =   104
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   661
            Caption         =   "조회"
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
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   106
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "검사일자 :"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   105
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "작업일자 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3660
         TabIndex        =   87
         Top             =   570
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "접수일자 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   85
         Top             =   570
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         DrawMode        =   5  '카피 펜이 아님
         X1              =   4650
         X2              =   10350
         Y1              =   5940
         Y2              =   5940
      End
      Begin VB.Label Label8 
         Caption         =   "● Information List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   7950
         TabIndex        =   41
         Top             =   5790
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "재검/QC :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   8070
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   8850
      Top             =   5160
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   6690
      Top             =   5190
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7860
      Top             =   5130
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8340
      Top             =   5130
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   7230
      Top             =   5130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3844
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":3DDE
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":4378
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":4912
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":4EAC
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":5446
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   5715
      Top             =   5130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":59E0
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":5F7A
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":6514
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":6AAE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":7340
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":749A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm.frx":75F4
            Key             =   "NOF"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   45
      TabIndex        =   1
      Top             =   9015
      Width           =   15315
      Begin VB.Timer tmrDummy 
         Enabled         =   0   'False
         Left            =   13380
         Top             =   120
      End
      Begin VB.Timer tmrOk 
         Left            =   14400
         Top             =   180
      End
      Begin VB.Timer tmrWorking 
         Interval        =   100
         Left            =   12930
         Top             =   150
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   12150
         Top             =   150
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   6795
         TabIndex        =   20
         Top             =   -90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Run"
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
         Left            =   7920
         TabIndex        =   21
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Stop"
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
         Left            =   9210
         TabIndex        =   22
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
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
         Left            =   10530
         TabIndex        =   23
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
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
         TransparentPicture=   "frmComm.frx":774E
         ImgOutLineSize  =   3
      End
      Begin VB.Image imgBack 
         BorderStyle     =   1  '단일 고정
         Height          =   1050
         Index           =   0
         Left            =   4020
         Picture         =   "frmComm.frx":8FD8
         Stretch         =   -1  'True
         Top             =   -210
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Image imgLogo 
         Height          =   240
         Index           =   0
         Left            =   3630
         Picture         =   "frmComm.frx":A7AB
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "팝업용 ==>"
         Height          =   225
         Index           =   1
         Left            =   2700
         TabIndex        =   58
         Top             =   210
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "작업대기 중.."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   960
         TabIndex        =   6
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " 상태 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   5
         Top             =   225
         Width           =   615
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   873
      Border          =   1
      CaptionBackColor=   16777215
      Caption         =   " Communication"
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Receive : "
         Height          =   180
         Left            =   13875
         TabIndex        =   4
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   12840
         TabIndex        =   3
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Index           =   0
         Left            =   11745
         TabIndex        =   2
         Top             =   165
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14745
         Picture         =   "frmComm.frx":AD35
         Top             =   135
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13455
         Picture         =   "frmComm.frx":B2BF
         Top             =   135
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12255
         Picture         =   "frmComm.frx":B849
         Top             =   135
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_KEY       As String = "K"
Private Const COL_EQP_NUM   As String = "EQP_ID"

Private Const KEY_SEQ       As String = "KEY_SEQ"   ' "순서"
Private Const KEY_PTID      As String = "KEY_PTID"  ' "등록번호"
Private Const KEY_PTNM      As String = "KEY_PTNM"  ' "성  명"
Private Const KEY_SPCNO     As String = "KEY_SPCNO" ' "검체번호"
Private Const KEY_EQPNO     As String = "KEY_EQPNO" ' "검체번호"
Private Const KEY_STAT      As String = "KEY_STAT"  ' "상 태"
Private Const KEY_TEST      As String = "KEY_TEST"  ' "검사항목"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름
Private Const TEST_VALUES   As String = "VALUES"    '결과

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const fs  As String = ""
Const Rs  As String = ""

Const Field_      As String = "|"
Const Repeat_     As String = "\"
Const Component_  As String = "^"
Const Escape_     As String = "&"
Const Slash_      As String = "/"
Dim cntField_     As Integer '|
Dim cntRepeat_    As Integer '\
Dim cntComponent_ As Integer '^
Dim cntEscape_    As Integer '&
Dim cntSlash_     As Integer '/



Dim sStxCheck As Integer
Dim sEtxCheck As Integer
Dim sLfCheck  As Integer
Dim sCrcheck  As Integer
' --------------------------------------------------------------
Dim strOrdLst As String

Dim ELEC1010(100)   As String
Dim fELEC1010       As Variant
Dim fELEC1010_1     As Variant
Dim fELEC1010_2     As Variant
Dim fELEC1010_3     As Variant
Dim SendData(10)     As String
Dim SendCount        As String
Dim Or_Seq           As Integer
Dim SendBuffW           As String
Dim SendBuffT           As String
Dim IntRow          As Integer
Dim brStr           As String

Dim cntCheckSum      As Integer
Dim ReceiveData      As String

Dim SendFlg          As Boolean
Dim HostOutput       As String

Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult

Private mAdoRs      As ADODB.Recordset
Private mAdoRs1     As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer     As String
Private f_strJOB_FLAG   As String
Private f_strOrdList    As String
Private f_intSampleNo   As Integer

Private f_blnWorkList   As Boolean
Private f_lngWork_Row   As Long

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

Dim fTBA40FR(50) As String
Dim fCellDynSize(50, 1) As Integer
Dim fChannel() As String
Dim pName   As String
Dim pNo     As String
Dim chkEnq  As Integer

Dim flgETB           As Boolean
Dim flgETX           As Boolean

Private Type SugaMatch
    TestId          As String
    Sugacd          As String
    Testcd          As String
    DecPoint        As Long
End Type

Dim SMatch() As SugaMatch
Dim CountTest As Integer, sErrorFlag As Boolean

Private Type TYPE_CD
    strEqpCd    As String
    intCnt      As Integer
    strTestcd(50) As String
End Type

Private f_typCode() As TYPE_CD

Dim RecordChk As Boolean

Dim strGumCd As String
Dim strJinCd As String
Dim fRcvString As String

Dim PatientID As String    'Q Message Pattern Check
Dim PatientSeq As String
Dim PatientDisk As String
Dim PatientRack As String
Dim PatientPos As String

Dim SeqNo As String
'Dim RecordChk   As Boolean

Dim G_CLVALU    As String
Dim G_CHVALU    As String
Dim G_EVALUATE  As String
Dim G_PANIC     As String
Dim G_DELTA     As String
Dim strFrameNo  As Integer
Dim OrderCnt As Integer
Dim vRow As Integer
Dim sPatiant_No As String


Private Type typeElecsys2010
    TestDate      As String
    TestTime      As String
    RunType       As String 'N, E, R, C, S, B
    SampleNo      As String
    SID           As String 'Sid
    SampleTy      As String '1~5
    RackNo        As String
    Position      As String '1~5
    priority      As String
    TestId(50)   As String
    Result(50)   As String
    Status(50)   As String
    Rerun(50)    As String
End Type

Dim Elecsys2010 As typeElecsys2010
Dim fElecsys2010(100) As String
Dim fElecsys2010_1(100) As String

Dim OrderSort_Flag As Integer
Dim Patiant_Recevid As Boolean

Dim gspdResultRow  As Integer

'-- 2010.03.11 osw 추가 : 검사결과 팝업메세지
'Private WithEvents mobjPopups   As PopUpMessages

'Private mobjDefault             As PopUpMessage

'-- Interface Class
Private cInterface              As New clsInterface
Private objIntInfo              As clsIntInfo           '검체정보 클래스
Private objOrder                As New clsIntOrder          '오더정보 클래스
Private objResult               As clsIntResults        '결과정보
Private objIntNm                As New clsIntTest       '검사정보
'Private objErrInfo              As clsErrInfo            '알람정보 클래스


Const SPCLEN As Integer = "11"
Dim mFrameNo As Integer

Dim strQCResult  As String   '수신한 QC결과
Dim strQC_LCResult  As String   '수신한 QC결과
Dim strQC_HCResult  As String   '수신한 QC결과
Dim strAlarm As String
Dim lngDummyTime As Long

Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer

Dim strSNOCode(100) As String
'Dim gwTmp1 As String

Const colBANO = 1   '바코드번호
Const colORDT = 2   '처방일자
Const colORQN = 3   '처방번호
Const colPANM = 4   '환자명
Const colPAID = 5   '병록번호
Const colOIFL = 6   '입/외구분
Const colSENO = 7   '일련번호
Const colSEXS = 8   '성별
Const colAGES = 9   '나이
Const colNWNO = 10  '내원번호
Const colSQNO = 11  'SeqNo



Private Sub cmdCLR_Click()
    
    txtDump.text = ""

End Sub

Private Sub cmdWorkPrint_Click()
    
    With spdResult1
        .Col = 2
        .Row = 8
        .Col2 = .maxrows
        .Row2 = 8
        .BlockMode = True
    
'        .CellBorderType = 1
'        .CellBorderStyle = SS_BORDER_STYLE_SOLID
'        .CellBorderColor = RGB(0, 0, 0)
        .Action = ActionPrint
    
        .BlockMode = False

    End With
    'spdResult1.Action = ActionPrint

End Sub

Private Sub cmdMake_Click()
    Dim strDoc      As String

'GoTo Rst
    strDoc = AS400DOWNF(INS_CODE, Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
    If strDoc = "0" Then
        'MsgBox "성공"
    ElseIf strDoc = "44" Then
        '-- 장비코드 오류
        MsgBox "장비코드 설정을 확인하세요", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    ElseIf strDoc = "48" Then
        '시작일오류
        MsgBox "작업일자 시작일을 확인하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    ElseIf strDoc = "52" Then
        '종료일오류
        MsgBox "작업일자 종료일을 확인하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If

End Sub

Private Sub cmdResultPrint_Click()
Dim objclsCommon As New clsCommon

Dim Tmp_Testnm As String
Dim Row_cnt As Integer, Col_cnt As Integer, TmpPrintline As Integer
Dim vTmp As Variant
Dim stragesex As String
Dim varTmp As Variant

Const TmpLine = "──────────────────────────────────────────────────────────────────────────────────────────────"
Const TmpLine1 = "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

    
    If spdResult2.maxrows >= 1 Then
        Printer.Orientation = 2 '가로
        With objclsCommon
            .PrintText 15, 3, Format(Date, "yyyy/mm/dd") & "  ResultList Report..( " & App.EXEName & " )", "Arial", 12
            
            .PrintText 0.5, 5, TmpLine
            .PrintText 0.5, 6, "순", , 9
            .PrintText 2, 6, "접수일자", , 9
            .PrintText 7, 6, "환자성명", , 9
            .PrintText 12, 6, "접수번호", , 9
            .PrintText 0.5, 7, TmpLine
            
            '-- 검사명찍기
            For Col_cnt = 11 To spdResult2.MaxCols
                spdResult2.GetText Col_cnt, 0, varTmp
                .PrintText 14 + ((Col_cnt - 10) * 4), 6, CStr(varTmp)
            Next
            
            TmpPrintline = 8
        
            For Row_cnt = 1 To spdResult2.maxrows
                spdResult2.Row = Row_cnt
                
                If (Row_cnt Mod 22) <> 0 Then
                                        .PrintText 0.5, TmpPrintline, Row_cnt, , 9                          ' 순
                    spdResult2.Col = 3: .PrintText 2, TmpPrintline, Mid(spdResult2.text, 3), , 9                    ' 처방일자
                    spdResult2.Col = 10: .PrintText 7, TmpPrintline, Trim(spdResult2.text), 9              ' 검체번호
                    spdResult2.Col = 4: .PrintText 12, TmpPrintline, Trim(spdResult2.text), , 9             ' 이    름
                    
                    For Col_cnt = 11 To spdResult2.MaxCols
                        spdResult2.Row = Row_cnt:            spdResult2.Col = Col_cnt
                        .PrintText 14 + ((Col_cnt - 10) * 4), TmpPrintline, spdResult2.text, , 7.5
                    Next Col_cnt
                    .PrintText 0.5, TmpPrintline + 1, TmpLine1
                    
                    TmpPrintline = TmpPrintline + 2
                    Tmp_Testnm = ""
                Else
                
                    '-------------------------------------------------------
                
                                        .PrintText 0.5, TmpPrintline, Row_cnt, , 9                          ' 순
                    spdResult2.Col = 3: .PrintText 2, TmpPrintline, Mid(spdResult2.text, 3), , 9
                    spdResult2.Col = 10: .PrintText 7, TmpPrintline, Trim(spdResult2.text), 9
                    spdResult2.Col = 4: .PrintText 12, TmpPrintline, Trim(spdResult2.text), , 9
                    
                    
                    For Col_cnt = 11 To spdResult2.MaxCols
                
                        spdResult2.Row = Row_cnt:            spdResult2.Col = Col_cnt
                        .PrintText 14 + ((Col_cnt - 10) * 4), TmpPrintline, spdResult2.text, , 7.5
                        
                    Next Col_cnt
                    .PrintText 0.5, TmpPrintline + 1, TmpLine1
                    
                    TmpPrintline = TmpPrintline + 2
                    Tmp_Testnm = ""
                    
                    '-------------------------------------------------------
                
                        .PrintText 0.5, TmpPrintline, TmpLine
                        .PrintText 1, TmpPrintline + 1, "── Next Report ──", , 9, True
                        Printer.NewPage
                        
                        .PrintText 0.5, 5, TmpLine
                        .PrintText 0.5, 6, "순", , 9
                        .PrintText 2, 6, "접수일자", , 9
                        .PrintText 7, 6, "환자성명", , 9
                        .PrintText 12, 6, "접수번호", , 9
                        '.PrintText 16, 6, "장비결과", , 9
                        .PrintText 0.5, 7, TmpLine
                        
                        For Col_cnt = 11 To spdResult2.MaxCols
                            spdResult2.GetText Col_cnt, 0, varTmp
                            .PrintText 14 + ((Col_cnt - 10) * 4), 6, CStr(varTmp)
                        Next
                        
                        TmpPrintline = 9
                End If
            
            Next Row_cnt
            
            .PrintText 0.5, TmpPrintline, TmpLine
            .PrintText 1, TmpPrintline + 1, "── End of Report ──", , 9, True
        
        End With
        
        Printer.NewPage
        Printer.EndDoc
        
        MsgBox Format(Date, "yyyy/mm/dd") & "일자의 " & App.EXEName & "의 장비 검사 WorkList가 Print되었습니다..       " & vbCrLf & vbCrLf & "다음 작업을 진행하십시요..", vbInformation + vbOKOnly, App.Title
    Else
        MsgBox Format(Date, "yyyy/mm/dd") & "일자의 " & App.EXEName & "의 장비 검사 WorkList가  Load 되어 있지 않습니다..       " & vbCrLf & vbCrLf & "자료를 확인 하십시요..", vbInformation + vbOKOnly, App.Title
    End If
    
    '
    ' 마지막 저장
    '
    spdResult2.SaveTabFile App.Path & "\" & REG_INSNAME & "_Request.txt"
    

End Sub

Private Sub Command2_Click()

'''Dim oSOAP As MSSOAPLib30.SoapClient30
'''Dim send As Long
'''
'''    Set oSOAP = New SoapClient30
'''
'''
'''
'''    oSOAP.ClientProperty("ServerHTTPRequest") = True
'''    oSOAP.MSSoapInit "http://test2.i-mftg.com:8888/CoChainWebService/WSCoChain.asmx?wsdl" '      <== 이부분 오류
'''
'''
'''    send = oSOAP.Adminlist("", "")
'''
'''    send = oSOAP.SetActualResult("bQ8GFj48eww=", "07178", 100, "2007-09-27")
'''
'''    DoEvents

'''    Dim client As SoapClient30
'''    Dim send As String
'''
'''    Set client = New SoapClient30
'''    client.ClientProperty("ServerHTTPRequest") = True
'''
'''
'''
'''    client.MSSoapInit ("http://loclahost/WebService.asmx?wsdl")
'''
'''
'''    send = client.Adminlist("", "")
'''
'''    Dim strSendData As String
'''
'''    '기본uRL
'''    strSendData = "http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?"
'''    '1. 인터페이스로그인 사용자 이름 조회
'''    strSendData = strSendData & "submit_id=TRLII00104&business_id=lis&ex_interface=12345678|01&instcd=012&userid=12345678&"
'''    '==> output : usernm  사용자이름
'''    '2. 장비정보조회
'''    'strSendData = strSendData & "submit_id=TRLII00103&business_id=lis&ex_interface=12345678|01&refgbn=2&eqmtcd=I24&instcd=012&"
'''
'''    '3. 워크리스트조회
'''    'strSendData = strSendData & "submit_id=TRLII00119&business_id=lis&ex_interface=12345678|01&instcd=012&eqmtcd=I24&bcno=I24NK1BR0&"
'''
'''    '4. 결과저장
'''    'strSendData = strSendData & "submit_id=TXLII00101&business_id=lis&ex_interface=12345678|01&bcno=I24NK1BR0&eqmtcd=I24&instcd=012&userid=12345678&"
'''
'''    Winsock1.SendData strSendData
'''
    Dim sHtmlLine
    Dim sUrl, sPost, sParam As String
    Dim sRcvData, sData As String

        sUrl = txtURL.text
        
        '-- 검사코드
                 sParam = "submit_id=TRLII00103&"
        sParam = sParam & "business_id=lis&"
        sParam = sParam & "ex_interface=" & CurrUser.CuUserID & "|" & Trim(txtInstCd.text) & "&" '사용자ID|기관코드
        sParam = sParam & "refgbn=2&"                         '장비코드
        sParam = sParam & "eqmtcd=" & INS_CODE & "&"                         '장비코드
'        sParam = sParam & "eqmtcd=I12&"                         '장비코드
        sParam = sParam & "instcd=" & Trim(txtInstCd.text)                                  '기관코드
        
        '-- QC
'                 sParam = "submit_id=TXLQI00102&"
'        sParam = sParam & "business_id=lis&"
'        sParam = sParam & "ex_interface=" & CurrUser.CuUserID & "|" & Trim(txtInstCd.text) & "&" '사용자ID|기관코드
'        sParam = sParam & "ifid=2&"                         '장비코드
'        sParam = sParam & "result=2&"                         '장비코드
''        sParam = sParam & "eqmtcd=" & INS_CODE & "&"                         '장비코드
'        sParam = sParam & "eqmtcd=I12&"                         '장비코드
'        sParam = sParam & "instcd=" & Trim(txtInstCd.text) & "&"                                 '기관코드
'        sParam = sParam & "userid=" & CurrUser.CuUserID
        
        sRcvData = OpenURLWithIE2(sUrl & sParam, Inet1)
        
            
        If InStr(1, sRcvData, "<?xml version") > 0 Then
            'gwTmp1 = ""
        End If
        
        'gwTmp1 = gwTmp1 & sRcvData
        
'<?xml version='1.0' encoding='utf-8'?>
'<root><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd>
'<testcdtodd>99991231</testcdtodd>
'<dispseq>1</dispseq>
'<tclscd>LIA196</tclscd>
'<testdispnm><![CDATA[SLE Profile]]></testdispnm>
'<spccd>007</spccd>
'<ifreqcd></ifreqcd>
'<ifrsltcd></ifrsltcd>
'<ifspccd></ifspccd>
'<qctestcd></qctestcd>
'<rsltreptkind></rsltreptkind>
'<etc1></etc1>
'<etc2></etc2>
'<testcdseqno>1</testcdseqno>
'</eqmttclscdlist><eqmttclscdlist>
'
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>2</dispseq><tclscd>LIA19601</tclscd><testdispnm><![CDATA[RNP/Sm]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd>
'<qctestcd></qctestcd>
'<rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>2</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>3</dispseq><tclscd>LIA19602</tclscd><testdispnm><![CDATA[Sm]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>3</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>4</dispseq><tclscd>LIA19603</tclscd><testdispnm><![CDATA[SS-A Native (60 KDa)]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>4</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>5</dispseq><tclscd>LIA19604</tclscd><testdispnm><![CDATA[Ro-52 (52 kDa)]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>5</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>6</dispseq><tclscd>LIA19605</tclscd><testdispnm><![CDATA[SS-B]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>6</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>7</dispseq><tclscd>LIA19606</tclscd><testdispnm><![CDATA[Scl-70]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>7</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>8</dispseq><tclscd>LIA19608</tclscd><testdispnm><![CDATA[Jo-1]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>8</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>9</dispseq><tclscd>LIA19609</tclscd><testdispnm><![CDATA[Centromere B]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>9</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>10</dispseq><tclscd>LIA19611</tclscd><testdispnm><![CDATA[dsDNA]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>10</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>11</dispseq><tclscd>LIA19614</tclscd><testdispnm><![CDATA[Ribosamal-P Protein]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>11</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>12</dispseq><tclscd>LIA19616</tclscd><testdispnm><![CDATA[Chromatin]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>12</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>13</dispseq><tclscd>LIA19617</tclscd><testdispnm><![CDATA[RNP (A)]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>13</testcdseqno></eqmttclscdlist><eqmttclscdlist>
'<testcdfromdd>20120308</testcdfromdd><testcdtodd>99991231</testcdtodd><dispseq>14</dispseq><tclscd>LIA19618</tclscd><testdispnm><![CDATA[PNP (68)]]></testdispnm><spccd>007</spccd><ifreqcd></ifreqcd><ifrsltcd></ifrsltcd><ifspccd></ifspccd><qctestcd></qctestcd><rsltreptkind></rsltreptkind><etc1></etc1><etc2></etc2><testcdseqno>14</testcdseqno></eqmttclscdlist>
'
'<resultKM error="no" type="status" clear="true" description="info||정상적으로 처리되었습니다." updateinstance="true" source="1331627467915"/>
'</root>

End Sub

Private Sub Command3_Click()
    Dim Arr()   As Byte
    Dim strTmp  As String
    Dim sDATE As String
    Dim eDate As String
    
'    sDATE = "20120122"
'    eDate = "20120122"
    
    'strTmp = AS400DOWNF(Text1.text, Text2.text, Text3.text)
    
    MsgBox AS400DOWNF(Text1.text, Text2.text, Text3.text)
    
    If strTmp = "0" Then
        MsgBox "성공"
    ElseIf strTmp = "44" Then
        '장비코드오류
        MsgBox "장비코드오류"
    ElseIf strTmp = "48" Then
        '시작일오류
        MsgBox "시작일오류"
    ElseIf strTmp = "52" Then
        '종료일오류
        MsgBox "종료일오류"
    End If
    
End Sub

''-- 2010.03.11 osw 추가 : 검사결과 팝업메세지
'Private Sub AddPopup(ByVal strSPnm As String, ByVal strSPid As String)
'Dim objPopUp    As PopUpMessage
'
'    Set objPopUp = New PopUpMessage
'    With objPopUp
'        .Caption = INS_NAME
'        .Message = strSPnm & "(" & strSPid & ") 님" & vbCrLf & vbCrLf & " 검사결과 전송성공" & vbCrLf & ""
'        .Clickable = False
'        .Sticky = False
'        Set .Background = imgBack.Item(0)
'        Set .Logo = imgLogo.Item(0)
'        .WavFile = App.Path & "\sounds\type.wav"
'    End With
'    mobjPopups.Show objPopUp
'
'End Sub

'
'Private Sub dtpRsltDay_Change()
'Dim sqlDoc As String
'Dim AdoRs   As New ADODB.Recordset
'
'    sqlDoc = "Select DISTINCT mid(SPCNO,2,6) as SPCNO " & _
'             "  From INTERFACE003" & _
'             " Where TRANSDT = '" & Format(mskRstDate.text, "yyyymmdd") & "'" & _
'             "   And EQUIPCD = '" & INS_CODE & "'"
'
'    If cboRstgbn(1).ListIndex = 0 Then
'        sqlDoc = sqlDoc & "   And SERVERGBN = ''"
'    ElseIf cboRstgbn(1).ListIndex = 1 Then
'        sqlDoc = sqlDoc & "   And SERVERGBN = 'Y'"
'    End If
'
''    '-- 내/외국인 구분
''    If cboChk1.ListIndex = 0 Then
''        sqlDoc = sqlDoc & "   And IOFLAG = '0' "
''    ElseIf cboChk1.ListIndex = 1 Then
''        sqlDoc = sqlDoc & "   And IOFLAG = '1' "
''    End If
'
'    'sqlDoc = sqlDoc & " Order By SPCNO"
'    cboOrdDt.Clear
'
'    AdoRs.CursorLocation = adUseClient
'    AdoRs.Open sqlDoc, AdoCn_Jet
'    If AdoRs.RecordCount > 0 Then AdoRs.MoveFirst
'    Do While Not AdoRs.EOF
'        cboOrdDt.AddItem "20" & AdoRs.Fields("SPCNO")
'        AdoRs.MoveNext
'    Loop
'
'    If AdoRs.RecordCount > 0 Then cboOrdDt.ListIndex = 0
'
'End Sub

''-- 2010.03.11 osw 추가 : 검사결과 팝업메세지
'Private Sub Form_Unload(Cancel As Integer)
'    Set mobjPopups = Nothing
'    Set mobjDefault = Nothing
'End Sub
'
''-- 2010.03.11 osw 추가 : 검사결과 팝업메세지
'Private Sub SetupDefaultPopup()
'    Set mobjDefault = New PopUpMessage
'    With mobjDefault
'        Set .Background = imgBack.Item(1)
'        .ForeColor = vbWhite
'        Set .Logo = imgLogo.Item(1)
'        .WavFile = App.Path & "\newemail.wav"
'        .Caption = "New Email"
'        .Message = "You have received" & vbCrLf & "4 new emails." & vbCrLf & "Downloading..."
'        .Clickable = True
'        .ProgressBar = True
'    End With
'End Sub

Private Function f_funGet_ConvertResult(ByVal strRstval As String) As String

    Dim intPos  As Integer
    Dim strTmp1 As String, strTmp2  As String
    
    intPos = InStr(strRstval, "E")
    If intPos > 0 Then
        strTmp1 = Mid$(strRstval, 1, intPos - 1)
        strTmp2 = Mid$(strRstval, intPos + 1)
        
        If Mid$(strTmp2, 1, 1) = "-" Then
            f_funGet_ConvertResult = Round(Val(strTmp1) * (0.1 ^ Val(Mid$(strTmp2, 2))), 2)
        Else
            f_funGet_ConvertResult = Round(Val(strTmp1) * (10 ^ Val(Mid$(strTmp2, 2))), 2)
        End If
    Else
        f_funGet_ConvertResult = strRstval
    End If
    
End Function

Private Function MakeCS(Source As String) As String
    Dim X      As Long
    Dim ChkCS  As String
    Dim SumCS  As String
    Dim AddCS  As Long
    For X = 1 To Len(Source)
        AddCS = AddCS + Asc(Mid(Source, X, 1))
    Next X
    SumCS = Hex(AddCS)
    ChkCS = Mid(SumCS, Len(SumCS) - 1, 1)
    ChkCS = ChkCS & Right(SumCS, 1)
    MakeCS = ChkCS
End Function

Private Function f_funGet_SpreadRow(ByVal objSpd As vaSpread, ByVal intCol As Integer, _
                                    ByVal strPara As String) As Integer

    Dim varTmp  As Variant
    Dim IntRow  As Integer
    
    f_funGet_SpreadRow = 0
    
    With objSpd
        For IntRow = 1 To .maxrows
            .GetText intCol, IntRow, varTmp
            If Trim$(varTmp) = strPara Then
                f_funGet_SpreadRow = IntRow
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub f_subGet_JobList(ByVal strKeyno As String, ByRef strOrder As String, _
                             ByRef intOrdCnt As Integer, ByRef strSpec As String, _
                             ByRef strPcFlag As String)

    Dim adoRS1  As New ADODB.Recordset
    Dim adoRS2  As New ADODB.Recordset
    Dim sqlDoc  As String
    
    strOrder = "":  strPcFlag = "  ":   strSpec = "SE": intOrdCnt = 0
    sqlDoc = "select ORD_CODE, CHART_NO From L3A01" & _
             " where SAMPLE_DATE = '" & Mid$(strKeyno, 1, 8) & "'" & _
             "   and SAMPLE_SEQ  = " & Format(Mid$(strKeyno, 9, 3), "##0") & "" & _
             "   and PART        = '" & Mid$(strKeyno, 12, 2) & "'"
    adoRS1.CursorLocation = adUseClient
    adoRS1.Open sqlDoc, AdoCn_SQL
    If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
    
    sqlDoc = "select TESTCD_EQP, TESTCD, REMARK, AUTOVERIFY from INTERFACE002 where (EQP_CD = " & STS(INS_CODE) & ") AND (TESTCD <> '')"
    adoRS2.CursorLocation = adUseClient
    adoRS2.Open sqlDoc, AdoCn_Jet
    If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
    Do While Not adoRS2.EOF
        If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
        adoRS1.Find "ORD_CODE = " & STS(Trim(adoRS2("TESTCD") & ""))
        If Not adoRS1.EOF Then
            Select Case Trim(adoRS2(2) & "")
                Case "128": strSpec = "PL"
                Case Else:  strSpec = "SE"
            End Select
            
            If Trim(adoRS2("TESTCD_EQP") & "") = "XXX" Then
                strOrder = strOrder + "06A ," + Trim$(adoRS2("AUTOVERIFY") & "") + ",": strPcFlag = "PC"
            Else
                strOrder = strOrder + Trim(adoRS2("TESTCD_EQP") & "") + " ," + Trim$(adoRS2("AUTOVERIFY") & "") + ","
            End If
            intOrdCnt = intOrdCnt + 1
        End If
        adoRS2.MoveNext
    Loop
    adoRS2.Close:   Set adoRS2 = Nothing
    adoRS1.Close:   Set adoRS1 = Nothing
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
    
End Sub

Private Sub f_subGet_WorkList(ByRef strOrder As String, ByRef intOrdCnt As Integer, _
                              ByRef strSpec As String, ByRef strPcFlag As String, _
                              ByVal IntRow As Integer)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String

    Dim varTmp  As Variant
    Dim intCol  As Integer
    
    Dim itemX   As ListItems
    
    Set itemX = lvwCuData.ListItems
    
    strOrder = "":  strPcFlag = "  ": strSpec = "SE":   intOrdCnt = 0
    With spdWorklist
        For intCol = 5 To .MaxCols
            .Row = IntRow:  .Col = intCol
            If .BackColor = &HC6FEFF Then
                Select Case itemX.Item(intCol - 4).SubItems(11)
                    Case "128": strSpec = "PL"
                    Case Else:  strSpec = "SE"
                End Select
                .GetText intCol, 0, varTmp
                
                If itemX.Item(intCol - 4).tag = "XXX" Then
                    strOrder = strOrder + "06A ," + itemX.Item(intCol - 4).SubItems(10) + ",": strPcFlag = "PC"
                Else
                    strOrder = strOrder + itemX.Item(intCol - 4).tag + " ," & itemX.Item(intCol - 4).SubItems(10) + ","
                End If
                intOrdCnt = intOrdCnt + 1
            End If
        Next
    End With
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
   
End Sub

Private Sub f_subSet_ComCharacter()

    MSG_STX = Chr(COM_STX)
    MSG_ETX = Chr(COM_ETX)
    MSG_ENQ = Chr(COM_ENQ)
    MSG_EOT = Chr(COM_EOT)
    MSG_ACK = Chr(COM_ACK)
    MSG_NAK = Chr(COM_NACK)
    MSG_CR = Chr(COM_CR)
    MSG_LF = Chr(COM_LF)
    MSG_CRLF = Chr(COM_CR) & Chr(COM_LF)
    
End Sub

Private Sub f_subSet_ItemHeader()
    
    '검사코드 테이블
    With lvwCuData
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HideColumnHeaders = True
        With .ColumnHeaders
            .Clear
            Call .Add(, TEST_NM_EQP, "ID", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_CD_LIS, "검사코드", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "검 사 명", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "검사결과", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFL", "참고치(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFH", "참고치(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "재검", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "검체코드", (lvwCuData.Width - 310) * 0.5)
            Call .Add(, "TESTNO", "KIT코드", (lvwCuData.Width - 310) * 0.5)
            Call .Add(, "QCL", "QCLow코드", (lvwCuData.Width - 310) * 0.5)
            Call .Add(, "QCH", "QCHigh코드", (lvwCuData.Width - 310) * 0.5)
        End With
        .HideColumnHeaders = False
    End With
    
   
End Sub

Private Sub f_subSet_ItemComplete(lvw As Listview)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemH           As ColumnHeader
    Dim objHeadeItem    As clsCommon
    
    Dim intCol  As Integer
    
    lvw.ColumnHeaders.Clear
    Call lvw.ColumnHeaders.Add(, "EQP_ID", "검체 번호")
    
    intCol = 4
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) AS TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = '" & INS_CODE & "') AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With lvw
            .Enabled = True
            Set itemH = .ColumnHeaders.Add
            With itemH
                '컬럽 헤더키를 장비검사 코드로
                .Key = COL_KEY & Trim(adoRS.Fields("TESTCD_EQP") & "")
                '컬럽명은 검사 항목 이름
                .text = Trim(adoRS.Fields("TESTNM") & "")
                '테그는 검사 코드로
                .tag = Trim(adoRS.Fields("TESTCD") & "")
                .Width = 700
                .Alignment = lvwColumnCenter
            End With
            Set itemH = Nothing
        End With
        
        With spdWorklist
            intCol = intCol + 1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1:  .ColWidth(.MaxCols) = 6.5
            
            .SetText intCol, 0, adoRS.Fields("TESTNM")
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Function f_subSet_WorkList(ByVal strSchDate As String, ByVal strFrNo As String, ByVal strToNo As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim strYear     As String
    Dim strDate     As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
        strYear = Mid(strSchDate, 1, 4)
        If Mid(strFrNo, 1, 1) = "A" Then
            strFrNo = Mid(strFrNo, 8)
        End If
        If Mid(strToNo, 1, 1) = "A" Then
            strToNo = Mid(strToNo, 8)
        End If
        
        Set AdoRs_ORACLE = New ADODB.Recordset
        
                            '접수일자,접수번호,항목코드,부속코드,대표장비코드,'
                            '바코드ID,항목검체,작업일자(검사일자),작업번호,결과유형,
                            '단위,항목명,검체명,검사부서코드,수진자명,
                            '거래처명 , 성별, 연령, 검사담당자
                 sqlDoc = " Select C96JSDT, C96JSNO, C96GSCD, C96ATCD, C96DEVC,"
        sqlDoc = sqlDoc & "        C96BARC, C96KCCD, C96KSDT, C96GSSQ, C96TYPE,"
        sqlDoc = sqlDoc & "        C96UNTA, C96GSNM, C96KCNM, C96DECD, C96PTNM,"
        sqlDoc = sqlDoc & "        C96CUNM, C96SEXX, C96AGEE, C96GSMN " & vbNewLine
        sqlDoc = sqlDoc & "  From GC96MP " & vbNewLine
        '-- 작업일자로 조회
        sqlDoc = sqlDoc & " Where C96KSDT = '" & Format(dtpStopDt.Value, "yyyymmdd") & "' " & vbNewLine
        sqlDoc = sqlDoc & "   And C96DEVC = '" & INS_CODE & "' " & vbNewLine
        sqlDoc = sqlDoc & "   And C96TRYN1 <> 'Y' " & vbNewLine
        sqlDoc = sqlDoc & " Order By 1,2,3,4,5"
        
        Set AdoRs_ORACLE = New ADODB.Recordset
        
        AdoRs_ORACLE.CursorLocation = adUseClient
        AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE
        
        If AdoRs_ORACLE.RecordCount = 0 Then
            Set f_subSet_WorkList = Nothing
            RecordChk = False
            Set AdoRs_SQL = Nothing
            Exit Function
        Else
            Set f_subSet_WorkList = AdoRs_ORACLE
            RecordChk = True
        End If
    
        Set AdoRs_ORACLE = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    
    Call ErrMsgProc(CallForm)
    
End Function

'-- 바코드번호로 환자정보를 가져온다.
Private Function f_subGet_PatInfo(ByVal strBarCd As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_PatInfo() As ADODB.Recordset"
    
        Set AdoRs_SQL = New ADODB.Recordset
        
        If cboChk.ListIndex = 0 Then
'            sqlDoc = ""
'            sqlDoc = "         Select a.PRSNVSDT as 검진일자,Rtrim(a.PRSNCODE) + lTrim(a.PRSNSUBC) as 처방코드 , b.ABHJNAME as 수진자명 ,b.ABHJMRNO as 챠트번호 ,b.ABHJSCBT + b.ABHJSCID as 주민등록번호 , a.PRSNRSLT as 검사결과, b.ABHJPSEX as 성별 "
'            sqlDoc = sqlDoc & "  From MediEHE..PRSNUMBM a, MediEHE..ABHJMSTM b"
'            sqlDoc = sqlDoc & " Where a.PRSNVSDT between  '" & strDate & "' and '" & strDate1 & "'"
'            sqlDoc = sqlDoc & "   And a.PRSNRSLT = '' "
'            sqlDoc = sqlDoc & "   And a.PRSNCODE in('21AC','21AN','21AD','21AE','21AF','21AK','21AF1')"
'            sqlDoc = sqlDoc & "   And a.PRSNSUBC In('','001','002','003','004','005')"
'            sqlDoc = sqlDoc & "   And a.PRSNMRNO = b.ABHJMRNO"
'            sqlDoc = sqlDoc & " Order By a.PRSNMRNO"

        Else
'             sqlDoc = "         Select a.*, b.수진자명,b.챠트번호,b.주민등록번호,  b.주민등록번호 as 성별 from TB_검사항목 a, TB_인적사항 b"
'            sqlDoc = sqlDoc & " Where a.진료년+a.진료월+a.진료일 between '" & strDate & "' and '" & strDate1 & "'"
'            sqlDoc = sqlDoc & "   And a.진료지원상태 < 5"
'            sqlDoc = sqlDoc & "   And a.진료지원상태 <> 5"
'            sqlDoc = sqlDoc & "   and 처방코드 in('C4802','C4812','C2243') "
'            sqlDoc = sqlDoc & "   and 서브코드 in('','001','002','003','004','005','006','007') "
'            sqlDoc = sqlDoc & "   And a.챠트번호 = b.챠트번호"
'            sqlDoc = sqlDoc & " Order By a.챠트번호"

        End If
        
        Set AdoRs_SQL = New ADODB.Recordset
        
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subGet_PatInfo = Nothing
            RecordChk = False
            Set AdoRs_SQL = Nothing
            Exit Function
        Else
            Set f_subGet_PatInfo = AdoRs_SQL
            RecordChk = True
        End If
    
        Set AdoRs_SQL = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    
    Call ErrMsgProc(CallForm)
    
End Function

Private Function f_subSet_WorkList_Barcode(ByVal strORDT As String, Optional ByVal strPAID As String, Optional ByVal strSENO As String)
''    Dim sqlRet      As Integer
''    Dim sqlDoc      As String
''    Dim stryy, strmm, strdd, strDate  As String
''
''On Error GoTo ErrorTrap
''    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
''
''
''        Set AdoRs_SQL = New ADODB.Recordset
''
''                            '접수일자,접수번호,항목코드,부속코드,대표장비코드,'
''                            '바코드ID,항목검체,작업일자(검사일자),작업번호,결과유형,
''                            '단위,항목명,검체명,검사부서코드,수진자명,
''                            '거래처명 , 성별, 연령, 검사담당자
''                 sqlDoc = " Select C96JSDT, C96JSNO, C96GSCD, C96ATCD, C96DEVC,"
''        sqlDoc = sqlDoc & "        C96BARC, C96KCCD, C96KSDT, C96GSSQ, C96TYPE,"
''        sqlDoc = sqlDoc & "        C96UNTA, C96GSNM, C96KCNM, C96DECD, C96PTNM,"
''        sqlDoc = sqlDoc & "        C96CUNM, C96SEXX, C96AGEE, C96GSMN "
''        sqlDoc = sqlDoc & "  From GC96MP "
''        sqlDoc = sqlDoc & " Where C96DEVC = '" & INS_CODE & "' "
''        sqlDoc = sqlDoc & "   And C96BARC = '" & strBarno & "' "
''        sqlDoc = sqlDoc & "  Order by C96GSCD "
''
''        Set AdoRs_SQL = New ADODB.Recordset
''        AdoRs_SQL.CursorLocation = adUseClient
''        AdoRs_SQL.Open sqlDoc, AdoCn_ORACLE
''
''        If AdoRs_SQL.RecordCount = 0 Then
''            Set f_subSet_WorkList_Barcode = Nothing
''            RecordChk = False
''            Set AdoRs_SQL = Nothing
''            Exit Function
''        Else
''            Set f_subSet_WorkList_Barcode = AdoRs_SQL
''            RecordChk = True
''        End If
''
''        Set AdoRs_SQL = Nothing
''
''Exit Function
''
''ErrorTrap:
''    Set AdoRs_SQL = Nothing
''
''    Call ErrMsgProc(CallForm)

    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
                       
                 sqlDoc = "SELECT DiSTINCT b.SCP42IDNOA, a.SCP41NAME, a.SCP41JDATE, b.SCP42SUGACD "
        sqlDoc = sqlDoc & vbCrLf & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b "
        sqlDoc = sqlDoc & vbCrLf & " WHERE a.SCP41PCODE = b.SCP42PCODE"
        sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41JDATE = b.SCP42JDATE"
        sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SID   = b.SCP42SID"
        sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SPMNO2 = '" & strORDT & "'"
        sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42SUGACD in (" & strGumCd & "," & strJinCd & ")"
        
'        sqlDoc = sqlDoc & vbCrLf & "   AND (b.SCP42RSTCD <> 'N' OR b.SCP42RSTCD IS null)"
'        sqlDoc = sqlDoc & vbCrLf & "   AND (b.SCP42RSTCD <> 'N' OR b.SCP42RSTCD IS null OR b.SCP42PROFLG  <> 'M')"
        
        '-- 2012.04.13 수정
'        sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SNDYN = 'N' "
        sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RESULT IS NULL "
        
'        sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RSTCD <> 'N'"
        
'        sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SNDYN  = 'N'"
   
        '-- 2012.04.03 추가
        'sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SNDYN  <> 'N' " '--고정값:         'N'"
       ' sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41RSTYN  <> 'Y' " '--고정값:         'Y'"
        'sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RSTCD  = '' " '-- 결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'"
        'sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42RESULT = ''   " '-- 결과값"
        
        
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_ORACLE
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subSet_WorkList_Barcode = Nothing
            RecordChk = False
            Set AdoRs_SQL = Nothing
            Exit Function
        Else
            Set f_subSet_WorkList_Barcode = AdoRs_SQL
            RecordChk = True
        End If
    
        Set AdoRs_SQL = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    
    Call ErrMsgProc(CallForm)

Exit Function
    
End Function
'Private Function f_subSet_WorkList_Barcode(ByVal strBarno As String)
'    Dim sqlRet      As Integer
'    Dim sqlDoc      As String
'
'On Error GoTo ErrorTrap
'    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
'
'
'        Set AdoRs_SQL = New ADODB.Recordset
'
'        If Len(strBarno) > 8 Then
'            sqlDoc = " SELECT a.per_gumjin_date, a.per_gum_num, a.edpscode, a.result, a.send_date, a.per_name " & _
'                    " FROM mdck..gumjin_interface a, mdck..bag_interfacecode b " & _
'                    " WHERE substring(a.per_gumjin_date,3,8) = '" & Mid(strBarno, 1, 6) & "'" & _
'                    " AND a.per_gum_num = '" & Val(Mid(strBarno, 7)) & "' " & _
'                    " AND a.result = '' " & _
'                    " AND substring(b.kind,1,1) = 'C' " & _
'                    " AND a.edpscode=b.meditem " & _
'                    " ORDER BY a.per_gumjin_date, a.per_gum_num "
'        Else
'            sqlDoc = " SELECT a.EnterDate, b.Status, b.waitseqno, b.MAP2SEQNO, b.DispDesc, b.RVALUEKIND, b.NORMLOW, b.NORMHIGH, b.NORMALVALUE, b.RVALUEKIND , " & _
'                    " a.ChartNo, b.GumsaKind, c.sujinname, b.status " & _
'                    " FROM medicom..WaitPrsnp a, medicom..jun370_resulttb b, medicom..pewprsnp c, medicom..BAGMAP2PREF d " & _
'                    " WHERE a.Chartno = '" & strBarno & "' " & _
'                    " AND a.WaitSeqNo = b.WaitSeqNo " & _
'                    " AND a.status = '1' " & _
'                    " AND d.labno = 4 " & _
'                    " AND b.jun370no = d.map2seqno " & _
'                    " AND b.status = '0' " & _
'                    " AND a.chartno = c.chartno " & _
'                    " ORDER BY a.chartno "
'        End If
'
'        Set AdoRs_SQL = New ADODB.Recordset
'        AdoRs_SQL.CursorLocation = adUseClient
'        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
'
'        If AdoRs_SQL.RecordCount = 0 Then
'            Set f_subSet_WorkList_Barcode = Nothing
'            RecordChk = False
'            Set AdoRs_SQL = Nothing
'            Exit Function
'        Else
'            Set f_subSet_WorkList_Barcode = AdoRs_SQL
'            RecordChk = True
'        End If
'
'        Set AdoRs_SQL = Nothing
'
'Exit Function
'
'ErrorTrap:
'    Set AdoRs_SQL = Nothing
'
'    Call ErrMsgProc(CallForm)
'
'
'End Function

Private Function f_subSet_SearchList(ByVal strBarcode As String)
    Dim sqlRet      As Integer
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
    
    Set AdoRs_SQL = New ADODB.Recordset
'    gSql = "select IN_CODE from EXAM_TOC Where RE_RCID = '" & strBarcode & "'"
    
    gSql = "select a.IN_CODE,a.RE_RCID,b.JU_NAME,b.JU_PERID from EXAM_TOC A,JUMN_TMA B,RECE_TJU C" _
            & " Where a.RE_RCID = '" & strBarcode & "'" _
            & " And b.HE_UNID = 'HC-46101'" _
            & " And a.RE_RCID = c.RE_RCID" _
            & " And b.JU_PERID = c.JU_PERID " _
            & " And a.EX_INST = '2'" _
            & " And a.IN_CODE like 'SE%'"

    AdoRs_SQL.Open gSql, AdoCn_ORACLE, adOpenStatic, adLockReadOnly
    
    If AdoRs_SQL.RecordCount = 0 Then
        Set f_subSet_SearchList = Nothing
    Else
        Set f_subSet_SearchList = AdoRs_SQL
    End If

    Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intCnt  As Integer, IntRow  As Integer
    
    Dim intPos1 As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 10
    intCol2 = 1
    IntRow = 1
    With spdWorklist
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by OUT_SEQ, TESTCD_EQP"
'             " order by TESTCD_EQP, TESTCD"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then
        adoRS.MoveFirst
        ReDim fChannel(adoRS.RecordCount)
        strJinCd = ""
        strGumCd = ""
    End If
    
    Do While Not adoRS.EOF
        If Trim(adoRS.Fields("TESTCD")) <> "" Then
            intPos1 = InStr(Trim(adoRS.Fields("TESTCD")), ",")
            If intPos1 = 0 Then
                strGumCd = strGumCd & "'" & Trim(adoRS.Fields("TESTCD")) & "',"
'                strJinCd = strJinCd & "'" & Trim(AdoRs.Fields("TESTCD")) & "',"
            Else
                strGumCd = strGumCd & "'" & Mid(Trim(adoRS.Fields("TESTCD")), 1, intPos1 - 1) & "',"
'                strJinCd = strJinCd & "" & Mid(Trim(AdoRs.Fields("TESTCD")), intPos1 + 1) & ","
                strJinCd = strJinCd & "'" & Mid(Trim(adoRS.Fields("TESTCD")), intPos1 + 1) & "',"
            End If
        End If
        
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFL") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFH") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REMARK") & "")
            itemX.tag = Trim(adoRS.Fields("TEST_EQP") & "")
            itemX.text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdWorklist
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
            .Col = intCol:  .ColHidden = True
        End With
        
        With spdResult1
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 12.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdRstview
            If IntRow > .maxrows Then
                IntRow = 1
                intCol2 = intCol2 + 2
            End If
            
            .SetText intCol2, IntRow, Trim$(adoRS("TESTNM") & "")
            IntRow = IntRow + 1
            
        End With
        
        With spdResult2
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 12.5
            End If
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        fChannel(intCol - colNWNO) = adoRS.Fields("TEST_EQP")
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(adoRS.Fields("TEST_EQP"))
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(adoRS.Fields("TESTCD"))
        intPos = InStr(strTmp, ",")
        Do While intPos > 0
            f_strOrdList = f_strOrdList + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
            
            f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
            f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = Mid$(strTmp, 1, intPos - 1)
            
            strTmp = Mid$(strTmp, intPos + 1)
            
            intPos = InStr(strTmp, ",")
        Loop
        f_strOrdList = f_strOrdList + "'" + strTmp + "',"
        f_typCode(intCnt).intCnt = f_typCode(intCnt).intCnt + 1
        f_typCode(intCnt).strTestcd(f_typCode(intCnt).intCnt) = strTmp
        
        intCol = intCol + 1
        
        adoRS.MoveNext
    Loop
    Set adoRS = Nothing
    
    If Trim(strGumCd) <> "" Then strGumCd = Mid(strGumCd, 1, Len(strGumCd) - 1)
    If Trim(strJinCd) <> "" Then strJinCd = Mid(strJinCd, 1, Len(strJinCd) - 1)
    
    With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol:  .ColHidden = True
        .ColWidth(intCol) = 12
    End With

Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Function f_funGet_CODE(ByVal strOrdcd As String) As String

    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funGet_CODE = ""
    
    For intIdx1 = 1 To UBound(f_typCode)
        For intIdx2 = 1 To f_typCode(intIdx1).intCnt
            If Trim$(strOrdcd) = Trim$(f_typCode(intIdx1).strTestcd(intIdx2)) Then
                f_funGet_CODE = f_typCode(intIdx1).strEqpCd
                Exit Function
            End If
        Next
    Next
    
End Function

Private Function f_subSet_ComList() As String
    
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap

    CallForm = "clsCommon - Public Function f_subSet_ComList() As ADODB.Recordset"
    
   
        Set AdoRs_SQL = New ADODB.Recordset
        
        sqlDoc = "         SELECT B.COM_CODE, B.COM_NAME " & vbCr
        sqlDoc = sqlDoc & "  FROM MDCK..GUMJIN_INTERFACE A, MDCK..TB_COMPANY B, MDCK..BAG_INTERFACECODE C " & vbCr
        sqlDoc = sqlDoc & " WHERE A.Per_com_Code = B.COM_CODE " & vbCr
        sqlDoc = sqlDoc & "   AND A.per_gumjin_date BETWEEN '" & Format(dtpStartDt.Value, "yyyymmdd") & "' AND '" & Format(dtpStopDt.Value, "yyyymmdd") & "'" & vbCr
        sqlDoc = sqlDoc & "  AND SUBSTRING(C.KIND, 1, 1) = 'C' " & vbCr
        sqlDoc = sqlDoc & "   AND A.EDPSCODE = C.MEDITEM " & vbCr
        sqlDoc = sqlDoc & " GROUP BY B.COM_CODE, B.COM_NAME " & vbCr
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount > 0 Then
            AdoRs_SQL.MoveFirst
            cboComNm.Clear
            cboComNm.AddItem "전체"
            Do Until AdoRs_SQL.EOF
                cboComNm.AddItem AdoRs_SQL.Fields("COM_NAME") & ""
                AdoRs_SQL.MoveNext
            Loop
            cboComNm.ListIndex = 0
        End If
        
        AdoRs_SQL.Close:  Set AdoRs_SQL = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub cboChk_Click()
    If Trim(cboChk.text) = "검진" Then
'        cboComNm.Visible = True
'        mskOrdtime.Visible = False
'        Label10.Visible = False
'        Call f_subSet_ComList
    Else
        cboComNm.Visible = False
'        mskOrdtime.Visible = True
'        Label10.Visible = True
        cboComNm.Clear
    End If
End Sub

Private Sub cboComNm_DropDown()
        Call f_subSet_ComList
End Sub

Private Sub cmdAppend_Click(Index As Integer)
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String

    Dim varTmp  As Variant, strErrMsg   As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestcd    As String, strTestRst   As String
    Dim strTestnm       As String
    Dim strRef          As String
    Dim strUnit         As String
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String

    Dim IntRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As vaSpread
    Dim sqlRet  As Integer
    Dim flgSave As Boolean
    Dim SaveGbn As Integer
    
    Dim strDate As String
    Dim strTime As String
    
    Dim strBarno As String
    Dim strSPnm As String
    Dim strSPid As String
    Dim strChartNo As String
    Dim strEqpCd As String
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD As String
    Dim strRefVal As String
    Dim pName As String
    Dim pNo As String
    Dim j As Integer
    Dim s As Integer
    Dim varTstCd As Variant
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"

On Error GoTo ErrorRoutine

    Me.MousePointer = 11

    If Index = 0 Then
        Set objSpd = spdResult1
    Else
        Set objSpd = spdResult2
    End If

    With objSpd
        For IntRow = 1 To .maxrows
            .GetText colORDT, IntRow, varTmp:    strORDT = Trim$(varTmp)
            .GetText colORQN, IntRow, varTmp:    strORQN = Trim$(varTmp)
            .GetText colPANM, IntRow, varTmp:    strPANM = Trim$(varTmp)
            .GetText colPAID, IntRow, varTmp:    strPAID = Trim$(varTmp): strBarno = strPAID
            .GetText colOIFL, IntRow, varTmp:    strOIFL = Trim$(varTmp)
            .GetText colSENO, IntRow, varTmp:    strSENO = Trim$(varTmp)
            .GetText colSEXS, IntRow, varTmp:    strSEXS = Trim$(varTmp)
            .GetText colAGES, IntRow, varTmp:    strAGES = Trim$(varTmp)
            .GetText colNWNO, IntRow, varTmp:    strNWNO = Trim$(varTmp)

            .GetText colBANO, IntRow, varTmp

            If strPAID = "" Then Exit For

            intCnt = 0: Erase strOrdcd ': Erase strRstval
            
            If Trim$(varTmp) = "1" And Len(strBarno) = 10 Then
                For intCol = 10 To .MaxCols
                    strDate = Format$(Now, "YYYYMMDD"):    strTime = Format$(Now, "HHMMSS")
                    .GetText intCol, IntRow, varTmp
                        If Trim$(varTmp) <> "" Then
                            .GetText intCol, 0, varTmp
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                .GetText intCol, IntRow, varTmp: strRstval = varTmp
                                strTestcd = itemX.ListSubItems(1)
                                For j = 0 To 100
                                    'intPos = InStr(strTestcd, ",")
                                    varTstCd = Split(strTestcd, ",")
                                    For s = 0 To UBound(varTstCd)
                                        If varTstCd(s) = strSNOCode(j) Then
                                            strTestcd = strSNOCode(j)
                                            Exit For
                                        End If
                                    Next
                                Next
                                
                                'intPos = InStr(strTestcd, ",")
                                strEqpCd = strTestcd
                                                    
                                If strEqpCd <> "" Then
                                    'If intPos > 0 Then
                                    '바코드번호로 오더찾아오기
                                    Set mAdoRs = f_subSet_WorkList_Barcode(strBarno)
                                    
                                    If RecordChk = True Then
                                        Do Until mAdoRs.EOF
'                                            '-- 로컬저장
'                                            sqlDoc = "Update INTERFACE003" & _
'                                                     "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
'                                                     " where SPCNO   = '" & strBarno & "'" & _
'                                                     "   and EQPNUM  = '" & itemX.tag & "'" & _
'                                                     "   and TRANSDT = '" & strDate & "'" & _
'                                                     "   and TRANSTM = '" & strTime & "'"
'                                            AdoCn_Jet.Execute sqlDoc
'
'                                            sqlDoc = "insert into INTERFACE003(" & _
'                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
'                                                     "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
'                                                     "            '" & strDate & "', '" & strTime & "'," & _
'                                                     "            '" & strRstval & "', '" & strRefVal & "'," & _
'                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
'                                            AdoCn_Jet.Execute sqlDoc
                                            
                                            'If InStr(strEqpCd, Trim(mAdoRs.Fields("SCP42SUGACD").Value)) > 0 Then
                                            
                                            '   3-1. 검사정보 MASTER
                                                     sqlDoc = "UPDATE JAIN_SCP.SCPRST41 SET "
                                            sqlDoc = sqlDoc & "       SCP41TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                                            sqlDoc = sqlDoc & "       SCP41SNDYN  = 'N',"                               '고정값 : 'N'
                                            sqlDoc = sqlDoc & "       SCP41RSTYN  = 'Y',"                               '고정값 : 'Y'
                                            sqlDoc = sqlDoc & "       SCP41TSTUID = '" & CurrUser.CuUserID & "'"        '검사자사번
                                            sqlDoc = sqlDoc & " WHERE SCP41SPMNO2 = '" & strBarno & "'"                 '바코드번호
                                            
                                            AdoCn_ORACLE.Execute sqlDoc
                                            
                                            
                                            '   3-2. 검사정보 DETAIL
                                                     sqlDoc = "UPDATE JAIN_SCP.SCPRST42 SET "
                                            sqlDoc = sqlDoc & "       SCP42TSTDAT = '" & Format(Now, "YYYYMMDD") & "'," '결과일자 => YYYYMMDD"
                                            sqlDoc = sqlDoc & "       SCP42RSTCD  = 'N',"                               '결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'
                                            sqlDoc = sqlDoc & "       SCP42RESULT = '" & strRstval & "'"                '결과값
                                            sqlDoc = sqlDoc & " WHERE SCP42SPMNO2 = '" & strBarno & "'"                 '바코드번호
                                            sqlDoc = sqlDoc & "   AND SCP42SUGACD = '" & strEqpCd & "'"              '수가코드
        
                                            AdoCn_ORACLE.Execute sqlDoc
                                            
                                            lblStatus.Caption = "저장 성공!!"
                                                                                                                            
                                            .Row = IntRow: .Col = colBANO: .Value = 0
                                                           .Col = colORDT: .BackColor = HNC_Cyan
                                                           .Col = colORQN: .BackColor = HNC_Cyan
                                                           .Col = colPANM: .BackColor = HNC_Cyan
                                                           .Col = colPAID: .BackColor = HNC_Cyan
                                                           .Col = colOIFL: .BackColor = HNC_Cyan
                                                           .Col = colSENO: .BackColor = HNC_Cyan
                                                           '.Col = colSEXS: .BackColor = HNC_Cyan
                                                           '.Col = colAGES: .BackColor = HNC_Cyan
                                                           '.Col = colNWNO: .BackColor = HNC_Cyan
                                                                                                                        
                                                                        
                                            'End If
                                        mAdoRs.MoveNext
                                        Loop
                                    End If
                                                                            
                                End If
                            Set itemX = Nothing
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    Me.MousePointer = 0
    
    If lblStatus.Caption = "저장 성공!!" Then
        'MsgBox "▒ EMR SERVER에 결과를 Upload 완료되었습니다. ▒      " & vbCrLf & vbCrLf & "     LIS 결과조회 화면에서 결과를 확인 하십시요..  ", vbInformation, App.Title
    End If
    
    Exit Sub
ErrorRoutine:

    Set AdoRs_SQL = Nothing

    Set itemX = Nothing

    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)
    
End Sub


Private Sub cmdEot_Click()
    Call COM_OUTPUT(EOT)
End Sub

Private Sub cmdExcel_Click()
'Dim sRow As Integer, sCol As Integer, sCnt As Integer
'Dim sSave As Boolean
'Dim fName As String
'
'    If chkExcel.Value = 1 Then
'        With CommonDialog1
'             .FileName = App.Path & "\" & fName & ".xls"
'             .DialogTitle = "Save As New Excel Spread"
'             .FileName = REG_INSNAME & "  " & Format(mskRstDate, "####-##-##") & " 검사현황대장"
'             .Filter = "New Excel file(*.xls)"
'             .ShowSave
'            sSave = spdResult2.ExportToExcel(.FileName, Format(mskRstDate, "####-##-##") & " TBA20FR", "\log.txt")
'        End With
'    Else
'        Call gsp_SetSpdTExcelExport(spdResult2, True)
'    End If

    Dim strTmp As String
    Dim lngRows As Long
    
    If spdResult2.DataRowCnt = 0 And spdResult2.DataRowCnt = 0 Then Exit Sub
    
    With spdResult2
        .Row = 0: .Row2 = .maxrows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .maxrows
    End With
 
    With tblexcel
        .maxrows = spdResult2.maxrows + 1
        .MaxCols = spdResult2.MaxCols
        .Row = 1: .Row2 = .maxrows
        .Col = 1: .Col2 = spdResult2.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.filter = "ExCelFile(*.XLS)|*.XLS"
    CommonDialog1.FileName = REG_INSNAME & "  " & Format(mskRstDate, "####-##-##") & " 검사현황대장"
    CommonDialog1.ShowSave

    tblexcel.SaveTabFile (CommonDialog1.FileName)

End Sub

Private Sub cmdOrder_Click()

    comEQP.Output = ENQ
    cInterface.state = "Q"
    cInterface.Snd_Phase = 0
    tmrOrder.Enabled = False
    
End Sub

Private Sub cmdPosNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                .Row = sCnt
                .Col = 7:       .text = Trim(sAdd + Val(sNo))
                If Trim(sAdd + Val(sNo)) = 14 Then sNo = 0
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If
End Sub

Private Sub cmdPrint_Click()
Dim objclsCommon As New clsCommon

Dim Tmp_Testnm As String
Dim Row_cnt As Integer, Col_cnt As Integer, TmpPrintline As Integer
Dim vTmp As Variant
Dim stragesex As String

Const TmpLine = "────────────────────────────────────────────────────────────"

    If spdResult1.maxrows >= 1 Then
        With objclsCommon
            .PrintText 15, 3, Format(Date, "yyyy/mm/dd") & "  WorkList Report..( " & App.EXEName & " )", "Arial", 12
            
            .PrintText 0.5, 5, TmpLine
            .PrintText 0.5, 6, "순", , 9
            .PrintText 2, 6, "접수일자", , 9
            .PrintText 7, 6, "환자성명", , 9
            .PrintText 12, 6, "접수번호", , 9
            .PrintText 17, 6, "거래처명", , 9
            .PrintText 27, 6, "장비검사종목", , 9
            .PrintText 0.5, 7, TmpLine
            
            TmpPrintline = 8
        
        For Row_cnt = 1 To spdResult1.maxrows
            spdResult1.Row = Row_cnt
            
            If (Row_cnt Mod 34) <> 0 Then
                                    .PrintText 0.5, TmpPrintline, Row_cnt, , 9                          ' 순
                spdResult1.Col = 3: .PrintText 2, TmpPrintline, Mid(spdResult1.text, 3), , 9           ' 처방일자
                spdResult1.Col = 10: .PrintText 7, TmpPrintline, Trim(spdResult1.text), 9              ' 검체번호
                spdResult1.Col = 4: .PrintText 12, TmpPrintline, Trim(spdResult1.text), , 9             ' 이    름
                spdResult1.Col = 9: .PrintText 17, TmpPrintline, Trim(spdResult1.text), , 9             ' 병원명
                
                
                For Col_cnt = 11 To spdResult1.MaxCols
            
                    spdResult1.Row = Row_cnt:            spdResult1.Col = Col_cnt
                    
                    If spdResult1.BackColor = &HC6FEFF Then
                        spdResult1.GetText Col_cnt, 0, vTmp
                        Tmp_Testnm = Tmp_Testnm & ", " & vTmp
                    End If
                    
                Next Col_cnt
                
                spdResult1.Col = 5: .PrintText 27, TmpPrintline, Mid(Trim(Tmp_Testnm), 2), , 7.5
                
                TmpPrintline = TmpPrintline + 2
                Tmp_Testnm = ""
            Else
            
                '-------------------------------------------------------
            
                                    .PrintText 0.5, TmpPrintline, Row_cnt, , 9                          ' 순
                spdResult1.Col = 3: .PrintText 2, TmpPrintline, Mid(spdResult1.text, 3), , 9                   ' 처방일자
                spdResult1.Col = 10: .PrintText 7, TmpPrintline, Trim(spdResult1.text), 9              ' 검체번호
                spdResult1.Col = 4: .PrintText 12, TmpPrintline, Trim(spdResult1.text), , 9             ' 이    름
                spdResult1.Col = 9: .PrintText 17, TmpPrintline, Trim(spdResult1.text), , 9             ' 병원명
                
                
                For Col_cnt = 11 To spdResult1.MaxCols
            
                    spdResult1.Row = Row_cnt:            spdResult1.Col = Col_cnt
                    
'                    If Trim(spdResult1.text) <> "" Then
                    If spdResult1.BackColor = &HC6FEFF Then
                        spdResult1.GetText Col_cnt, 0, vTmp
                        Tmp_Testnm = Tmp_Testnm & ", " & vTmp
                    End If
                    
                Next Col_cnt
                
                spdResult1.Col = 5: .PrintText 27, TmpPrintline, Mid(Trim(Tmp_Testnm), 2), , 7.5
                
                TmpPrintline = TmpPrintline + 2
                Tmp_Testnm = ""
                
                '-------------------------------------------------------
            
                    .PrintText 0.5, TmpPrintline, TmpLine
                    .PrintText 1, TmpPrintline + 1, "── Next Report ──", , 9, True
                    Printer.NewPage
                    
'                    .PrintText 0.5, 5, TmpLine
'                    .PrintText 0.5, 6, "순", , 9
'                    .PrintText 2, 6, "접수번호", , 9
'                    .PrintText 7, 6, "환자성명", , 9
'                    .PrintText 12, 6, "처방일자", , 9
'                    .PrintText 17, 6, "거래처명", , 9
'                    .PrintText 27, 6, "장비검사종목", , 9
'                    .PrintText 0.5, 7, TmpLine
                                        
                                        
                    .PrintText 0.5, 5, TmpLine
                    .PrintText 0.5, 6, "순", , 9
                    .PrintText 2, 6, "접수일자", , 9
                    .PrintText 7, 6, "환자성명", , 9
                    .PrintText 12, 6, "접수번호", , 9
                    .PrintText 17, 6, "거래처명", , 9
                    .PrintText 27, 6, "장비검사종목", , 9
                    .PrintText 0.5, 7, TmpLine
                                        
                    TmpPrintline = 9
            End If
        
        Next Row_cnt
        .PrintText 0.5, TmpPrintline, TmpLine
        .PrintText 1, TmpPrintline + 1, "── End of Report ──", , 9, True
        
        End With
        Printer.NewPage
        Printer.EndDoc
        
        MsgBox Format(Date, "yyyy/mm/dd") & "일자의 " & App.EXEName & "의 장비 검사 WorkList가 Print되었습니다..       " & vbCrLf & vbCrLf & "다음 작업을 진행하십시요..", vbInformation + vbOKOnly, App.Title
    Else
        MsgBox Format(Date, "yyyy/mm/dd") & "일자의 " & App.EXEName & "의 장비 검사 WorkList가  Load 되어 있지 않습니다..       " & vbCrLf & vbCrLf & "자료를 확인 하십시요..", vbInformation + vbOKOnly, App.Title
    End If
    
    '
    ' 마지막 저장
    '
    spdResult1.SaveTabFile App.Path & "\" & REG_INSNAME & "_Request.txt"
    

End Sub

Private Sub cmdRackNo_Click()
'    Dim sNo As String, sCnt As Integer, sAdd As Integer
'    Dim aROW    As Integer, aCOL   As Integer
'    Dim varChk  As Variant, varBar As Variant, varNum As Variant
'    Dim iRow    As Integer, iCnt   As Integer
'    Dim strRack_tmp As String
'
'
'AgainInput:
'    sNo = InputBox("시작 번호를 입력하세요 !")
'    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
'        If Not IsNumeric(sNo) Then
'            MsgBox "숫자만 입력하세요.!", vbCritical
'            GoTo AgainInput
'        End If
'
'        With spdResult1
'            iCnt = 1
'            .GetText 1, 1, varChk
'            .GetText 2, 1, varBar
'            varNum = sNo
'            If Trim(varChk) = "1" And Trim(varBar) <> "" Then
'                For iRow = 1 To .maxrows
'                    .SetText 6, iRow, varNum
'                    .SetText 7, iRow, ((iCnt Mod 101) + 1) - 1
'                    iCnt = iCnt + 1
'                    If (iCnt Mod 101) = 1 Then varNum = varNum + 1
'                Next
'            End If
'        End With
'    End If

Dim sNo As String, sCnt As Integer, sAdd As Integer
Dim fNum1 As Integer, fNum2 As Integer
Dim intRow1 As Integer

AgainInput:
    fNum1 = 1: fNum2 = 0
    sNo = InputBox("시작 렉번호를 입력하세요!")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
'        sNo = UCase(sNo)
        
'        If Asc(sNo) < 65 Or Asc(sNo) > 70 Then
'            MsgBox "a~f까지의 문자만 입력하세요.!", vbCritical
'            GoTo AgainInput
'        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                intRow1 = intRow1 + 1
                .Row = sCnt
                .Col = 1
                If .Value >= 1 Then
                    .Col = 6
                    If intRow1 = (14 * fNum1) + 1 Then
                        fNum1 = fNum1 + 1: fNum2 = 0
                    End If
                    fNum2 = fNum2 + 1
                    .text = Chr(fNum1 + Asc(sNo) - 1)
                    
                    .Col = 7
                    .text = fNum2
                End If
            Next sCnt
        End With
    End If

End Sub

Private Sub cmdRequist_Click(Index As Integer)
    Dim ret As Integer
    
    ret = spdResult1.LoadFromFile(App.Path & "\" & REG_INSNAME & "_Request.txt")

End Sub

Private Sub cmdSearch_Click()
    Dim strDoc      As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim IntRow      As Integer
    Dim pGrid_Point As Integer
    Dim intCnt      As Integer
    Dim strBarno As String
    Dim itemX As ListItem
    Dim strEqpCd As String
    Dim strBartmpNo As String
    Dim blt As Boolean
    Dim intWorkNo  As Integer
    Dim intCol As Integer
    '-- WorkList조회
    Dim strTime As String

'GoTo Rst
'    strDoc = AS400DOWNF(INS_CODE, Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'
'    If strDoc = "0" Then
'        'MsgBox "성공"
'    ElseIf strDoc = "44" Then
'        '-- 장비코드 오류
'        MsgBox "장비코드 설정을 확인하세요", vbOKOnly + vbCritical, Me.Caption
'        Exit Sub
'    ElseIf strDoc = "48" Then
'        '시작일오류
'        MsgBox "작업일자 시작일을 확인하세요", vbOKOnly + vbInformation, Me.Caption
'        Exit Sub
'    ElseIf strDoc = "52" Then
'        '종료일오류
'        MsgBox "작업일자 종료일을 확인하세요", vbOKOnly + vbInformation, Me.Caption
'        Exit Sub
'    End If

'Rst:

    With spdResult1
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    blt = True


On Error GoTo ErrorTrap

    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
        
    Set AdoRs_ORACLE = New ADODB.Recordset
       
    strTime = mskOrdtime.text
    Set mAdoRs = f_subSet_WorkList(Format(dtpStopDt.Value, "yyyymmdd"), Format(dtpStartDt.Value, "yyyymmdd"), Trim(txtToNo.text))
    
    If RecordChk = False Then
        MsgBox Format(dtpStopDt.Value, "yyyymmdd") & "일의 검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    Else
        strBarno = ""
        mAdoRs.MoveFirst
        
        With spdResult1
            If RecordChk = True Then
                Do Until mAdoRs.EOF
                    strBarno = Trim(mAdoRs("C96BARC")) & ""
                    pGrid_Point = SeqSearch(spdResult1, strBarno, 7)

                    If pGrid_Point = 0 Then
                        pGrid_Point = SeqNullSearch(spdResult1, strBarno, 7)
                        If pGrid_Point = 0 Then
                            spdResult1.maxrows = spdResult1.maxrows + 1: pGrid_Point = spdResult1.maxrows
                        End If
                    End If
                
                    spdResult1.Row = pGrid_Point
                    spdResult1.BackColor = vbGreen '&HC6FEFF '&H80C0FF
                    
                    spdResult1.SetText 1, pGrid_Point, "1"
                    'spdResult1.SetText 2, pGrid_Point, strRackNo & "," & strPos 'pGrid_Point   '-- No
                    spdResult1.SetText 3, pGrid_Point, Format(mAdoRs("C96JSDT"), "####-##-##") '-- 검체번호
                    spdResult1.SetText 4, pGrid_Point, Trim(mAdoRs("C96JSNO")) & "" '-- 접수번호
                    spdResult1.SetText 5, pGrid_Point, Trim(mAdoRs("C96GSCD")) & "" '-- 항목코드
                    spdResult1.SetText 6, pGrid_Point, Trim(mAdoRs("C96ATCD")) & "" '-- 부속코드
                    spdResult1.SetText 7, pGrid_Point, Trim(mAdoRs("C96BARC")) & "" '-- 바코드
                    spdResult1.SetText 8, pGrid_Point, Trim(mAdoRs("C96KCCD")) & "" '-- 항목검체
                    spdResult1.SetText 9, pGrid_Point, Trim(mAdoRs("C96CUNM")) & "" '-- 거래처명
                    spdResult1.SetText 10, pGrid_Point, Trim(mAdoRs("C96PTNM")) & "" '-- 수진자명
                    'spdResult1.SetText 11, pGrid_Point, Trim(mAdoRs("C96SEXX")) & "/" & Trim(mAdoRs("C96AGEE")) '-- 성별/나이
                    
                    .Row = pGrid_Point: .Col = 1: .ForeColor = HNC_Black
                                        .Col = 2: .ForeColor = HNC_Black
                                        .Col = 3: .ForeColor = HNC_Black
                                        .Col = 4: .ForeColor = HNC_Black
                    
'                    If blt = False Then
'                        .Row = pGrid_Point - 1
'                        .Action = ActionDeleteRow
'                        .maxrows = .maxrows - 1
'                    Else
'                        blt = False
'                    End If
                    
    
                    strEqpCd = f_funGet_CODE(Trim(mAdoRs("C96GSCD")) & Trim(mAdoRs("C96ATCD")))
                    
                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                    If Not itemX Is Nothing Then
                        'spdWorklist.SetText 1, pGrid_Point, "0"
                        .Col = itemX.Index + 10
                        .Row = pGrid_Point
                        .BackColor = &HC6FEFF   '&HC6FEFF
                        blt = True
                    End If
'                    intWorkNo = pGrid_Point
                    mAdoRs.MoveNext
                Loop
            End If
            
            Set mAdoRs = Nothing
                            
    
            If blt = False Then
                .Row = pGrid_Point
                .Action = ActionDeleteRow
                .maxrows = .maxrows - 1
            End If
        End With
    End If
    
    Set AdoRs_SQL = Nothing

Exit Sub

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub cmdACK_Click()
'
'    Call COM_OUTPUT(charCOM_Convert(COM_ACK))
Call COM_OUTPUT(Chr(1))
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0:     Call cmdRun
        Case 1:     Call cmdStop
        Case 2:     Call cmdClear
        Case 3:     Call cmdExit
        Case Else
    End Select
    
    IntRow = 0
    
End Sub

Private Sub cmdClear()
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    Or_Seq = 1
    txtDump.text = ""
    txtChart.text = ""
    
    With spdWorklist

        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
        .maxrows = 1
        
    End With
    
    With spdResult1
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 13
    End With

    With spdResult2
        .maxrows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    Dim Rowcnt As Integer
    Dim Colcnt As Integer

    With spdRstview
        .maxrows = 40
        For Rowcnt = 1 To 40
            For Colcnt = 2 To 6 Step 2
                .Row = Rowcnt
                .Col = Colcnt
                .BackColor = &HFFFFFF
                .text = ""
'                .Col = Colcnt + 1
'                .BackColor = &HFFFFFF
'                .text = ""
            
            Next Colcnt
        Next Rowcnt
    End With

End Sub

Private Sub cmdExit()
    
    Unload Me

End Sub

Private Sub cmdRun()
    
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    
    If Not comEQP.PortOpen Then comEQP.PortOpen = True
    If comEQP.PortOpen Then
        Call ShowMessage("연결 되었습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    Else
        Call ShowMessage("연결 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
    End If
        
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdStop()
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If comEQP.PortOpen Then comEQP.PortOpen = False
    If comEQP.PortOpen Then
        Call ShowMessage("중지 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    Else
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
    End If
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

'Private Sub cmdAppend_Click(Index As Integer)
'
'    Dim adoRS   As New ADODB.Recordset
'    Dim sqlDoc  As String
'
'    Dim varTmp  As Variant, strErrMsg   As String
'    Dim strSampleno()   As String, strBarno     As String, strTime      As String
'    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
'    Dim strTmp1()       As String, strTmp2()    As String
'    Dim intPos          As String, strTestcd    As String, strTestRst   As String
'    Dim strTestnm       As String
'    Dim strRef          As String
'    Dim strUnit         As String
'    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
'
'    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
'    Dim itemX   As ListItem
'    Dim objSpd  As vaSpread
'    Dim sqlRet  As Integer
'    Dim flgSave As Boolean
'    Dim SaveGbn As Integer
'    Dim strDate As String
'
'    CallForm = "frmComm - Private Sub cmdAppend_Click()"
'
'On Error GoTo ErrorRoutine
'
'    Me.MousePointer = 11
'
'    If Index = 0 Then
'        Set objSpd = spdResult1
'    Else
'        Set objSpd = spdResult2
'    End If
'
'    With objSpd
'        For intRow = 1 To .maxrows
'
''            .GetText 2, intRow, varTmp:         strDate = Trim$(varTmp)
''            .GetText 3, intRow, varTmp:         strBarno = Trim$(varTmp)
''            .GetText .MaxCols, intRow, varTmp:  strTime = Trim$(varTmp)
'
'            .GetText 2, pGrid_Point, varTmp:   strDate = Trim$(varTmp)
'            .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
'            .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
'            .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
'
'
'            .GetText 1, intRow, varTmp
'
'            If strBarno = "" Then Exit For
'
'            intCnt = 0: Erase strOrdcd: Erase strRstval
'            If Trim$(varTmp) = "1" Then
'                For intCol = 6 To .MaxCols
'                    .GetText intCol, intRow, varTmp
'                    If Trim$(varTmp) <> "" Then
'                        .GetText intCol, 0, varTmp
'                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
'                        If Not itemX Is Nothing Then
'                            .GetText intCol, intRow, varTmp
'                            strTestcd = itemX.ListSubItems(1)
'                            intPos = InStr(strTestcd, ",")
'                            If intPos > 0 Then
'                                Do While intPos > 0
'
'                                    blnFlag = False
'                                    Set mAdoRs = f_subSet_TestList(strBarno)
'                                    Do Until mAdoRs.EOF
'                                        If mAdoRs("LCODE") = Mid$(strTestcd, 1, intPos - 1) Then blnFlag = True: Exit Do
'                                        mAdoRs.MoveNext
'                                    Loop
'                                    strTestcd = Mid$(strTestcd, intPos + 1)
'                                    intPos = InStr(strTestcd, ",")
'
'                                    AdoCn_ORACLE.BeginTrans
'
'                                    sqlDoc = "insert into lab_result" & _
'                                             "   (RESULTNO, LCODE, LSEQ, LNAME, LRESULT, UNIT, REFV,LTYPE,RESULT_DATE, REPORTER, PT_SEQ) " & _
'                                             " values('" & mAdoRs("resultno") & "', '" & strTestcd & "'," & _
'                                             "       '1', '" & Trim(itemX.ListSubItems(2)) & "'," & _
'                                             "       '" & Trim(varTmp) & "', '" & Trim(itemX.ListSubItems(9)) & "'," & _
'                                             "       '" & Trim(itemX.ListSubItems(8)) & "','0',sysdate," & _
'                                             "       '70001','" & mAdoRs("pt_seq") & "')"
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    sqlDoc = ""
'                                    sqlDoc = sqlDoc & "update ipd_order_date set req_result2 = '*'"
'                                    sqlDoc = sqlDoc & " where patient_no = '" & mAdoRs("patient_no") & "'"
'                                    sqlDoc = sqlDoc & "   and order_date = to_date('" & strDate & "','yyyy-mm-dd')"
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    sqlDoc = ""
'                                    sqlDoc = sqlDoc & "update lab_order set r_flag='1'"
'                                    sqlDoc = sqlDoc & " where resultno = '" & mAdoRs("resultno") & "'"
'                                    sqlDoc = sqlDoc & "   and patient_no = '" & mAdoRs("patient_no") & "'"
'                                    sqlDoc = sqlDoc & "   and lcode = '" & mAdoRs("lcode") & "'"
'                                    sqlDoc = sqlDoc & "   and pt_seq = '" & mAdoRs("pt_seq") & "'"
'
'                                    AdoCn_ORACLE.Execute sqlDoc
'
'                                    AdoCn_ORACLE.CommitTrans
'
'                                    lblStatus.Caption = "저장 성공!!"
'                                    Set adoRS = Nothing:    mAdoRs.Close
'                                Loop
'                            Else
'                                blnFlag = False
'                                Set mAdoRs = f_subSet_TestList(strBarno)
'                                Do Until mAdoRs.EOF
'                                    If Trim(mAdoRs("LCODE")) = strTestcd Then blnFlag = True: Exit Do
'                                    mAdoRs.MoveNext
'                                Loop
'                                If blnFlag Then
'                                    AdoCn_SQL.BeginTrans
'
'                                    If chkAuto.Value = "1" Then
'                                           If Mid(pName, 1, 2) = "검진" Then
'                                               sqlDoc = "Update MDCK..GUMJIN_INTERFACE" & _
'                                                        "   set RESULT = '" & strRstval & "'," & _
'                                                        "       ACT_RETURN_DATE = '" & strDate & "'" & _
'                                                        " where PER_GUMJIN_DATE = '" & strDate2 & "'" & _
'                                                        "   and PER_GUM_NUM = " & pNo & "" & _
'                                                        "   and EDPSCODE = '" & Mid(itemX.text, 1, 4) & "'"
'                                           Else
'                                               sqlDoc = "Update MEDICOM..jun370_resulttb" _
'                                                       & "   Set Result = '" & strRstval & "', status='1'" _
'                                                       & " Where WaitSeqNo = '" & pNo & "'" _
'                                                       & "   and map2seqno = '" & strEqpCd & "'"
'
'                                           End If
'                                           AdoCn_SQL.Execute sqlDoc
'                                    End If
'
'                                    AdoCn_SQL.CommitTrans
'
'                                    lblStatus.Caption = "저장 성공!!"
'                                    Set adoRS = Nothing:    mAdoRs.Close
'                                End If
'                            End If
'                        End If
'
'                        Set itemX = Nothing
'                    End If
'                Next
'                spdResult1.Row = intRow
'                spdResult1.Col = 2
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 3
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 4
'                spdResult1.BackColor = vbCyan
'                spdResult1.Col = 1: spdResult1.Value = 0
'
'                If strErrMsg = "" Then
'                    sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
'                             " where SPCNO   = '" & strBarno & "'" & _
'                             "   and TRANSDT = '" & mskRstDate.text & "'"
'                    AdoCn_Jet.Execute sqlDoc
'                Else
'                    MsgBox strErrMsg, vbInformation, Me.Caption
'                End If
'            End If
'        Next
'    End With
'    Me.MousePointer = 0
'    MsgBox "작업이 완료되었습니다.", vbInformation, Me.Caption
'
'    Exit Sub
'ErrorRoutine:
'    Set itemX = Nothing
'
'    Me.MousePointer = 0
'    Call ErrMsgProc(CallForm)
'End Sub

Public Function CheckSum_ECi_Tx(ByVal strPrmValue As String)

    Dim i                   As Integer
    Dim intValueLength      As Integer
    Dim intCheck            As Integer
    Dim strCheck            As String
    
    intCheck = 0
    
    intValueLength = LenA(strPrmValue)
    
    For i = 1 To intValueLength
        intCheck = intCheck + Asc(Mid(strPrmValue, i, 1))
    Next
    
    strCheck = Hex(intCheck)
    
    If Len(strCheck) = 1 Then
        CheckSum_ECi_Tx = "0" & strCheck
    Else
        CheckSum_ECi_Tx = Right(strCheck, 2)
    End If

End Function

Public Function LenA(strPrmString As String) As Integer

    Dim i                   As Integer
    Dim intStrLen           As Integer
    Dim intAnsiStrLen       As Integer
    Dim strTemp             As String
    
    intStrLen = Len(strPrmString)
    For i = 1 To intStrLen
        strTemp = Mid(strPrmString, i, 1)
        
        Select Case AscW(strTemp)
        Case 0 To 255
            intAnsiStrLen = intAnsiStrLen + 1
        
        Case Else
            intAnsiStrLen = intAnsiStrLen + 2
        
        End Select
    Next
    
    LenA = intAnsiStrLen

End Function

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdRstQuery_Click()
''
''    Dim AdoRs   As New ADODB.Recordset
''    Dim sqlDoc  As String, intRet   As Integer
''
''    Dim strSpcno    As String
''    Dim intRow      As Integer, intCol  As Integer
''    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
''    Dim strRackNo, strPos As String
''
''    Dim itemX       As ListItem
''
''    intRow = 0
''    With spdResult2
''        .maxrows = 1
''        .Col = 1:   .Col2 = .MaxCols
''        .Row = 1:   .Row2 = .maxrows
''        .BlockMode = True
''        .Action = ActionClearText
''        .BlockMode = False
''    End With
''
''    sqlDoc = "Select SPCNO, TESTCD, EQUIPCD, EQPNUM, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, PATID, PNM, SEX, SPCDT, TMP2 " & _
''             "  From INTERFACE003" & _
''             " Where TRANSDT = '" & Format(dtpRsltDay.Value, "yyyymmdd") & "'" & _
''             "   And EQUIPCD = '" & INS_CODE & "'"
''
''    sqlDoc = sqlDoc & " Order By SPCDT, SPCNO, TRANSTM"
''
''    AdoRs.CursorLocation = adUseClient
''    AdoRs.Open sqlDoc, AdoCn_Jet
''    If AdoRs.RecordCount > 0 Then AdoRs.MoveFirst
''    Do While Not AdoRs.EOF
''        With spdResult2
''        If strSpcno <> Trim$(AdoRs("TRANSDT") & "") + Trim$(AdoRs("SPCNO") & "") Then
''                intRow = intRow + 1
''                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
''                .SetText 1, intRow, "1"
''                .SetText 2, intRow, Format(Trim$(AdoRs("TRANSDT") & ""), "####-##-##")
''                .SetText 3, intRow, Format(Trim$(AdoRs("SPCDT") & ""), "####-##-##")
''                .SetText 4, intRow, Trim$(AdoRs("PATID") & "")
''                .SetText 7, intRow, Trim$(AdoRs("SPCNO") & "")
''                .SetText 8, intRow, Trim$(AdoRs("TMP2") & "")
''                .SetText 10, intRow, Trim$(AdoRs("PNM") & "")
''            End If
''            strSpcno = Trim$(AdoRs("TRANSDT") & "") + Trim$(AdoRs("SPCNO") & "")
''            Set itemX = lvwCuData.FindItem(Trim$(AdoRs("TESTCD") & ""), lvwTag, , lvwWhole)
''            If Not itemX Is Nothing Then
''                intCol = itemX.Index + 10
''                .SetText intCol, intRow, Trim$(AdoRs("RSTVAL")) & ""
''                .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(AdoRs("REFVAL") & "") <> "", vbRed, vbBlack)
''            End If
''        End With
''        AdoRs.MoveNext
''    Loop
''    AdoRs.Close:    Set AdoRs = Nothing
''


    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim IntRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    Dim itemX       As ListItem

    IntRow = 0
    With spdResult2
        .maxrows = 25
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select SPCNO, TESTCD, EQUIPCD, TRANSDT, RSTVAL, REFVAL, TRANSDT, EQPNUM, NAME, PNO" & _
             "  from INTERFACE003" & _
             " where TRANSDT >= '" & mskRstDate.text & "'" & _
             "   and EQUIPCD = '" & INS_CODE & "'"
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = ''"
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = 'Y'"
    End If
    sqlDoc = sqlDoc & " order by SPCNO, TRANSTM"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdResult2
        If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "") Then
                IntRow = IntRow + 1
                If IntRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                .SetText 1, IntRow, "1"
                .SetText 2, IntRow, Trim$(adoRS(3) & "")
                .SetText 3, IntRow, Trim$(adoRS(0) & "")
                .SetText 6, IntRow, Trim$(adoRS(8) & "")
                .SetText 7, IntRow, Trim$(adoRS(9) & "")
                '.SetText .MaxCols, intRow, Trim$(adoRS(6) & "")
            End If
                strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(6) & "")
                Set itemX = lvwCuData.FindItem(Trim$(adoRS(7) & ""), lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    intCol = itemX.Index + 8
                    .SetText intCol, IntRow, Trim$(adoRS(4)) & ""
                    .Col = intCol:  .Row = IntRow:  .ForeColor = IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
                End If
        End With
        adoRS.MoveNext
    Loop
'    spdResult2.MaxCols = spdResult2.MaxCols - 1
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim IntRow  As Integer
    
    If Index = 0 Or Index = 1 Then
        With spdResult2
            For IntRow = 1 To .maxrows
                .GetText 2, IntRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, IntRow, IIf(Index = 0, "1", "")
            Next
        End With
    ElseIf Index = 4 Or Index = 5 Then
        With spdResult1
            For IntRow = 1 To .maxrows
                .GetText 2, IntRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, IntRow, IIf(Index = 4, "1", "")
            Next
        End With
    Else
        With spdWorklist
            For IntRow = 1 To .maxrows
                .GetText 2, IntRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, IntRow, IIf(Index = 0, "1", "")
            Next
        End With
    End If
    
End Sub

Private Sub cmdStartNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                .Row = sCnt
                .Col = 7:       .text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If

End Sub

Private Function f_subSet_TestList(ByVal strDate As String, ByVal strSeq As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_TestList() As ADODB.Recordset"
        Set AdoRs_SQL = New ADODB.Recordset
               
                 sqlDoc = "select EDPSCODE from GUMJIN_INTERFACE"
        sqlDoc = sqlDoc & " where PER_GUMJIN_DATE = '" & strDate & "'"
        sqlDoc = sqlDoc & "   and PER_GUM_NUM = '" & strSeq & "'"
        sqlDoc = sqlDoc & "   and EDPSCODE in ('0208','0226','0225','0227','0207','0206','0205','0221','0222','0223','0224','0209')"
        sqlDoc = sqlDoc & "   and RESULT=''"
        
        Set AdoRs_SQL = New ADODB.Recordset
        AdoRs_SQL.CursorLocation = adUseClient
        AdoRs_SQL.Open sqlDoc, AdoCn_SQL
        
        If AdoRs_SQL.RecordCount = 0 Then
            Set f_subSet_TestList = Nothing
        Else
            Set f_subSet_TestList = AdoRs_SQL
        End If
    
        Set AdoRs_SQL = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_SQL = Nothing
    Set AdoRs_SQL = Nothing
    Call ErrMsgProc(CallForm)

    
End Function

Private Sub cmdWordQuery_Click()
'    On Error GoTo ErrRoutine
'    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
'
'    Dim strKeyno    As String
'    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
'    Dim strTestcd() As String, strTPid()    As String, strTPnm() As String
'    Dim strEqpCd    As String
'    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
'    Dim itemX   As ListItem
'
'    '-- WorkList조회
'    Set mAdoRs = f_subSet_WorkList(mskOrdDate.Text)
'
'    If RecordChk = False Then
'        Exit Sub
'    End If
'
''    With spdWorkList
''        .maxrows = 14
''        .Col = 1:   .Col2 = .MaxCols
''        .Row = 1:   .Row2 = .maxrows
''        .BlockMode = True
''        .Action = ActionClearText
''        .BlockMode = False
''        .RowHeight(-1) = 12
''    End With
'
'    intRow = 0
'    Do Until mAdoRs.EOF
'        intIdx = 0
'        With spdResult1
'            If strKeyno <> mAdoRs.Fields("EXAM_NO") Then
''                intRow = SeqNullSearch(spdResult1, "", 1)
''                If intRow = "0" Then
''                    .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
''                    intRow = .maxrows
''                Else
''                    intRow = intRow + 1
''                End If
'                intRow = intRow + 1
'                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
'
'                .SetText 1, intRow, "1"
'                .SetText 2, intRow, Trim(mAdoRs("REQUEST_DATE")) & ""
'                '.SetText 3, intRow, Trim(cboRegGbn.Text) & ""
'                .SetText 4, intRow, Trim(mAdoRs("PERSON_NAME")) & ""
'                .SetText 5, intRow, Trim(mAdoRs("EXAM_NO")) & ""
'                .SetText 6, intRow, Trim(mAdoRs("CHART_NO")) & ""
'                '.SetText 6, intRow, Trim(mAdoRs("COMPANY_NAME"))
'
'                '-- 검사항목조회
''                blnFlag = False
'                Set mAdoRs1 = f_subSet_TestList(mAdoRs.Fields("EXAM_NO"))
'                If Len(mAdoRs.Fields("EXAM_NO")) > 0 Then
'                    Do Until mAdoRs1.EOF
'                        strEqpCd = f_funGet_CODE(Trim(mAdoRs1("EXAM_CODE")))
'                        Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
'                        If Not itemX Is Nothing Then
''                            blnFlag = True
'                            spdResult1.Row = intRow
'                            spdResult1.Col = itemX.Index + 6
'                            spdResult1.BackColor = &HC6FEFF '&H80C0FF
'                            DoEvents
'                        End If
'                        mAdoRs1.MoveNext
'                    Loop
'                End If
'
'            End If
'            strKeyno = mAdoRs("EXAM_NO")
'        End With
'        intIdx = intIdx + 1
'        mAdoRs.MoveNext
'    Loop
'    Exit Sub
'
'ErrRoutine:
'
'    Call ErrMsgProc(CallForm)

End Sub

Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim itemX As ListItem
    Dim blnFlag As Boolean
    Dim strBarno    As String, strSPid  As String, strSPnm   As String, strChartNo As String, strSex As String
    Dim strWDate As String
    Dim strEqpCd    As String
    Dim tmpDate     As String
    Dim intCol      As Integer
    Dim intCnt      As Integer
    Dim strKey1 As String, strKey2 As String, strKey3 As String
    
    blnFlag = False
    On Error Resume Next
    
    With spdWorklist
        For intRow1 = 1 To .maxrows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, intRow1, varTmp:    strWDate = Trim$(varTmp)
                .GetText 3, intRow1, varTmp:    strBarno = Trim$(varTmp)
                .GetText 4, intRow1, varTmp:    strSPnm = Trim$(varTmp)
                .GetText 5, intRow1, varTmp:    strKey1 = Trim$(varTmp)
                .GetText 6, intRow1, varTmp:    strKey2 = Trim$(varTmp)
                .GetText 7, intRow1, varTmp:    strKey3 = Trim$(varTmp)
                
                .Row = intRow1:
                
                .Col = 1: .ForeColor = HNC_Red
                .Col = 2: .ForeColor = HNC_Red
                .Col = 3: .ForeColor = HNC_Red
                .Col = 4: .ForeColor = HNC_Red
                
                intRow2 = f_funGet_SpreadRow(spdResult1, 3, strBarno)
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 13
                        intRow2 = spdResult1.maxrows
                    End If

                    blnFlag = False
                    
                    tmpDate = Mid(strWDate, 1, 4) & Mid(strWDate, 6, 2) & Mid(strWDate, 9, 2)
                    
                    Set mAdoRs = f_subSet_WorkList(tmpDate, strBarno, strBarno)

                    '==================================================================================
                    If cboChk.ListIndex = 0 Then
                        If Len(strBarno) > 0 Then
                            Do Until mAdoRs.EOF
                                strEqpCd = f_funGet_CODE(Trim(mAdoRs("CODE")))
                                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                                If Not itemX Is Nothing Then
                                    blnFlag = True
                                    spdResult1.Row = intRow2
                                    spdResult1.Col = itemX.Index + 7
                                    spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                    spdResult1.text = " "
                                    
                                    DoEvents
                                End If
                                mAdoRs.MoveNext
                            Loop
                        End If
                    Else
                        If Len(strBarno) > 0 Then
                            For intCnt = 0 To mAdoRs.RecordCount - 1
                            'Do Until mAdoRs.EOF
                                strEqpCd = ""
                                For intCol = 1 To 5
                                    Select Case intCol
                                        Case 1: strEqpCd = "5"  '"033"
                                        Case 2: strEqpCd = "6"  '"034"
                                        Case 3: strEqpCd = "7"  '"035"
                                        Case 4: strEqpCd = "11" '"047"
                                        Case 5: strEqpCd = "13" '"055"
                                    End Select
                                
                                    Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                                    If Not itemX Is Nothing Then
                                        blnFlag = True
                                        spdResult1.Row = intRow2
                                        spdResult1.Col = itemX.Index + 7
                                        spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                        spdResult1.text = " "
                                        
                                        DoEvents
                                    End If
                                    mAdoRs.MoveNext
                                Next
                            'Loop
                            Next
                        End If
                    End If
                    '==================================================================================
                    
                    If blnFlag = True Then
                        spdResult1.SetText 1, intRow2, "1"
                        spdResult1.SetText 2, intRow2, strWDate
                        spdResult1.SetText 3, intRow2, strBarno
                        spdResult1.SetText 4, intRow2, strSPnm
                        spdResult1.SetText 5, intRow2, strKey1
                        spdResult1.SetText 6, intRow2, strKey2
                        spdResult1.SetText 7, intRow2, strKey3
                        
                        spdResult1.Row = intRow2:
                        spdResult1.Col = 7:
                        spdResult1.ForeColor = HNC_Red
                    Else
                        spdResult1.maxrows = spdResult1.maxrows - 1
                    End If
                End If
                
                .SetText 1, intRow1, ""
            End If
        Next
    End With
                
End Sub


'Private Sub spdResult1_DblClick(ByVal Col As Long, ByVal Row As Long)
'Dim TmpYesno As String
'Dim Tmpptno, TmpPtnm As String
'
'    If Row = 0 Then
'
'        If Col = 1 Then
'            Col = 2
'        End If
'
'        If OrderSort_Flag = 1 Then
'            Call SpreadSheetSort(spdResult1, Col, 2)
'            OrderSort_Flag = 2
'        Else
'            Call SpreadSheetSort(spdResult1, Col, 1)
'            OrderSort_Flag = 1
'        End If
'
'        Exit Sub
'    End If
'
'
'    If Col = 4 Or Col = 6 Then
'        With spdResult1
'            .Row = Row
'
'            ' 병록번호 불러오기
'            .Col = 6
'            Tmpptno = .text
'
'            ' 환자이름 불러오기
'            .Col = 4
'            TmpPtnm = .text
'        End With
'
'        If Len(Trim(Tmpptno)) >= 1 And Len(Trim(TmpPtnm)) >= 1 Then
'             TmpYesno = MsgBox(Tmpptno & " (  " & TmpPtnm & "  ) " & " 환자를 선택 하셨습니다..    " & vbCrLf & vbCrLf & "검사를 제외 하시겠습니까..??", vbCritical + vbYesNo, App.Title)
'
'             If TmpYesno = vbYes Then
'                spdResult1.Action = ActionDeleteRow
'                spdResult1.maxrows = spdResult1.maxrows - 1
'             End If
'        End If
'    End If
'
'End Sub

'Private Sub spdResult1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'
'Dim oMenu As cPopupMenu
'Dim lMenuChosen As Long
'
'    Set oMenu = New cPopupMenu
'
'    lMenuChosen = oMenu.Popup(" ▒ 검사자 추가", "-", " ▒ 검사자 삭제", "-", " ▒ 시작번호수정", "-", " ▒ 서버 저장")
'
'    Select Case lMenuChosen
'        Case 1
'            With spdResult1
'                .maxrows = .maxrows + 1
'                .Col = Col
'                .Row = Row
'                .Action = ActionInsertRow
'            End With
'        Case 3
'            With spdResult1
'                .Col = Col
'                .Row = Row
'                .Action = ActionDeleteRow
'                .maxrows = .maxrows - 1
'            End With
'        Case 5
'            Call cmdStartNo_Click
'        Case 7
'            Call cmdAppend_Click(0)
'    End Select
'End Sub


Private Sub comEQP_OnComm()
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim Arr()       As Byte
    Dim strdata     As String
    Dim brStr As String
    Dim sStxCheck As Integer, sEtxCheck As Integer, sCrcheck As Integer
    Dim com_sTemp As String
    Dim ii As Integer, jj As Integer
    Dim MHead  As String, pInfo As String
    Dim PatientID As String
    
    Dim Orderoutput As String
    Dim OutPutData  As String
    Dim Rev As Long
    Dim Test_Cd() As String, strPid()    As String, strPnm() As String
    Dim sRow As Integer
    Dim oPatNo As String
    Dim oRackNo As String
    Dim oPosNo As String
    Dim oIdNo As String
    
    Dim adoRS As ADODB.Recordset
    Dim sqlDoc As String
    Dim itemX As ListItem
    Dim strEqpCd1 As String
    
    Dim varTmp  As Variant
    Dim intCol  As Integer
    Dim strLevel() As String

    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            
            Buffer = comEQP.Input
            Print #1, "[Rx]" & Buffer;
            
'            Buffer = ReceiveData
            
            Debug.Print Buffer
            
            lngBufLen = Len(Buffer)
            With cInterface
                For i = 1 To lngBufLen
                    BufChar = Mid$(Buffer, i, 1)
                    
                    Select Case intPhase
                        Case 1      '## Estabilshment Phase
                            Select Case BufChar
                                Case ENQ
                                    Erase strRecvData
                                    intPhase = 2
                                    comEQP.Output = ACK
                                    Print #1, "[Tx]" & ACK;
                                Case ACK
                                    If strState = "Q" Then Call SendOrder
                                
                            End Select
                        Case 2      '## Transfer Phase
                            Select Case BufChar
                                Case ENQ
                                    Erase strRecvData
                                    comEQP.Output = ACK
                                    Print #1, "[Tx]" & ACK;
                                Case STX
                                    If intBufCnt = 0 Then
                                        intBufCnt = 1
                                        Erase strRecvData
                                        ReDim Preserve strRecvData(intBufCnt)
                                    Else
                                        intBufCnt = intBufCnt + 1
                                        ReDim Preserve strRecvData(intBufCnt)
                                    End If
                                Case ETB
                                    intPhase = 3
                                    blnIsETB = True
                                Case ETX
                                    intBufCnt = intBufCnt + 1
                                    ReDim Preserve strRecvData(intBufCnt)
                                    intPhase = 3
                                Case vbCr
                                    intBufCnt = intBufCnt + 1
                                    ReDim Preserve strRecvData(intBufCnt)
                                Case EOT
                                    intPhase = 1
                                Case Else
                                    If blnIsETB = False Then
                                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                                    Else
                                        blnIsETB = False
                                    End If
                            End Select
                        Case 3      '## Transfer Phase
                            Select Case BufChar
                                Case vbCr
                                    intPhase = 4
                                    comEQP.Output = ACK
                                    Print #1, "[Tx]" & ACK;
                            End Select
                        Case 4      '## Termination Phase
                            Select Case BufChar
                                Case STX
                                    intPhase = 2
                                Case EOT
                                    Call EditRcvData
                                    .clearRcvbuf
                                    If strState = "Q" Then
                                        intSndPhase = 1
                                        intFrameNo = 1
                                        comEQP.Output = ENQ
                                        Print #1, "[Tx]" & ENQ;
                                    End If
                                    intPhase = 1
                            End Select
                    End Select
                Next i
            End With
            
        Case comEvSend
        
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        Case comEvCTS
            strEVMsg = " CTS(Clear to Send) 변경 감지"
        Case comEvDSR
            strEVMsg = " DSR(Data Set Read) 변경 감지"
        Case comEvCD
            strEVMsg = " CD(Carrier Detecr) 변경 감지"
        Case comEvRing
            strEVMsg = " 전화 벨이 울리는 중"
        Case comEvEOF
            strEVMsg = " EOF(End Of File) 감지"

        ' 오류 메시지
        Case comBreak
            strERMsg = " 중단 신호 수신"
        Case comCDTO
            strERMsg = " 반송파 검출 시간 초과"
        Case comCTSTO
            strERMsg = " CTS(Clear to Send) 시간 초과"
        Case comDCB
            strERMsg = " 포트에 대한 장치 제어 블록(DCB) 검색 중 예기치 못한 오류"
        Case comDSRTO
            strERMsg = " DSR(Data Set Read) 시간 초과"
        Case comFrame
            strERMsg = " 프레이밍 오류"
        Case comOverrun
            strERMsg = " 패리티 오류"
        Case comRxOver
            strERMsg = " 수신 버퍼 초과"
        Case comRxParity
            strERMsg = " 패리티 오류"
        Case comTxFull
            strERMsg = " 전송 버퍼에 여유가 없음"
        Case Else
            strERMsg = " 알 수 없는 오류 또는 이벤트"
    End Select
    If Len(strERMsg) > 0 Then Call ShowMessage(strERMsg)
        
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, 스프레드에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
'Private Sub GetOrder(ByVal pBarNo As String)
'    Dim objOrder As clsIntOrder
'    'Dim intRow      As Integer
'    Dim strEqpCd    As String
'    'Dim i           As Integer
'    Dim strSexAge  As String
'    Dim itemX As ListItem
'
'    Dim i           As Integer
'    Dim intRow      As Long
'    Dim strItems    As String
'
'    intRow = -1
'    For i = 1 To spdResult1.DataRowCnt
'        If Trim(GetText(spdResult1, i, 4)) = pBarNo Then
'            intRow = i
'            Exit For
'        End If
'    Next i
'
'    If intRow < 0 Then
'        intRow = spdResult1.DataRowCnt + 1
'        If spdResult1.maxrows < intRow Then
'            spdResult1.maxrows = intRow
'        End If
'    End If
'
'    Call SetText(spdResult1, pBarNo, intRow, 4)  '3
''    Call SetText(spdResult1, mOrder.RackNo, intRow, colRack)       '4
''    Call SetText(spdResult1, mOrder.TubePos, intRow, colPos)         '5
'    Call vasActiveCell(spdResult1, intRow, 4)
'    Call ClearSpread(spdRstview)
'
''    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
'    If Mid(pBarNo, 1, 2) = "99" Then
'        Call Get_Sample_Info_QC(intRow)
'    Else
'        Call Get_Sample_Info(intRow)                        '2,6,7,8,9
'    End If
'
'    '-- 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.
'    '-- intRow 추가
'    strItems = GetEquipExamCode_VISTA(gEquip, pBarNo, intRow)
'
'    If strItems <> "" Then
'        Call ErrWrite(pBarNo & "의 검사항목이 없습니다")
'    End If
'
'    If Trim(strItems) = "" Then
'        mOrder.NoOrder = True
'        mOrder.Order = ""
'    Else
'        mOrder.NoOrder = False
'        mOrder.Order = strItems
'    End If
'
'
'End Sub


'Function GetEquipExamCode_VISTA(argEquipCode As String, argPID As String, Optional intRow As Long) As String
''검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
''한 장비 번호에 검사코드가 1개이상 존재
'Dim i As Integer
'Dim sExamCode As String
'Dim strExamCode As String
'Dim sSpecNo     As String
'Dim strChannel As String
'Dim rs_Vista As ADODB.Recordset
'Dim iRow        As Long
'Dim SpecNo      As String
'
'    GetEquipExamCode_VISTA = ""
'
'    If Trim(argEquipCode) = "" Then
'        Exit Function
'    End If
'
'    '-- 자검체는 11자리임 조회하기위하여 마지막 자리를 없앤다.
'    argPID = Mid(argPID, 1, 10)
'
'    If Mid(argPID, 1, 2) = "99" Then
'        'strExamCode = Proc_Order_LX_QC(argPID)
'
'        'iRow = frmInterface.spdResult1.DataRowCnt
'        iRow = intRow
'
'        SpecNo = Trim(GetText(frmInterface.spdResult1, iRow, colSpecNo))
'
'        sql = "SELECT QC_EXMN_CD "
'        sql = sql & vbCrLf & " FROM SPSLMQMST "
'        sql = sql & vbCrLf & "WHERE EQPM_CD = '" & Mid(SpecNo, 3, 3) & "' "     '//// 장비 번호
'        sql = sql & vbCrLf & "  AND SBSN_CD = '" & Mid(SpecNo, 6, 3) & "' "     '//// 검사명 번호
'        sql = sql & vbCrLf & "  AND LVL_CD = '" & Mid(SpecNo, 9, 1) & "' "      '//// 레벨 번호
'        sql = sql & vbCrLf & "  AND QC_EXMN_CD IN (" & gAllExam & ") "
'        res = db_select_Row(gServer, sql)
'        strExamCode = ""
'
'        For i = 0 To UBound(gReadBuf)
'            If gReadBuf(i) <> "" Then
'                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
'            Else
'                Exit For
'            End If
'        Next
'    Else
'
'        '바코드번호로 검체번호 불러오기
'        sql = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
'        res = db_select_Col(gServer, sql)
'        sSpecNo = Trim(gReadBuf(0))
'
'        '-- 검사코드 가져오기
'        sql = " Select EXMN_CD From SPSLHRRST " & CR & _
'              " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
'              "   and RSLT_NO IS NOT NULL"
'
'        res = db_select_Row(gServer, sql)
'        strExamCode = ""
'
'        For i = 0 To UBound(gReadBuf)
'            If gReadBuf(i) <> "" Then
'                strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
'            Else
'                Exit For
'            End If
'        Next
'    End If
'
'    If strExamCode = "" Then
''        MsgBox "미접수 환자"
'        GetEquipExamCode_VISTA = ""
'        Exit Function
'    End If
'    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
'    'EquipExamCode =
'
'    ClearSpread frmInterface.vasTemp1
''    sExamCode = ""
'
'    Set rs_Vista = New ADODB.Recordset
'
'    '-- 가져온 검사코드의 채널 찾기
'          sql = "Select distinct equipcode,testno "
'    sql = sql & "  From EquipExam "
'    sql = sql & " Where equipno  = '" & Trim(gEquip) & "' "
'    sql = sql & "   and examcode in (" & Trim(strExamCode) & ")"
'
'    'res = db_select_Row(gLocal, SQL)
'
'    strExamCode = ""
'    Set rs_Vista = cn.Execute(sql)
'    Do Until rs_Vista.EOF
'        If Trim(rs_Vista.Fields("testno").Value & "") <> "" And Trim(rs_Vista.Fields("equipcode").Value & "") <> "" Then
'            strChannel = Trim(rs_Vista.Fields("testno").Value & "") & "^^" & Trim(rs_Vista.Fields("equipcode").Value & "")
'            strExamCode = strExamCode & "\^^^" & strChannel
'        End If
'        rs_Vista.MoveNext
'    Loop
'
'    GetEquipExamCode_VISTA = Mid(strExamCode, 2)
'
'    Set rs_Vista = Nothing
'
'End Function

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarno     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
    Dim strIntResult As String   '수신한 결과
    Dim strRefVal    As String
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strRack      As String
    Dim strPos       As String
    Dim strTemp      As String
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim strTemp3     As String
    Dim varTemp3     As Variant
    Dim intCnt       As Integer
    Dim pDocount As Integer
    Dim itemX As ListItem
    Dim varTmp
    Dim intCol As Integer
    Dim pGrid_Point As Integer
    Dim sqlDoc As String
    Dim strName
    Dim strRackPos
    
    Dim strKey1, strKey2, strKey3
    
    Dim strItemX As String
    Dim strItemTag As String
    Dim blnItemX As Boolean
    
    Dim strSexAge As String
    Dim strSvrcData As Variant
    Dim varSvcData As Variant
    Dim varSvcData_1 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim blnSameCode As Boolean
    
    Dim strPatInfo  As String   '에러메세지 표시용 : 환자명(병록번호)
    Dim adoRS2      As ADODB.Recordset
    Dim strTstCD As String
    Dim strChannel, strItems As String
    Dim strBUN, strCREA As String
    Dim strGFR As String
    Dim strAGE, strSex As String
    Dim strDoc As String
    Dim strEqpCd As String
    Dim blnQCChk    As Boolean
    Dim strLevel    As String
    Dim strAnalyte  As String
    Dim strUnit     As String
    Dim strPR3 As String
    Dim strMPO As String
    Dim strPR3C As String
    Dim strMPOC As String
    Dim strSpcDt As String
    
    Dim sUrl, sPost, sParam As String
    Dim sRcvData, sData As String
    Dim varRcvData As Variant
    Dim varTstCode As Variant
    Dim varEqpCode As Variant
    Dim strQCFlag  As String
    Dim strItemName As String
    Dim strLot  As String
    Dim strSmartLot  As String
    Dim strSmartLvl  As String
    Dim varTstCd    As Variant
    Dim strSResult  As String
    Dim strTmp1  As String
    Dim varOrderInfo As Variant
    Dim varSEQinfo As Variant
    Dim varORDinfo As Variant
    Dim strSeqNo  As String
    Dim intIdx As Integer
    Dim sCol As Integer
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD As String
    Dim intOrdCnt As Integer
    Dim strRstval As String
    Dim strDate As String
    Dim strTime As String
'    Dim varTstCd As Variant
    Dim s As Integer
    Dim strOrder(4) As String
    
    Set objResult = New clsIntResults
    
    On Error Resume Next
    
    With cInterface
        For intCnt = 1 To UBound(strRecvData) '.bufcnt
            strRcvBuf = strRecvData(intCnt)
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                    strAlarm = ""
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                    '## 바코드번호, SEQ, Disk No, Tube Position 조회
                    If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                    
                    mOrder.NoOrder = False
                    mOrder.BarNo = strBarno
                    mOrder.Seq = mGetP(strTemp1, 3, "^")
                    mOrder.RackNo = mGetP(strTemp1, 4, "^")
                    mOrder.TubePos = mGetP(strTemp1, 5, "^")
                    
                    Set objIntInfo = New clsIntInfo
                    With objIntInfo
                        .BarNo = strBarno
                        .SpcPos = strPos & "/" & strRackNo
                        
                    End With
                                        
                    '-- 오더조회
                    
                    '바코드번호로 오더찾아오기
                    RecordChk = False
                    Set mAdoRs = f_subSet_WorkList_Barcode(strBarno)
                    
                    If RecordChk = True Then
                        strOrder(1) = ""
                        strOrder(2) = ""
                        strOrder(3) = ""
                        strOrder(4) = ""
                        
                        strTstCD = ""
                        mOrder.Testcd = ""
                        Do Until mAdoRs.EOF
                            intIdx = 0
                            With spdResult1
                                sCol = 5
                                pGrid_Point = SeqSearch(spdResult1, strBarno, sCol)
                                If pGrid_Point = 0 Then
                                    pGrid_Point = SeqNullSearch(spdResult1, strBarno, sCol)
                                    If pGrid_Point = 0 Then
                                        .maxrows = .maxrows + 1
                                        .RowHeight(.maxrows) = 13
                                    End If
                                    pGrid_Point = .maxrows
                                End If
                                
                                If intOrdCnt = 0 Then
                                    .SetText 1, pGrid_Point, "1"
                                    strSeqNo = mAdoRs("SCP42IDNOA")
                                    .SetText 2, pGrid_Point, strSeqNo 'mAdoRs("SCP42IDNOA") & ""
                                    .SetText 3, pGrid_Point, mAdoRs("SCP41NAME") & ""
                                    .SetText 4, pGrid_Point, mAdoRs("SCP41JDATE") & ""
                                    .SetText 5, pGrid_Point, strBarno
                                    .SetText 6, pGrid_Point, mAdoRs("SCP42SUGACD") & ""
                                End If
                                
                                For intCol = 10 To .MaxCols
                                    .GetText intCol, 0, varTmp
                                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                    If Not itemX Is Nothing Then
                                    
                                        If Trim(mAdoRs.Fields("SCP42SUGACD").Value) = "L070355" Or Trim(mAdoRs.Fields("SCP42SUGACD").Value) = "L070356" Then
                                            strTstCD = strTstCD & "'L070355,L070356',"
                                        ElseIf Trim(mAdoRs.Fields("SCP42SUGACD").Value) = "L07035501" Or Trim(mAdoRs.Fields("SCP42SUGACD").Value) = "L07035601" Then
                                            strTstCD = strTstCD & "'L07035501,L07035601',"
                                        ElseIf Trim(mAdoRs.Fields("SCP42SUGACD").Value) = "L070350" Or Trim(mAdoRs.Fields("SCP42SUGACD").Value) = "L070351" Then
                                            strTstCD = strTstCD & "'L070350,L070351',"
                                        Else
                                            strTstCD = strTstCD & "'" & Trim(mAdoRs.Fields("SCP42SUGACD").Value) & "',"
                                        End If
                                        
                                        mOrder.Testcd = mOrder.Testcd & Trim(mAdoRs.Fields("SCP42SUGACD").Value) & "|"
                                        
                                        
                                        varTstCd = Split(Trim(itemX.SubItems(1)), ",")
                                        
                                        For s = 0 To UBound(varTstCd)
                                            'If mAdoRs("SCP42SUGACD") & "" = Trim(itemX.SubItems(1)) Then
                                            If mAdoRs("SCP42SUGACD") & "" = Trim(varTstCd(s)) Then
                                                intOrdCnt = intOrdCnt + 1
                                                spdResult1.Row = pGrid_Point
                                                spdResult1.Col = itemX.Index + 9
                                                spdResult1.BackColor = &HC6FEFF   '&HC6FEFF
                                                GoTo Rst
                                                Exit For
                                            End If
                                        Next
                                    End If
                                    'Exit For
                                Next
Rst:
                            End With
                            mAdoRs.MoveNext
                        Loop
                    Else
                        lblStatus.Caption = "바코드 번호 " & strBarno & " 는 검사대상이 아닙니다"
                    End If
                                                                                                    
                    Set mAdoRs = Nothing
                    
                     
                     If Trim(strTstCD) <> "" Then
                         strTstCD = Mid(strTstCD, 1, Len(strTstCD) - 1)
                     End If
                    
                     strItems = ""
                     If Trim(strTstCD) <> "" Then
                         Set adoRS2 = New ADODB.Recordset
                         
                                  sqlDoc = "select TESTCD_EQP "
                         sqlDoc = sqlDoc & "  from INTERFACE002 "
                         sqlDoc = sqlDoc & " where EQP_CD = '" & INS_CODE & "'"
                         sqlDoc = sqlDoc & "   and TESTCD in (" & strTstCD & ")"
                         
                         adoRS2.CursorLocation = adUseClient
                         adoRS2.Open sqlDoc, AdoCn_Jet
                         
                         If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
                         
                         Do While Not adoRS2.EOF
                             If Trim(adoRS2.Fields("TESTCD_EQP").Value & "") <> "" Then
                                 strChannel = Trim(adoRS2.Fields("TESTCD_EQP").Value & "")
                                 'If strChannel = "1" Or strChannel = "2" Or strChannel = "3" Then
                                 '   strItems = "\^^^1\^^^2\^^^3"
                                 '   varEqpCode = varEqpCode & "|" & strChannel
                                 'Else
                                 '   strItems = strItems & "\^^^" & strChannel
                                 '   varEqpCode = varEqpCode & "|" & strChannel
                                 'End If
                                If strChannel = "1" Then
                                    strOrder(1) = "\^^^1"
                                ElseIf strChannel = "2" Then
                                    strOrder(2) = "\^^^2"
                                ElseIf strChannel = "3" Then
                                    strOrder(3) = "\^^^3"
                                ElseIf strChannel = "4" Then
                                    strOrder(4) = "\^^^4"
                                End If
                                
                                RecordChk = True
                             End If
                             adoRS2.MoveNext
                         Loop
                         Set adoRS2 = Nothing
                         adoRS2.Close
                     End If
                    
                    If strOrder(1) <> "" Or strOrder(2) <> "" Or strOrder(3) <> "" Then
                        strItems = "\^^^1\^^^2\^^^3" & strOrder(4)
                    Else
                        strItems = strOrder(1) & strOrder(2) & strOrder(3) & strOrder(4)
                    End If
                    If strItems <> "" Then
                        strItems = Mid(strItems, 2)
                    End If
                    
                    If varEqpCode <> "" Then
                        varEqpCode = Mid(varEqpCode, 2)
                    End If
                    
                    If Trim(strItems) = "" Then
                        mOrder.NoOrder = True
                        mOrder.Order = ""
                    Else
                        mOrder.NoOrder = False
                        mOrder.Order = strItems
                    End If
                    
                    strState = "Q"
                    
                Case "O"    '## Order
                    '## 바코드번호, Rack No, Tube Postion 조회
                    strState = ""
                    strResult = ""
                    strIntResult = ""
                    strSResult = ""
                    strQCFlag = Right(Trim(strRcvBuf), 1)
                    strTemp = Trim$(mGetP(strRcvBuf, 3, "|"))
                    strBarno = mGetP(strTemp, 1, "^")
'strBarno = "2512001437"
'                    If strQCFlag = "Q" Then
'                        strBarno = Trim$(mGetP(strRcvBuf, 4, "|"))
'                        strLot = mGetP(strTemp, 1, "^")
'                    Else
'                        strBarno = mGetP(strTemp, 1, "^")
'                        strLot = Trim$(mGetP(strRcvBuf, 4, "|"))
'                    End If
                    
                    If strBarno <> "" Then
                        pGrid_Point = SeqSearch(spdResult1, strBarno, colPAID)
                        If pGrid_Point = 0 Then
                            pGrid_Point = SeqNullSearch(spdResult1, strBarno, colPAID)
                        End If
                    Else
                        pGrid_Point = SeqNullSearch(spdResult1, strBarno, colPAID)
                    End If
                    
                    If pGrid_Point = 0 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        pGrid_Point = spdResult1.maxrows
                    End If
                    
                    '바코드번호로 오더찾아오기
                    RecordChk = False
                    If Len(strBarno) = 10 Then
                        Set mAdoRs = f_subSet_WorkList_Barcode(strBarno)
                    End If
                    If RecordChk = True Then
                        k = 0
                        For i = 0 To 100
                            strSNOCode(i) = ""
                        Next
                        k = 0
                        Do Until mAdoRs.EOF
                            intIdx = 0
                            With spdResult1
                                sCol = 5
                                pGrid_Point = SeqSearch(spdResult1, strBarno, colPAID)
                                If pGrid_Point = 0 Then
                                    pGrid_Point = SeqNullSearch(spdResult1, strBarno, colPAID)
                                    If pGrid_Point = 0 Then
                                        .maxrows = .maxrows + 1
                                        .RowHeight(.maxrows) = 13
                                    End If
                                    pGrid_Point = .maxrows
                                End If
                                
                                .SetText 1, pGrid_Point, "1"
                                .SetText 2, pGrid_Point, mAdoRs("SCP42IDNOA") & ""
                                .SetText 3, pGrid_Point, mAdoRs("SCP41NAME") & ""
                                .SetText 4, pGrid_Point, mAdoRs("SCP41JDATE") & ""
                                .SetText 5, pGrid_Point, strBarno
                                .SetText 6, pGrid_Point, mAdoRs("SCP42SUGACD") & ""
                            End With
                            
                            strSNOCode(k) = Trim(mAdoRs("SCP42SUGACD"))
                            k = k + 1
                            
                            For intCol = 10 To spdResult1.MaxCols
                                spdResult1.GetText intCol, 0, varTmp
                                Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                If Not itemX Is Nothing Then
                                    If mAdoRs("SCP42SUGACD") & "" = Trim(itemX.SubItems(1)) Then
                                        'strOrdBuffer = strOrdBuffer & "DSP|" & 29 + intOrdCnt & "||" & Trim(itemX.Tag) & "^^^|||" & vbCr + vbLf '검사채널(test id)
                                        'intOrdCnt = intOrdCnt + 1
                                        spdResult1.Col = itemX.Index + 9
                                        spdResult1.BackColor = &HC6FEFF   '&HC6FEFF
    
                                        Exit For
                                    End If
                                End If
                            Next
                            mAdoRs.MoveNext
                        Loop
                    Else
                        'spdResult1.SetText 1, pGrid_Point, "0"
                        'spdResult1.SetText 2, pGrid_Point, ""
                        'spdResult1.SetText 3, pGrid_Point, ""
                        'spdResult1.SetText 4, pGrid_Point, ""
                        spdResult1.SetText 5, pGrid_Point, strBarno
                        'spdResult1.SetText 6, pGrid_Point, ""
                        
                        lblStatus.Caption = "바코드 번호 " & strBarno & " 는 검사대상이 아닙니다"
                    End If
                                                                                                    
                    Set mAdoRs = Nothing
                    
                    
                    If pGrid_Point > 0 Then
                        spdResult1.GetText colORDT, pGrid_Point, varTmp:    strORDT = Trim$(varTmp)
                        spdResult1.GetText colORQN, pGrid_Point, varTmp:    strORQN = Trim$(varTmp)
                        spdResult1.GetText colPANM, pGrid_Point, varTmp:    strPANM = Trim$(varTmp)
                        spdResult1.GetText colPAID, pGrid_Point, varTmp:    strPAID = Trim$(varTmp)
                        spdResult1.GetText colOIFL, pGrid_Point, varTmp:    strOIFL = Trim$(varTmp)
                        spdResult1.GetText colSENO, pGrid_Point, varTmp:    strSENO = Trim$(varTmp)
                        spdResult1.GetText colSEXS, pGrid_Point, varTmp:    strSEXS = Trim$(varTmp)
                        spdResult1.GetText colAGES, pGrid_Point, varTmp:    strAGES = Trim$(varTmp)
                        spdResult1.GetText colNWNO, pGrid_Point, varTmp:    strNWNO = Trim$(varTmp)
                    End If
                
                
                Case "R"    '## Result
'                    strTemp = mGetP(strRcvBuf, 3, "|")
'                    strIntBase = mGetP(strTemp, 4, "^")
'                    strFlag = mGetP(strTemp, 5, "^")

                    strTemp = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp, 4, "^")
                    strFlag = mGetP(strRcvBuf, 9, "|")
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    
                    Select Case strFlag
                        Case "F"    '## 정량
                            'strIntBase = strIntBase & "N"
                            strResult = strIntResult
                        Case "I"    '## 정성
'                            strIntBase = strIntBase & "C"
'                            Select Case Mid$(strIntResult, 1, 1)
'                                Case "N":   strResult = "Negative"
'                                Case "G":   strResult = "GrayZone"
'                                Case "R":   strResult = "Positive"
'                                Case "P":   strResult = "Positive"
'                            End Select
                    End Select

                    If strIntBase <> "" And strResult <> "" Then
                        strRstval = strResult
                        For intCol = 10 To spdResult1.MaxCols
                            spdResult1.GetText intCol, 0, varTmp
                            
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnSameCode = False
                                
'                                If strQCFlag = "Q" Then
'                                    blnSameCode = True
'                                Else
'                                    Set adoRS2 = New ADODB.Recordset
'
'                                             sqlDoc = "SELECT TESTCD " & vbCrLf
'                                    sqlDoc = sqlDoc & "  From INTERFACE002 " & vbCrLf
'                                    sqlDoc = sqlDoc & " Where TESTCD_EQP = '" & varTmp & "'"
'
'                                    adoRS2.CursorLocation = adUseClient
'                                    adoRS2.Open sqlDoc, AdoCn_Jet
'                                    If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
'                                    Do While Not adoRS2.EOF
'                                        For i = 0 To UBound(varTstCD)
'                                            If adoRS2.Fields("TESTCD").Value = varTstCD(i) Then
'                                                blnSameCode = True
'                                                strTmp1 = adoRS2.Fields("TESTCD").Value
'                                                Exit Do
'                                            End If
'                                        Next
'                                        adoRS2.MoveNext
'                                    Loop
'                                    adoRS2.Close:    Set adoRS2 = Nothing
'                                End If


                                'If strMPO <> "" And strPR3 <> "" And strIntBase = itemX.tag Then
                                '    blnSameCode = True
                                'End If
                                
                                If strIntBase <> "" And strIntBase = itemX.tag Then
                                    strDate = Format$(Now, "YYYYMMDD"):    strTime = Format$(Now, "HHMMSS")
                                
                                    '-- 처방번호 찾기
                                    spdResult1.GetText intCol, pGrid_Point, varTmp: strORQN = varTmp
                                    spdResult1.Col = intCol
                                    spdResult1.ForeColor = vbBlack
                                    
                                    spdResult1.SetText intCol, pGrid_Point, strRstval
                                    spdResult1.Col = 7: spdResult1.ForeColor = vbRed: spdResult1.BackColor = vbCyan
                                    spdResult1.SetText 1, pGrid_Point, "1"
                                    
                                    '-- 로컬저장
                                    sqlDoc = "Update INTERFACE003" & _
                                             "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                                             " where SPCNO   = '" & strBarno & "'" & _
                                             "   and EQPNUM  = '" & itemX.tag & "'" & _
                                             "   and TRANSDT = '" & strDate & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc
                                    
                                    sqlDoc = "insert into INTERFACE003(" & _
                                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                             "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
                                             "            '" & strDate & "', '" & strTime & "'," & _
                                             "            '" & strRstval & "', '" & strRefVal & "'," & _
                                             "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                    AdoCn_Jet.Execute sqlDoc
                                    

                                    Set itemX = Nothing
                                    
                                    strState = "R"
                                    
                                    Call spdResult1_Click(2, pGrid_Point)
                                    
'                                    If strMPO <> "" And strPR3 <> "" Then
'                                        strIntBase = "ANCA sc"
'                                        If strMPO = "<0.2" And strPR3 = "<0.2" Then
'                                            strResult = "0/Negative"
'                                        Else
'                                            strResult = "0/Positive"
'                                        End If
'                                        strMPO = ""
'                                        strPR3 = ""
'                                        GoTo RST2:
'                                    End If
                                    strResult = ""
                                    Exit For
                                End If
                                Set itemX = Nothing
                            End If
                        Next
                        
'                        If strMPO <> "" And strPR3 <> "" Then
'                            strIntBase = "ANCA sc"
'                            If strMPO = "<0.2" And strPR3 = "<0.2" Then
'                                strResult = "0/Negative"
'                            Else
'                                strResult = "0/Positive"
'                            End If
'                            strMPO = ""
'                            strPR3 = ""
'                            GoTo RST2:
'                        End If
                        
                    End If
                Case "C"
                Case "L"    '## Terminator
                    '## DB에 결과저장
                    If chkAuto.Value = "1" And strState = "R" Then
                        Call cmdAppend_Click(0)
                        
                        Set objIntInfo = Nothing
                        strState = ""
                    End If

            End Select
        Next
    End With

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput   As String     '송신할 데이터
    Dim sqlDoc      As String
    Dim varTestCd   As Variant
    Dim i           As Integer
''
''    With cInterface
''        Select Case intSndPhase '.Snd_Phase
''            Case 0
''                strOutput = EOT
''                comEQP.Output = strOutput
''                Print #1, "[Tx]" & strOutput;
''                Debug.Print strOutput
''
''                strState = ""
''
''
''                Exit Sub
''
''            '-- 최초 오더전송
''            Case 1  '## Header
''                '## Header
''                strOutput = mIntLib.GetFrameNo & "H|\^&|||99^2.00" & vbCr & ETX
''
''                '## Order
''                If mOrder.NoOrder = False Then
''                    '## 접수정보가 있는경우
''                   'strOutput = strOutput & "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||S||||||||||Q" & vbCr
''                    strOutput = strOutput & "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||" & mOrder.TubePos & "||||||||||Q" & vbCr
''
''                    'Architect sample ==> O|1|MCC1||^^^16\^^^606|||20010223081223||||A|Hep|lipemic||serum||||||||||Q[CR]
''                Else
''                    '## 접수정보가 없는경우
''                    strOutput = strOutput & "O|1|" & mOrder.BarNo & "|||R||||||C||||||||||||||Q" & vbCr
''                End If
''
''                '## Termianator
''                strOutput = strOutput & "L|1|N" & vbCr
''                'strOutput = .FrameN & strOutput
''                strOutput = intFrameNo & strOutput
''
''            Case 2  '## Patient
''                strOutput = mIntLib.GetFrameNo & "P|1|||" & mOrder.AccInfo.PtId & "|^1^1^56|||19700505" & vbCr & ETX
''                intSndPhase = 4
''            Case 3  '## No Order
''                strOutput = strOutput & "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||" & mOrder.TubePos & "||||||||||Q" & vbCr
''                intSndPhase = 5
''            Case 4  '## Order
''                strOutput = mIntLib.GetFrameNo & "Q|1|^" & mOrder.BarNo & "||^^^ALL||||||||X" & vbCr & ETX
''                mIntLib.SndPhase = 5
''
''        End Select
''
''        If Len(strOutput) >= 230 Then
''            mOrder.Order = Mid$(strOutput, 231)
''            strOutput = Mid$(strOutput, 1, 230) & ETB
''            intSndPhase = 2
''            '.Snd_Phase = 2
''        Else
''            strOutput = strOutput & ETX
''            intSndPhase = 0
''            '.Snd_Phase = 0
''        End If
''
''        strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
''        comEQP.Output = strOutput
''    '    Save_Raw_Data "[Tx]" & strOutput
''        Print #1, "[Tx]" & strOutput;
''        Debug.Print strOutput
''    End With

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&|||99^2.00" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        
            If mOrder.NoOrder = False Then
                intSndPhase = 2
            Else
                intSndPhase = 3
            End If
        
        
        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1" & vbCr & ETX
                                                     '.ptid
            strOutput = intFrameNo & "P|1|||" & mOrder.BarNo & "|^1^1^56|||19700505" & vbCr & ETX
            intSndPhase = 4
            intFrameNo = intFrameNo + 1
            
        Case 3  '## No Order
            strOutput = intFrameNo & "Q|1|^" & mOrder.BarNo & "||^^^ALL||||||||X" & vbCr & ETX
            intSndPhase = 5
            
        Case 4  '## Order
            If mOrder.IsSending = False Then   '## 최초 보낼때
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R"
                            
                            '3O|1|9905300211|1^00014^1^^SAMPLE^NORMAL|ALL|R|20110613090006|||||X||||||||||||||O|||||
                            '90
                If Len(strOutput) > 230 Then
                    mOrder.IsSending = True
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 4
                Else
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 5
                End If
            Else                        '## 남은 문자열이 있을때
                strOutput = mOrder.Order
                If Len(strOutput) > 230 Then
                    mOrder.Order = Mid$(strOutput, 231)
                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                    intSndPhase = 4
                Else
                    mOrder.IsSending = False
                    strOutput = intFrameNo & strOutput & vbCr & ETX
                    intSndPhase = 5
                End If
            End If
            intFrameNo = intFrameNo + 1
        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1
            
        Case 6  '## EOT
            strState = ""
            comEQP.Output = EOT
            'Save_Raw_Data "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEQP.Output = strOutput
    Debug.Print strOutput
    'Save_Raw_Data "[Tx]" & strOutput
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열의 CheckSum을 구함
'   인수 :
'       - pMsg : 문자열
'   반환 : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 오더전송시 사용되는 FrameNo를 조회
'-----------------------------------------------------------------------------'
Public Function GetFrameNo() As Long
    mFrameNo = mFrameNo + 1
    If mFrameNo = 8 Then
        mFrameNo = 0
    End If
    GetFrameNo = mFrameNo
End Function


'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i As Long
    
    cInterface.clearRcvbuf
    
'    For i = mBuffers.Count To 1 Step -1
'        mBuffers.Remove i
'    Next i
End Sub


Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brspread As Object)

Dim sTemp       As String       ' On Com으로부터 넘겨받은 Receive Data
Dim Channel_No  As String       ' 문자형 변수
Dim Patiant_No  As String       ' 환자번호
Dim pGrid_Point As Integer      ' 해당 검사자 Point
Dim Max_Arary_Cnt As Integer    ' 검사 항목수
'-------------------------------' 임시 변수들.....
Dim sDeCnt      As Integer
Dim pDocount    As Integer
Dim Loop_count  As Integer
Dim sRtn As Integer, sChannel As String, sRstText As String, sRstValue As Single, sUnit As String
Dim itemX As ListItem
Dim strRstval(1 To 19) As String, strRefVal(1 To 19)  As String
Dim FunStr As String
Dim sqlDoc  As String
Dim intCol As Integer
Dim Test_Cd() As String, strPid()    As String, strPnm() As String
Dim Rev As Long
Dim ii As Integer
Dim tmpTstCd As String
Dim strLevel() As String
Dim chkPos  As Variant
Dim strResult As String
Dim strBarno    As String, strSPid  As String, strSPnm   As String
Dim strSex      As String, strOld   As String, strArea   As String
Dim varTmp  As Variant
Dim strDate As String, strTime  As String, sqlRet   As Integer
Dim strResultTmp As String

    On Error GoTo errDefine
    sRstText = brbarcd
'    Debug.Print "sRstText : " & sRstText
    '------------------------------<<< fElecsys2010() 배열 Clear 한다.         >>>----------
    For Loop_count = 1 To 100: fElecsys2010(Loop_count) = "": Next Loop_count
    '------------------------------<<< fElecsys2010() 배열에 구분하여 넣는다.  >>>----------
        
    pDocount = 0
'    sRstText = Mid(sRstText, STX)
    sRstText = Mid(sRstText, InStr(fRcvString, STX))
    Do While InStr(sRstText, "|") > 0
        pDocount = pDocount + 1
        fElecsys2010(pDocount) = Text_Redefine(sRstText, "|")
        sRstText = Mid$(sRstText, InStr(sRstText, "|") + 1)   ' 구분자가 "|" 이다....
        If pDocount > 99 Then
            sRstText = ""
            Exit Do
        End If
    Loop
    
    sRstText = ""
    If Mid$(fElecsys2010(1), 3, 1) = "H" Then          ' "H" Head Message Display
        comEQP.Output = ACK
'        Debug.Print "H [HOST] " & ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "P" Then      ' "P" Patiant Information Data Process
        comEQP.Output = ACK
'        Debug.Print "P [HOST] " & ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "C" Then
        comEQP.Output = ACK
'        Debug.Print "C [HOST] " & ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "Q" Then      ' "Q" Patiant Information Data Process
        comEQP.Output = ACK
'        Debug.Print "Q [HOST] " & ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "O" Then      ' "O" Order Data Process
        comEQP.Output = ACK
'        Debug.Print "O [HOST] " & ACK
        PatientID = fElecsys2010(4)
        pDocount = 0
        Do While InStr(fElecsys2010(4), "^") > 0
            pDocount = pDocount + 1
            Select Case pDocount
                Case 1:    PatientSeq = Text_Redefine(fElecsys2010(4), "^")
                Case 2:    PatientRack = Text_Redefine(fElecsys2010(4), "^")
                Case 3:    PatientPos = Text_Redefine(fElecsys2010(4), "^")
                Case Else: Exit Do
            End Select
            fElecsys2010(4) = Mid$(fElecsys2010(4), InStr(fElecsys2010(4), "^") + 1)   ' 구분자가 "^" 이다....
        Loop

        Patiant_Recevid = False        ' 환자번호 Flag
        sPatiant_No = PatientSeq ' 환자번호
        '-------------------------------------------<<< 해당검사결과와 해당환자를 찿는다.       >>>----------
        With brspread
            For pDocount = 1 To .maxrows
                .Row = pDocount: .Col = 7
                If Trim$(.text) = Trim$(Val(sPatiant_No)) Then
                    vRow = pDocount
                    Patiant_Recevid = True
                    Exit For
                End If
            Next pDocount
        End With

    ElseIf Mid$(fElecsys2010(1), 3, 1) = "R" Then
        comEQP.Output = ACK
'        Debug.Print "R [HOST] " & ACK
        Dim strChannel_No1 As String
        Dim strChannel_No2 As String
        
        If Patiant_Recevid = True Then
            strChannel_No1 = Mid(fElecsys2010(3), InStr(fElecsys2010(3), "^^^") + 3)
            strChannel_No2 = left(strChannel_No1, InStr(strChannel_No1, "^^") - 1)
            Channel_No = strChannel_No2
            With spdResult1
                For pDocount = 8 To .MaxCols
                    .Row = vRow
                    .Col = pDocount
                    .GetText 2, vRow, varTmp:    strBarno = Trim$(varTmp)
                    .GetText 4, vRow, varTmp:    strSPnm = Trim$(varTmp)
                    .GetText 7, vRow, varTmp:    strSPid = Trim$(varTmp)
                    
                    .GetText pDocount, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Channel_No = itemX.tag Then
                        If Trim(fElecsys2010(4)) <> "" Then
                            Select Case Channel_No
                                Case "900"
                                    strResult = Mid(fElecsys2010(4), InStr(fElecsys2010(4), "^") + 1)
                                Case Else
                                    strResult = Trim(fElecsys2010(4))
                            End Select
                             .text = strResult
                        Else
                            .text = ""
                        End If


                        If strResult <> "" Then
                            strDate = Format$(Now, "YYYYMMDD"): strTime = Format$(Now, "MMSS")
                            
                            sqlDoc = "Update INTERFACE003" & _
                                     "   set RSTVAL  = '" & strResult & "', REFVAL = ''" & _
                                     " where SPCNO   = '" & strBarno & "'" & _
                                     "   and EQPNUM  = '" & itemX.tag & "'" & _
                                     "   and TRANSDT = '" & strDate & "'" & _
                                     "   and TRANSTM = '" & strTime & "'"
                            AdoCn_Jet.Execute sqlDoc

                            sqlDoc = "insert into INTERFACE003(" & _
                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                     "    values( '" & strBarno & "', '" & itemX.text & "', '" & itemX.tag & "'," & _
                                     "            '" & strDate & "', '" & strTime & "'," & _
                                     "            '" & strResult & "', ''," & _
                                     "            '" & INS_CODE & "', '', '" & strSPnm & "', '" & strSPid & "')"
                            AdoCn_Jet.Execute sqlDoc
                            
                            '-- 서버결과등록
'                            If chkAuto.Value = "1" Then
'                                sqlDoc = "Update EXAM_TOC set EX_INRV = '" & Trim(strResult) & "',EX_INST = '2',EX_DATE = '" & Format$(Now, "YYYYMMDD") & "',EX_INEM='1271'" _
'                                       & " where RE_RCID ='" & strSPid & "' And IN_CODE='" & itemX.text & "'"
'
'                                AdoCn_ORACLE.Execute (sqlDoc)
'                                lblStatus.Caption = "저장 성공!!"
'                                AdoCn_ORACLE.Execute sqlDoc
'                            End If

                            Set itemX = Nothing
                        End If
                    End If
                    .Col = 7: .ForeColor = vbRed: .BackColor = vbCyan
                    .SetText 1, vRow, 1
                Next pDocount
            End With
        End If
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "L" Then      ' "L" Data Last
        comEQP.Output = ACK
        Debug.Print "L [HOST] " & ACK
        Patiant_Recevid = False                        ' 환자 번호  Flag
    Else
        comEQP.Output = ACK
    End If
                        
    Exit Sub
errDefine:

End Sub
'Private Sub ComReceive(ByRef RecData As String)
'    Dim intIdx1     As Integer, intIdx2     As Integer
'    Dim strTmp1     As String, strTmp2      As String
'    Dim intPos1     As Integer, intPos2     As Integer
'    Dim strDta()    As String, intCnt       As Integer
'    Dim strRec      As String, strbuff      As String
'
'    Debug.Print RecData
'    strRec = RecData
'    Print #1, strRec;
'    Call COM_INPUT(strRec)
'    Debug.Print strRec
'
'    For intIdx1 = 1 To Len(strRec)
'        strbuff = Mid$(strRec, intIdx1, 1)
'
'        Select Case Asc(strbuff)
'            Case 2 '-- STX
'                        f_strBuffer = strbuff
'            Case 3 '-- ETX
'                        If Mid$(f_strBuffer, 2, 2) = "R " Or Mid$(f_strBuffer, 2, 2) = "RH" Then
''                            Call RequestDefine(f_strBuffer, fChannel(), spdResult1)
'                        ElseIf Mid$(f_strBuffer, 2, 2) = "D " Or Mid$(f_strBuffer, 2, 2) = "DH" Then
'                            Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
'                        End If
'                        f_strBuffer = ""
'            Case Else
'                        f_strBuffer = f_strBuffer + strbuff
'        End Select
'    Next
'End Sub

'Private Sub ComReceive(ByRef RecData As String)
'
'    Dim strRec  As String, strBuff  As String
'    Dim strTmp  As String, intIdx   As Integer
'    Dim intPos1 As Integer, intPos2 As Integer
'
'    Dim strdata()   As String, intCnt   As Integer
'
'    Static OrgMsg As String
'    strRec = RecData ' StrConv(RecData, vbUnicode)
'    Debug.Print strRec
'
'    Print #1, strRec;
'
'    strTmp = strRec
'    Call COM_INPUT(strTmp)
'
'    For intIdx = 1 To Len(strRec)
'        strBuff = Mid$(strRec, intIdx, 1)
'        Select Case Asc(strBuff)
'            Case 2  '-- STX
'                    f_strBuffer = strBuff
'
'            Case 3  '-- ETX
'                    f_strBuffer = f_strBuffer + strBuff
'                    intCnt = 0
'                    strTmp = f_strBuffer
'                    Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
'            Case Else
'                    f_strBuffer = f_strBuffer + strBuff
'        End Select
'     Next
'End Sub

Private Sub ComReceive(ByRef RecData As Variant)
    
Dim strRec  As String, strBuff  As String
    Dim strTmp  As String, intIdx   As Integer
    Dim intPos0 As Integer, intPos1 As Integer, intPos2 As Integer
    
    Dim AGE As String
    Dim i As Integer
    Dim tt As Boolean
    Dim sHead       As String
    Dim sPInfo      As String
    Dim sRtypeId    As String  'Record Type ID(1)
    Dim sSNumber    As String  'Sequence Number(6)
    '---[Specimen ID]----------------------------
    Dim sSampleNo   As String  'Sample No(5)
    Dim sSampleId   As String  'Sample ID(13)
    Dim sSampleType As String  'Sample Type(1)
    Dim sRackId     As String  'Rack Id(5)
    Dim sPositionNo As String  'Position No(1)
    '--------------------------------------------
    Dim sSpecimenID As String  'Specimen ID(2)
    '---[Universal Test Id]----------------------
    Dim sAppCode    As String  'Application Code(3)
    Dim sIdc        As String  'Inc,Dec or Cir(3)
    '--------------------------------------------
    Dim sPriority   As String  'Priority(1)
    Dim sRDateTime  As String  'Requested/Ordered Date and Time
    Dim sSDateTime  As String  'Specimen Collection Date and Time(14)
    Dim sCEndTime   As String  'Collection End Time
    Dim sCvolume    As String  'Collection Volume
    Dim sCId        As String  'Collection Id
    Dim sACode      As String  'Action Code(1)
    Dim sDCode      As String  'Danger Code
    Dim sRcinfo     As String  'Relevant Clinical Information(7)
    Dim sDtSpeR     As String  'Date/Time Specimen Received
    Dim sSpeDesc    As String  'Specimen Descriptor(2)
    Dim sOrderPh    As String  'Ordering Physician
    Dim sPtNum      As String  'Physician's Telephone Number
    Dim sUserF1     As String  'User Field No1(6)
    Dim sUserF2     As String  'User Field No2(104)
    Dim sLaboF1     As String  'Laboratory Field No.1
    Dim sLaboF2     As String  'Laboratory Field No.2
    Dim sDtRr       As String  'Date/Time Result(14)
    Dim sIccs       As String  'Instrument Charge to Computer System
    Dim sIsId       As String  'Instrument Section ID
    Dim sReportT    As String  'Report Types(1)
    Dim ii As Integer
    Dim sTempid As String
    Dim Orderoutput As String
    Dim OutPutData As String
    Dim Testcd As String, sOrderLst As String
    Dim Loop_count As Integer, pDocount, pChnoCount As Integer
    Dim SEX As String
    Dim intldx As Integer
    Dim sStxCheck As Integer
    Dim sEtxCheck As Integer
    
    Static OrgMsg As String
    
    strRec = RecData
    Print #1, strRec;
    Call COM_INPUT(strRec)
'    Debug.Print strRec
    
    For intIdx = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx, 1)
        Select Case strBuff
            Case STX
                    sStxCheck = InStr(strBuff, STX)
            Case ETX
                    Debug.Print strTmp
                    sEtxCheck = InStr(strBuff, ETX)
                    If sStxCheck <> 0 And sEtxCheck <> 0 Then
                        Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
                        GoSub ClearReceiveData
                    End If
            Case ETB
                    If Mid(f_strBuffer, intIdx, 2) = vbCrLf Then
                        f_strBuffer = left(f_strBuffer, Len(f_strBuffer) - 2) 'Remove CR & LF
                    End If
                    cntCheckSum = cntCheckSum + 1
                    Call COM_OUTPUT(ACK)
'                        Debug.Print "[HOST] " & ACK
                    flgETB = True
            Case vbCr

            Case vbLf

            Case ENQ
                    Call COM_OUTPUT(ACK)
            Case ACK
                    Dim varTmp      As Variant
                    Dim IntRow      As Integer, intCol  As Integer
                    Dim strBarno    As String, strTest  As String
                    Dim strRack     As String, strCup   As String
                    Dim intCnt      As Integer
                    Dim itemX       As ListItem

                    With spdResult1
                        For IntRow = 1 To .maxrows
                            .Row = IntRow
                            .Col = 2
                            If .BackColor = vbWhite Then
                                sAppCode = ""
                                intCnt = 0
                                .GetText 3, IntRow, varTmp: strBarno = Trim$(varTmp)
                                .GetText 5, IntRow, varTmp: strRack = Trim$(varTmp)
                                .GetText 6, IntRow, varTmp: strCup = Trim$(varTmp)
'                                .GetText 1, intRow, varTmp
                                For intCol = 7 To .MaxCols - 15
                                    spdResult1.GetText intCol, 0, varTmp
                                    If Trim$(varTmp) = "" Then Exit For
                                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                    If Not itemX Is Nothing Then
                                        spdResult1.Col = intCol:    'spdResult1.Row = OrderCnt
                                        If spdResult1.BackColor = &HC6FEFF Then
                                            sAppCode = sAppCode + "^^^" & Trim(Val(itemX.tag)) & "/\"
                                            intCnt = intCnt + 1
                                        End If
                                    End If
                                    Set itemX = Nothing
                                Next intCol
                                If sAppCode <> "" Then
                                    .Row = IntRow
                                    .Col = 2: .BackColor = vbCyan
                                    .Col = 3: .BackColor = vbCyan
                                    .Col = 4: .BackColor = vbCyan
                                End If
                                Exit For
                            End If
                        Next IntRow
                    End With
                    If IntRow > spdResult1.maxrows Then
                        Exit Sub
                    Else
                        sHead = "H|\^&|||HOST^2|||||H7600^1|TSDWN^BATCH|P|1" + Chr(13)
                        sPInfo = "P|1" + Chr(13)
                        sRtypeId = "O"
                        sSNumber = "1"
                        sSampleNo = Format$(IntRow, "0000")
                        'sSampleId = Space(2) & Right(H7600.SID, 10)
                        sSampleType = "1"
                        sRackId = Format$(strRack, "0000")
    '                    sPositionNo = H7600.Position
                        sSpecimenID = "R1"
                        'sAppCode = ""
                        sIdc = ""
                        sPriority = "R" 'H7600.Priority
                        sRDateTime = ""
                        sSDateTime = Format(Now, "YYYYMMDDHHMMSS")
                        sCEndTime = ""
                        sCvolume = ""
                        sCId = ""
                        sACode = "N"
                        sDCode = ""
                        sRcinfo = ""
                        sDtSpeR = ""
                        sSpeDesc = ""
                        sOrderPh = ""
                        sPtNum = ""
                        sUserF1 = ""
                        sUserF2 = sSampleNo + "                          ^^^^"
                        sLaboF1 = ""
                        sLaboF2 = ""
                        sDtRr = ""
                        sIccs = ""
                        sIsId = ""
                        sReportT = "O"
                        sRcinfo = "^^"

                        HostOutput = sHead & sPInfo & _
                                       sRtypeId & Field_ & sSNumber & _
                                       Field_ & sSampleNo & Component_ & "             " & _
                                       Component_ & sSampleType & Component_ & sRackId & _
                                       Component_ & strCup & Field_ & sSpecimenID & _
                                       Field_ & left(sAppCode, Len(sAppCode) - 1) & _
                                       Field_ & sPriority & Field_ & sRDateTime & _
                                       Field_ & sSDateTime & Field_ & sCEndTime & _
                                       Field_ & sCvolume & Field_ & sCId & _
                                       Field_ & sACode & Field_ & sDCode & _
                                       Field_ & sRcinfo & Field_ & sDtSpeR & _
                                       Field_ & sSpeDesc & Field_ & sOrderPh & _
                                       Field_ & sPtNum & Field_ & sUserF1 & _
                                       Field_ & sUserF2 & Field_ & sLaboF1 & _
                                       Field_ & sLaboF2 & Field_ & sDtRr & _
                                       Field_ & sIccs & Field_ & sIsId & _
                                       Field_ & sReportT & Chr(13) & "L|1|N" & Chr(13)

                        SendCount = Int((Len(HostOutput) / 230)) + 1

                        For i = 1 To SendCount
                            SendData(SendCount - i + 1) = i & Mid(HostOutput, (i - 1) * 230 + 1, 230)
                            If i = SendCount Then
                               SendData(SendCount - i + 1) = SendData(SendCount - i + 1) & ETX
                            Else
                               SendData(SendCount - i + 1) = SendData(SendCount - i + 1) & ETB
                            End If
                            SendData(SendCount - i + 1) = STX & SendData(SendCount - i + 1) & MakeCS(SendData(SendCount - i + 1)) & Chr(13) & EOT
                        Next i
                        Call COM_OUTPUT(ENQ)
                        Call COM_OUTPUT(SendData(SendCount))
                        Debug.Print " T:" & ENQ & SendData(SendCount)
                        sAppCode = ""
                    End If
            Case NAK

            Case EOT
                    Call COM_OUTPUT(ACK)
                    GoSub ClearReceiveData
            Case Else
                    f_strBuffer = f_strBuffer + strBuff
        End Select
    Next
    
ClearReceiveData:
    ReceiveData = ""
    cntField_ = 0
    cntRepeat_ = 0
    cntComponent_ = 0
    cntEscape_ = 0
    cntSlash_ = 0
    f_strBuffer = ""
'    Return
     
End Sub
Private Sub ReceiveTheData(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)
    
    
    Dim sTemp      As String
    Dim Channel_No As String        ' 검사항목 번호 : Channel No
    Dim pGrid_Point As Integer
    Dim pDocount   As Integer
    Dim Loop_count As Integer
    Dim FunStr As String
    Dim Max_Arary_Cnt As Integer    ' 검사 항목수
    Dim sAdd As Integer, sPosition As Integer
    Dim itemX As ListItem
    Dim strRstval As String, strRefVal  As String
    Dim sqlDoc  As String
    Dim intCol As Integer
    Dim Gnum   As String
    Dim ii As Integer, jj As Integer, kk As Integer
    Dim Test_Cd() As String
    Dim Rev As Long
    Dim tmpTstCd As String
    Dim tmpMXD As Variant
    Dim sSeq, strTmp, varTmp, strBarno, strDate, strDate1, strTime As String
    Dim sCol As Integer
    Dim sDeCnt As Integer
    Dim Float_rate1 As String
    Dim Float_rate2 As String
    Dim Float_rate  As String
    Dim IntRow, intIdx As Integer
    Dim chrChk As Boolean
    Dim seqChk As Variant
    Dim chkGbn As Variant
    Dim strEqpCd As String
    
    On Error Resume Next
       
    CallForm = "frmInterface - Privete sub psDataDefine()"

    pDocount = 0
    Do While InStr(strdata, "|") > 0
        pDocount = pDocount + 1
        fTBA40FR(pDocount) = Text_Redefine(strdata, "|")
        strdata = Mid$(strdata, InStr(strdata, "|") + 1)   ' 구분자가 "|" 이다....
        If pDocount > 99 Then
            strdata = ""
            Exit Do
        End If
    Loop
    
    pGrid_Point = 0
    strTmp = ""
    
    If Mid$(fTBA40FR(1), 3, 1) = "H" Then          ' "H" Head Message Display
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
    ElseIf Mid$(fTBA40FR(1), 3, 1) = "P" Then      ' "P" Patiant Information Data Process
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
    ElseIf Mid$(fTBA40FR(1), 3, 1) = "C" Then
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
    ElseIf Mid$(fTBA40FR(1), 3, 1) = "Q" Then      ' "Q" Patiant Information Data Process
        comEQP.Output = ACK
    ElseIf Mid$(fTBA40FR(1), 3, 1) = "O" Then      ' "O" Order Data Process
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
        Patiant_Recevid = False                        ' 환자 번호  Flag
        strBarno = Val(Text_Redefine(fTBA40FR(4), "^"))  '' 환자번호  "5450^0^57"
        '-------------------------------------------<<< 해당검사결과와 해당환자를 찿는다.       >>>----------
        If optSeq.Value = 1 Then
            sCol = 7
        Else
            sCol = 3
        End If
        pGrid_Point = SeqSearch(brspread, strBarno, sCol)
        Patiant_Recevid = (pGrid_Point > 0)
    ElseIf Mid$(fTBA40FR(1), 3, 1) = "R" Then      ' "R" Result Data Process
        Dim ssChannel() As String
        comEQP.Output = ACK
        
        If Patiant_Recevid = True Then
            ssChannel = Split(fTBA40FR(3), "^")
            If UBound(ssChannel) > 3 Then
                fTBA40FR(3) = ssChannel(3)
                Channel_No = fTBA40FR(3)
            Else
                Channel_No = 0
            End If
'            fTBA40FR(3) = fclsFunc.Text_Change(fTBA40FR(3), "^", "")    ' channel
'            Channel_No = Val(fTBA40FR(3) / 10)                                   ' channel
            '-------------------------------------------<<< 해당검사결과를 찿는다.       >>>----------
            Max_Arary_Cnt = brspread.MaxCols - 6   ' 앞에서부터 5까지는 환자 정보 이기때문에.... -6를 한다.
                                                   ' 해당 배열은  brItem(),brChannel() 이다.
            With brspread
                '----------------------------------------------<<<<<<<<<,  세부검사항목을 찿는다.  >>>>>>>----------

                For pDocount = 1 To Max_Arary_Cnt
                    .Col = pDocount + 6
                    If Channel_No > 0 And Channel_No = Val(brChannel(pDocount)) Then          ' 검사결과가 있으면...
                        If Trim(fTBA40FR(4)) <> "" Then
                            fTBA40FR(4) = Text_Change(fTBA40FR(4), ">", "")
                            fTBA40FR(4) = Text_Change(fTBA40FR(4), "<", "")

                            If InStr(fTBA40FR(4), "^") > 0 Then
                                .text = Trim(Mid$(fTBA40FR(4), InStr(fTBA40FR(4), "^") + 1))
                            Else
                                .text = Trim(fTBA40FR(4))
                            End If
                        Else
                            .text = ""
                        End If

                    End If

                Next pDocount
            End With
        End If
    
        IntRow = 0
        With spdResult1
            sCol = 8
            pGrid_Point = SeqNullSearch(spdResult1, sSeq, sCol)
            
            
    
            .GetText 2, pGrid_Point, varTmp:   strDate1 = Trim$(varTmp)
            .GetText 3, pGrid_Point, varTmp:   strBarno = Trim$(varTmp)
            .GetText 4, pGrid_Point, varTmp:   pName = Trim$(varTmp)
            .GetText 5, pGrid_Point, varTmp:   pNo = Trim$(varTmp)
            chkGbn = Split(pName, ":")

            .GetText 2, pGrid_Point, varTmp ':   strBarno = Trim$(varTmp)

            If pGrid_Point > 0 Then
                Set mAdoRs = f_subSet_WorkList_Barcode(strBarno, Mid(pName, 1, 2))
                For intCol = 8 To .MaxCols
                    strRstval = ""
                    .GetText intCol, 0, varTmp
                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                    If Not itemX Is Nothing Then
                        For intIdx = 1 To .MaxCols
                            If Len(fELEC1010(9)) > 0 Then
                                strEqpCd = ""
                                Do Until mAdoRs.EOF
                                    If Mid(pName, 1, 2) = "검진" Then
                                        If Mid(itemX.text, InStr(itemX.text, ",") + 1) = Trim(mAdoRs.Fields("EDPSCODE")) Then
                                            strEqpCd = Trim(mAdoRs.Fields("EDPSCODE"))
                                            Exit Do
                                        End If
                                    Else
                                        If Mid(itemX.text, InStr(itemX.text, ",") + 1) = Trim(mAdoRs.Fields("MAP2SEQNO")) Then
                                            strEqpCd = Trim(mAdoRs.Fields("MAP2SEQNO"))
                                            Exit Do
                                        End If
                                    End If
                                    mAdoRs.MoveNext
                                Loop
                                mAdoRs.MoveFirst
                                
                                If Trim(strEqpCd) <> "" Then
                                    fELEC1010_2 = Split(Trim(fELEC1010(intIdx + 10)), "=")
                                    Channel_No = Trim(fELEC1010_2(0))
                                    If UCase(Channel_No) = UCase(itemX.tag) Then
                                        fELEC1010_3 = Split(Trim(fELEC1010_2(1)), " ")
                                        If UCase(Channel_No) = "CL" Then
                                            strRstval = ""
                                        Else
                                            strRstval = Trim(Mid(fELEC1010_2(1), 3, 3))
                                        End If
                                         strDate1 = Format$(Now, "YYYYMMDD"):     strTime = Format$(Now, "MMSS")
                                        .SetText intCol, pGrid_Point, strRstval
                                        .Col = intCol:  .Row = pGrid_Point
                                                        .ForeColor = IIf(Trim$(strRefVal) <> "", vbRed, vbBlack)
    
                                        sqlDoc = "Update INTERFACE003" & _
                                                 "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                                                 " where SPCNO   = '" & strBarno & "'" & _
                                                 "   and EQPNUM  = '" & itemX.tag & "'" & _
                                                 "   and TRANSDT = '" & strDate1 & "'" & _
                                                 "   and TRANSTM = '" & strTime & "'"
                                        AdoCn_Jet.Execute sqlDoc
    
                                        If cboChk.ListIndex = 0 Then
                                            sqlDoc = "insert into INTERFACE003(" & _
                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                     "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
                                                     "            '" & strDate1 & "', '" & strTime & "'," & _
                                                     "            '" & strRstval & "', '" & strRefVal & "'," & _
                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                        Else
                                            sqlDoc = "insert into INTERFACE003(" & _
                                                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                                     "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.tag & "'," & _
                                                     "            '" & strDate1 & "', '" & strTime & "'," & _
                                                     "            '" & strRstval & "', '" & strRefVal & "'," & _
                                                     "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                        End If
    
                                        AdoCn_Jet.Execute sqlDoc
    
                                        If chkAuto.Value = "1" Then
                                            If Mid(pName, 1, 2) = "검진" Then
                                                sqlDoc = "Update MDCK..GUMJIN_INTERFACE" & _
                                                         "   set RESULT = '" & strRstval & "'," & _
                                                         "       ACT_RETURN_DATE = '" & strDate1 & "'" & _
                                                         " where PER_GUMJIN_DATE = '" & strDate & "'" & _
                                                         "   and PER_GUM_NUM = " & pNo & "" & _
                                                         "   and EDPSCODE = '" & strEqpCd & "'"
                                            Else
                                                sqlDoc = "Update MEDICOM..jun370_resulttb" _
                                                        & "   Set Result = '" & strRstval & "', status='1'" _
                                                        & " Where WaitSeqNo = '" & pNo & "'" _
                                                        & "   and map2seqno = '" & strEqpCd & "'"
                                            End If
                                            AdoCn_SQL.Execute sqlDoc
                                        End If
    
                                        spdResult1.Row = pGrid_Point
                                        spdResult1.Col = 2
                                        spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 3
                                        spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 4
                                        spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 5
                                        spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 6
                                        spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 7
                                        spdResult1.BackColor = vbCyan
                                        spdResult1.Col = 1: spdResult1.Value = 0
                                        Exit For
    
                                    End If
                                End If
                            End If
                        Next intIdx
                    End If
                    Set itemX = Nothing
                Next
            End If
        End With
    
        Set mAdoRs = Nothing
    
    ElseIf Mid$(fTBA40FR(1), 3, 1) = "L" Then      ' "L" Data Last
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
        Patiant_Recevid = False                        ' 환자 번호  Flag
    End If
    
    Exit Sub

ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .maxrows
            .Row = sCnt
            .Col = brCol
            If Trim(.text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Function f_funAdd_Server(ByVal strBarno As String, ByVal strTestcd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_Server = False
    
    strTmp = strTestcd: intPos = InStr(strTmp, ",")
    Do While intPos > 0
        blnFlag = False
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = Mid$(strTmp, 1, intPos - 1) Then
                blnFlag = True
                strTmp = Mid$(strTmp, 1, intPos - 1)
                Exit Do
            End If
        Next
        
        strTmp = Mid$(strTmp, intPos + 1)
        intPos = InStr(strTmp, ",")
    Loop
    
    If Not blnFlag Then
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = strTmp Then blnFlag = True: Exit For
        Next
    End If
    
    If blnFlag Then
        ReDim Preserve strSampleno(1 To 1) As String
        ReDim Preserve strOrdcd(1 To 1) As String
        ReDim Preserve strRstval(1 To 1) As String
        ReDim Preserve strTmp1(1 To 1) As String
        ReDim Preserve strTmp2(1 To 1) As String
        
        strSampleno(1) = strBarno
        strOrdcd(1) = strTmp
        strRstval(1) = strTestval
        strTmp2(1) = INS_CODE
        
        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "0" Then
            f_funAdd_Server = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
    End If
                                
End Function

Private Function f_funAdd_QcServer(ByVal strBarno As String, ByVal strTestcd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_QcServer = False
    
    strTmp = strTestcd: intPos = InStr(strTmp, ",")
    Do While intPos > 0
        blnFlag = False
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = Mid$(strTmp, 1, intPos - 1) Then
                blnFlag = True
                strTmp = Mid$(strTmp, 1, intPos - 1)
                Exit Do
            End If
        Next
        
        strTmp = Mid$(strTmp, intPos + 1)
        intPos = InStr(strTmp, ",")
    Loop
    
    If Not blnFlag Then
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = strTmp Then blnFlag = True: Exit For
        Next
    End If
    
    If blnFlag Then
        ReDim Preserve strSampleno(1 To 1) As String
        ReDim Preserve strOrdcd(1 To 1) As String
        ReDim Preserve strRstval(1 To 1) As String
        ReDim Preserve strTmp1(1 To 1) As String
        ReDim Preserve strTmp2(1 To 1) As String
        
        strSampleno(1) = strBarno
        strOrdcd(1) = strTmp
        strRstval(1) = strTestval
        strTmp2(1) = INS_CODE
        
        Call sl_online_pc_98&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "" Then
            f_funAdd_QcServer = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
    End If
                                
End Function

Public Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
End Function

Public Function Text_Change(FSend_Str As String, FCheck_Char As String, FChange_Char As String) As String
Dim Pos_point As Integer
    Do
        Pos_point = InStr(FSend_Str, FCheck_Char)
        If Pos_point < 1 Then
            Exit Do
        ElseIf Pos_point = 1 Then
            FSend_Str = FChange_Char + Mid$(FSend_Str, 2)
        Else
            FSend_Str = Mid$(FSend_Str, 1, Pos_point - 1) + FChange_Char + Mid$(FSend_Str, Pos_point + 1)
        End If
    Loop
    Text_Change = FSend_Str
    
End Function

Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brspread
        If optSeq.Value = False Then
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Val(.text) = brSeq Then
                    SeqSearch = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        Else
            For sCnt = 1 To .maxrows
                .Row = sCnt
                .Col = brCol
                If Trim(.text) = brSeq Then
                    SeqSearch = sCnt 'brSeq
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            Next sCnt
        End If
    End With

End Function

Private Sub Command1_Click()
'
'    Dim Arr()   As Byte
'    Dim strTmp  As String
'
'
'    strTmp = AS400DOWNF(INS_CODE, "20111230", "20111230")
'
'    If strTmp = "44" Then
'        '-- 장비코드 오류
'    ElseIf strTmp = "48" Then
'        '시작일오류
'    ElseIf strTmp = "52" Then
'        '종료일오류
'    End If
'    'strTmp = AS400UPF(INS_CODE, "20111230")
   
    ReceiveData = ENQ
    ReceiveData = ReceiveData & "1H|\^&||||||||||P||" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "2P|1|||||||||||||||||||||||||||||||||" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "3O|1|0001|1^0001^1^^SAMPLE^NORMAL|ALL|R|20030722194828|||||X||||||||||||||O|||||" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "4R|1|^^^250^^0|22.30|ng/ml|25.00^72.00|L||F|||20030722195530|20030722200528|" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "5C|1|I|48^Below expected value range|I  DA"
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "6R|2|^^^10^^0|0.058|mIU/l|0.270^4.20|L||F|||20030722195448|20030722201310|" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "7C|1|I|48^Below expected value range|I" & vbCr
    ReceiveData = ReceiveData & "05" & vbCrLf
    ReceiveData = ReceiveData & "0L|1" & vbCr
    ReceiveData = ReceiveData & "04" & vbCrLf
    ReceiveData = ReceiveData & ""
 
    ReceiveData = ENQ
    ReceiveData = ReceiveData & "1H|\^&|||cmc|||||||P||20120313131558" & vbCr
    ReceiveData = ReceiveData & "P|1||1020135856||^|||||||||^|^^^|^|^||||||||" & vbCr
    ReceiveData = ReceiveData & "O|1||Sample #^250|^^^|||||||||||Whole Blood^||||||||||F" & vbCr
    ReceiveData = ReceiveData & "C|1|I|1196^A default Hct value was used|I" & vbCr
    ReceiveData = ReceiveData & "L|1|N" & vbCr
    ReceiveData = ReceiveData & ""

 
    ReceiveData = ENQ
    ReceiveData = ReceiveData & "1H|\^&|||cmc|||||||P||20120313131558" & vbCr
    ReceiveData = ReceiveData & "P|1" & vbCr
    ReceiveData = ReceiveData & "O|1|Q24IL0030^^B09MP^1||^^^Ribo P|R|20111115151541" & vbCr
'    ReceiveData = ReceiveData & "O|1|O24IK0010^^B09MP^1||^^^Ribo P|R|20111115151541" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^Ribo P^QUANT|<0.2|AI||<||F|||20120206155741|20120206164019" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^Ribo P^QUAL|Negative|||||F|||20120206155741|20120206164019" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^Jo-1^QUANT|<0.2|AI||<||F|||20120206155741|20120206164019" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^Jo-1^QUAL|Negative|||||F|||20120206155741|20120206164019" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^Sm^QUANT|>8.0|AI||>||F|||20120202131305|20120202135533" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^Sm^QUAL|POSITIVE|||||F|||20120202131305|20120202135533" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^RNP 68^QUANT|3.1|AI||||F|||20120202131305|20120202135533" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^RNP 68^QUAL|POSITIVE|||||F|||20120202131305|20120202135533" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^SS-B^QUANT|1.1|AI||||F|||20120202131341|20120202135611" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^SS-B^QUAL|POSITIVE|||||F|||20120202131341|20120202135611" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^SS-A60^QUANT|>8.0|AI||>||F|||20120202131341|20120202135611" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^SS-A60^QUAL|POSITIVE|||||F|||20120202131341|20120202135611" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^Cent B^QUANT|<0.2|AI||<||F|||20120206122302|20120206130533" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^Cent B^QUAL|Negative|||||F|||20120206122302|20120206130533" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^dsDNA^QUANT|<0.3|AI||<||F|||20120206122302|20120206130533" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^dsDNA^QUAL|Negative|||||F|||20120206122302|20120206130533" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^GBM^QUANT|<0.2|AI||<^GBM: -43126 No QC||F|||20120329111934|20120329120237" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^GBM^QUAL|Negative|||^GBM: -43126 No QC||F|||20120329111934|20120329CC" & vbCr
    
    ReceiveData = ReceiveData & "L|1|N" & vbCr
    ReceiveData = ReceiveData & "04" & vbCr
    ReceiveData = ReceiveData & ""



    ReceiveData = ENQ
    ReceiveData = ReceiveData & "1H|\^&|||neodin|||||||P||20120403143925" & vbCr
    ReceiveData = ReceiveData & "P|1" & vbCr
    ReceiveData = ReceiveData & "O|1|02093110051^^B0AEI^4||^^^PR3|R|20120403104521" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^PR3^QUANT|<0.2|AI||<||F|||20120403120512|20120403124755" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^PR3^QUAL|Negative|||||F|||20120403120512|20120403124755" & vbCr
    ReceiveData = ReceiveData & "L|1|N" & vbCr
    ReceiveData = ReceiveData & "A2" & vbCr

    ReceiveData = ReceiveData & "2H|\^&|||neodin|||||||P||20120403143925" & vbCr
    ReceiveData = ReceiveData & "P|1" & vbCr
    ReceiveData = ReceiveData & "O|1|02093110051^^B0AEI^4||^^^MPO|R|20120403104521" & vbCr
    ReceiveData = ReceiveData & "R|1|^^^MPO^QUANT|<0.2|AI||<||F|||20120403120512|20120403124755" & vbCr
    ReceiveData = ReceiveData & "R|2|^^^MPO^QUAL|Negative|||||F|||20120403120512|20120403124755" & vbCr
    ReceiveData = ReceiveData & "L|1|N" & vbCr
    ReceiveData = ReceiveData & "E8" & vbCr
    ReceiveData = ReceiveData & "" & vbCr


'    ReceiveData = ENQ
'    ReceiveData = ReceiveData & "1H|\^&|||cmc|||||||P||20120313123507" & vbCr
''    ReceiveData = ReceiveData & "Q|1|^208051007710||ALL||||||||O" & vbCr
''    ReceiveData = ReceiveData & "Q|1|^208044027410||ALL||||||||O" & vbCr
'    ReceiveData = ReceiveData & "Q|1|^208013006012||ALL||||||||O" & vbCr
'    ReceiveData = ReceiveData & "L|1|N" & vbCr
'    ReceiveData = ReceiveData & "84" & vbCr
'    ReceiveData = ReceiveData & "" & vbCr
'''
    'ReceiveData = ACK
 
'    Call EditRcvData
    
    Call comEQP_OnComm
    
End Sub

Private Sub Form_Activate()

    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
        
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    
    Call f_subSet_ComCharacter  ' 통신문자
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun                 ' 실행
    
    mskRstDate.text = Now
    dtpStartDt.Value = (Now - 30)
    dtpStopDt.Value = Now
    mskOrdtime.text = Format$(Now, "HHMM")
    
    Open App.Path + "\Log\" + REG_INSNAME + "_" + Format(Now, "YYYYMMDD") + ".log" For Append As #1


    Print #1, Chr(13) + Chr(10);
    
    Open App.Path + "\ErrorLog\" + REG_INSNAME + "_" + Format(Now, "YYYYMMDD") + ".sql" For Append As #2

    Print #2, Chr(13) + Chr(10);
   
    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    tabWork.Tab = 0
    Or_Seq = 1
    IntRow = 0
    chkEnq = 0
    cboChk.ListIndex = 0
    cboChk1.ListIndex = 0
    
    gspdResultRow = 0
    
    
    COM_MODE = "1"
    
    
    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    
End Sub


Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""

    GetSetup = False

    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "URL", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtURL.text = Trim(db_tmp)

    db_tmp = ""
    Call GetPrivateProfileString("CONFIG", "INSTCD", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtInstCd.text = Trim(db_tmp)

    GetSetup = True

End Function

Private Sub f_subGet_Setting()
    
    Dim objComSetting As clsCommon
    Dim Baudratio As String
    Dim Paritybit As String
    Dim Databit As String
    Dim Stopbit As String
    
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subGet_Setting()"
    Set objComSetting = New clsCommon
    
    With objComSetting
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_EqpProperty(INS_CODE)
    End With
    Set objComSetting = Nothing
    
    If mAdoRs Is Nothing Then
        IS_SET = False
        MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
        Exit Sub
    Else
        If mAdoRs.EOF Then
            IS_SET = False
            MsgBox INS_CODE & " 에 대한 장비 통신 구성이 없습니다. 통신 설정후 다시 시도 하십시오.", vbExclamation
            Set mAdoRs = Nothing
            Exit Sub
        Else
            IS_SET = True
            Baudratio = Trim(mAdoRs.Fields("COM_SPEED") & "")
            Paritybit = Trim(mAdoRs.Fields("COM_PARITYBIT") & "")
            Databit = Trim(mAdoRs.Fields("COM_DATABIT") & "")
            Stopbit = Trim(mAdoRs.Fields("COM_STOPBIT") & "")
            
            With comEQP
                .CommPort = Trim(mAdoRs.Fields("COM_PORT") & "")
                .Handshaking = Trim(mAdoRs.Fields("COM_HANDSHAK") & "")
'                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
'                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
'                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
'                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
'                .InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
'                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
'                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
'                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")
                .RThreshold = Trim(mAdoRs.Fields("COM_RTH") & "")
                .SThreshold = Trim(mAdoRs.Fields("COM_STH") & "")
                .Settings = Baudratio & "," & Paritybit & "," & Databit & "," & Stopbit
            End With
            Call Del_OldData
        End If
    End If
    
    Set mAdoRs = Nothing
Exit Sub

ErrRoutine:
    Set objComSetting = Nothing
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call cmdStop
    Set Result = Nothing
    
    Close #1
    Close #2
End Sub

Private Sub FrameError_Click()
    txtResult.Visible = True
    'List1.Visible = False
End Sub

Private Sub imgPort_DblClick()
    
    If lvwCuData.Visible Then
        lvwCuData.Visible = False
    Else
        lvwCuData.Visible = True
        lvwCuData.ZOrder 0
    End If
    
End Sub

Private Sub imgReceive_DblClick()

    If pnlCom2.Visible = True Then
        pnlCom2.Visible = False
    Else
        pnlCom2.Visible = True
        pnlCom2.ZOrder 0
    End If
    
End Sub

Private Sub imgSend_DblClick()
    
    If pnlCom.Visible = True Then
        pnlCom.Visible = False
    Else
        pnlCom.Visible = True
        pnlCom.ZOrder 0
    
        tmrWorking.interval = 20000
        tmrWorking.Enabled = True
    End If

End Sub


Private Sub Label6_DblClick()
    If Frame4.Visible = False Then
        Frame4.Visible = True
    Else
        Frame4.Visible = False
    End If
End Sub

Private Sub Label9_DblClick()

    If COM_MODE = "1" Then
        COM_MODE = "0"
        ShowMessage "인터페이스 내용을 화면에 출력하지 않습니다."
    Else
        COM_MODE = "1"
        ShowMessage "인터페이스 내용을 화면에 출력합니다."
    End If
End Sub

Private Sub tabWork_Click(PreviousTab As Integer)
    cboRstgbn(1).ListIndex = 2
'    spdResult2.maxrows = 0
'    Call dtpRsltDay_Change
'    If PreviousTab = 0 Then
'        cmdAppend(1).Visible = False
'    Else
'        cmdAppend(1).Visible = True
'    End If
End Sub

Private Sub tmrDummy_Timer()
Dim strSvrcData As String
'   - service : CC_SYSDATE_S
'     input   : N/A
'     output  : S_DATETIME1  /* system time (yyyy-mm-dd hh24:mi:ss)           (s) */
     
'    tmrDummy.interval = 1000 '65000
'    tmrDummy.Enabled = True
    
    lngDummyTime = lngDummyTime + 1
    
    If lngDummyTime >= 20 Then
        strSvrcData = getSvrcInfo("CC_SYSDATE_S", "")
        lngDummyTime = 0
    End If
     
End Sub

Private Sub tmrOk_Timer()
    'fraOK.Visible = False
    'tmrOk.Enabled = False
End Sub

Private Sub tmrOrder_Timer()
    Dim blnAllSend As Boolean
    
    blnAllSend = True
    
    With spdResult1
        For IntRow = 1 To .maxrows
            .Row = IntRow
            .Col = 1
            If .Value = "1" Then
                Call cmdOrder_Click
                blnAllSend = False
            End If
        Next
    End With
    
    tmrOrder.Enabled = False

'    If blnAllSend = True Then
'        With spdResult1
'            For intRow = 1 To .maxrows
'                .Row = intRow
'                .Col = 1
'                .Value = "1"
'            Next
'        End With
'    End If
    
End Sub

Private Sub tmrWorking_Timer()
    pnlCom.Visible = False
End Sub

'Private Sub mskOrdDate_GotFocus()
'
'    With mskOrdDate
'        .SelStart = 8
'        .SelLength = Len(.text)
'    End With
'
'End Sub


'Private Sub mskOrdDate_KeyPress(KeyAscii As Integer)
'
'    If Not KeyAscii = vbKeyBack Then mskOrdDate.SelLength = 1
'
'End Sub


'Private Sub mskRstDate_GotFocus()
'
'    With mskRstDate
'        .SelStart = 0
'        .SelLength = Len(.text) + 2
'    End With '
'
'End Sub
'
'
'Private Sub mskRstDate_KeyPress(KeyAscii As Integer)
'
'    If Not KeyAscii = vbKeyBack Then mskRstDate.SelLength = 1
'
'End Sub

Private Sub Order_Ready(ByVal ACK As String)

    Static msgIndex As Long
    
    Select Case ACK
        Case Chr(COM_ENQ)
            msgIndex = 1
        Case Chr(COM_ACK)
            msgIndex = msgIndex + 1
        Case Chr(COM_NACK)
            msgIndex = msgIndex
        Case Chr(COM_EOT)
            msgIndex = 7
            Set Order = Nothing
        Case Else
        
    End Select
    
    Select Case msgIndex
        Case 1
            Call COM_OUTPUT(Order.MSG_ENQ)
        Case 2
            Call COM_OUTPUT(Order.MSG_HEADER)
        Case 3
            Call COM_OUTPUT(Order.MSG_PATIENT)
        Case 4
            Call COM_OUTPUT(Order.MSG_ORDER)
        Case 5
            Call COM_OUTPUT(Order.MSG_TERMINATION)
        Case 6
            Call COM_OUTPUT(Order.MSG_EOT)
        Case Else
    End Select
    
End Sub

'Private Sub spdRstview_Click(ByVal Col As Long, ByVal Row As Long)
'Dim iCnt, rCnt As Integer
'Dim intCol, intRow As Integer
'Dim tCol As Integer
'Dim iresult As String
''
'' 결과 시작 Position
''
'Const sResultPos As Integer = 8
'    With spdRstview
'        For iCnt = 2 To .MaxCols Step 2
'            For rCnt = 1 To .maxrows
'                .Row = rCnt: .Col = iCnt
'                iresult = Trim(.text)
'
'                With spdResult1
'                    .Row = gspdResultRow:  .Col = sResultPos + tCol
'                    If Len(Trim(iresult)) <> 0 Then
'                        .text = iresult
'                    End If
'                    DoEvents
'                End With
'                tCol = tCol + 1
'
'            Next rCnt
'            rCnt = 0
'        Next iCnt
'    End With
'End Sub

'Private Sub spdRstview_EnterRow(ByVal Row As Long, ByVal RowIsLast As Long)
'    Call spdRstview_Click(Row, RowIsLast)
'End Sub

'
'
'
'
'Private Sub spdRstview_KeyPress(KeyAscii As Integer)
'
'Dim iCnt, rCnt As Integer
'Dim intCol, intRow As Integer
'Dim tCol As Integer
'Dim iresult As String
'
''
'' 결과 시작 Position
''
'Const sResultPos As Integer = 8
'
'    ' 처방 존재 유무 확인..
'    With spdRstview
'        .Row = .ActiveRow: .Col = .ActiveCol
'        If .BackColor <> &HC6FEFF And Len(.text) >= 1 Then
'            .text = ""
'            MsgBox "▒ OCS/EMR의 검사 처방이 없는 항목 입니다.." & Space(5), vbOKOnly + vbInformation, App.Title
'            spdRstview.SetFocus
'            Exit Sub
'        End If
'    End With
'
'    ' Enter Key 유무..
'    If KeyAscii = vbKeyReturn Then
'
'        If gspdResultRow < 1 Then
'            With spdRstview
'                .Row = .ActiveRow:  .Col = .ActiveCol
'                .text = ""
'            End With
'
'            MsgBox "▒ 수정을 원하는 검사 Sample을 선택 후 수정 하십시요.." & Space(5), vbOKOnly + vbInformation, App.Title
'            Exit Sub
'        End If
'
'        ' 수정된 결과 본 Spread로 옮기기..
'        With spdRstview
'            For iCnt = 2 To .MaxCols Step 2
'                For rCnt = 1 To .maxrows
'                    .Row = rCnt: .Col = iCnt
'                    iresult = .text
'
'                    With spdResult1
'                        .Row = gspdResultRow:  .Col = sResultPos + tCol
'                        If Len(Trim(iresult)) <> 0 Then
'                            .text = iresult
'                        End If
'                    End With
'                    tCol = tCol + 1
'                Next rCnt
'            Next iCnt
'        End With
'    End If
'
'End Sub
'
'
'Private Sub spdRstview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Dim objResult As clsResult
'   Dim lngCol As Long
'
'   If gspdResultRow = 0 Then Exit Sub
'
'   If 2280 >= X And X >= 1410 Then
'      lngCol = 2
'   ElseIf 4125 >= X And X >= 3210 Then
'      lngCol = 4
'   ElseIf 5055 >= X And X >= 5955 Then
'      lngCol = 8
'   ElseIf 6885 >= X And X >= 7755 Then
'      lngCol = 8
'   Else
'      lngCol = 9
'   End If
'
'   If Y < 330 Then Exit Sub
'
'   Select Case lngCol
'      Case 2, 4, 6, 8
'        spdRstview_TextTipFetch lngCol, gspdResultRow, 1, 6500, "", True
'      Case Else
'        Exit Sub
'   End Select
'
'End Sub


Private Sub spdRstview_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    
    Dim pDate, pPtnm, pPtno, pSex, pPos As String
    
    With spdResult1
        .Row = gspdResultRow
        .Col = 2: pDate = .text
        .Col = 4: pPtnm = .text
        .Col = 5: pSex = .text
        .Col = 6: pPtno = .text
        .Col = 7: pPos = .text
    End With
            
'    Debug.Print pDate, pPtnm, pPtno, pSex, pPos
            
    With spdRstview
        .Row = Row
        .Col = Col
         MultiLine = 1
         TipWidth = 3000
         .SetTextTipAppearance "굴림체", 9, False, False, &HEEFDF2, vbBlack
         .TextTip = TextTipFloating
         
    
         .SetTextTipAppearance "굴림체", 9, False, False, &HEEFDF2, vbBlue
         
         TipText = "" & vbNewLine & _
                   "   ▒ 처방일자 ; " & pDate & vbNewLine & _
                   "   ▒ 환 자 명 ; " & pPtnm & vbNewLine & _
                   "   ▒ 병록번호 ; " & pPtno & vbNewLine & _
                   "   ▒ 성    별 ; " & pSex & vbNewLine & vbNewLine & _
                   "   ▒ 검사 POS ; " & pPos & vbNewLine
                   
         ShowTip = True
       
    End With
End Sub


Private Sub spdRstview_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
    
    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 이전 환자", "-", " ▒ 다음 환자")

    Select Case lMenuChosen
        Case 1
            With spdResult1
                Col = .ActiveCol
                Row = .ActiveRow
            End With
            
            If gspdResultRow >= 1 Then
                Call spdResult1_Click(Col, gspdResultRow - 1)
            ElseIf gspdResultRow = 0 Then
                MsgBox "▒ 처음 자료입니다." & Space(5), vbOKOnly + vbInformation, App.Title
            Else: Exit Sub
            End If
            
        Case 3
            With spdResult1
                Col = .ActiveCol
                Row = .ActiveRow
            End With
            
            If gspdResultRow < spdResult1.maxrows Then
                Call spdResult1_Click(Col, gspdResultRow + 1)
            ElseIf gspdResultRow = spdResult1.maxrows Then
                MsgBox "▒ 마지막 자료입니다." & Space(5), vbOKOnly + vbInformation, App.Title
            Else: Exit Sub
    
            End If
    
    End Select


End Sub

Private Sub cmdNext_Click()
Dim Col, Row As Integer
    
    With spdResult1
        Col = .ActiveCol
        Row = .ActiveRow
    End With
    
    If gspdResultRow < spdResult1.maxrows Then
        Call spdResult1_Click(Col, gspdResultRow + 1)
    ElseIf gspdResultRow = spdResult1.maxrows Then
        MsgBox "▒ 마지막 자료입니다." & Space(5), vbOKOnly + vbInformation, App.Title
    Else: Exit Sub
    
    End If
    
End Sub

Private Sub cmdPrevious_Click()
Dim Col, Row As Integer
    With spdResult1
        Col = .ActiveCol
        Row = .ActiveRow
    End With
    
    If gspdResultRow >= 1 Then
        Call spdResult1_Click(Col, gspdResultRow - 1)
    ElseIf gspdResultRow = 0 Then
        MsgBox "▒ 처음 자료입니다." & Space(5), vbOKOnly + vbInformation, App.Title
    Else: Exit Sub
    End If
End Sub

Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol1 As Integer
    Dim intCol2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
    Dim iCnt    As Integer
    Dim strResult   As String
    
    If Row = 0 Then
        gspdResultRow = 0:        Exit Sub
    Else
        gspdResultRow = Row
    End If
    
    intCol1 = 11
    intCol2 = 2
    intRow1 = 1
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstview.Row = intRow1
            spdRstview.Col = intCol2
            spdRstview.BackColor = vbWhite
            spdRstview.ForeColor = vbBlack
            
            If .BackColor = &HC6FEFF Then
                spdRstview.BackColor = &HC6FEFF
            Else
                spdRstview.BackColor = &H80000005
            End If
                        
            strResult = .text
            
            'If InStr(strResult, "/") > 0 Then
                spdRstview.Col = intCol2
                spdRstview.text = strResult
'                spdRstview.Col = intCol2 + 1
'                spdRstview.text = mGetP(strResult, 2, "/")
'            Else
'                spdRstview.Col = intCol2
'                spdRstview.text = strResult
'                spdRstview.Col = intCol2 + 1
'                spdRstview.text = strResult
            'End If
            
            intRow1 = intRow1 + 1
            intCol1 = intCol1 + 1
            
            If intRow1 > spdRstview.maxrows Then
                intRow1 = 1
                intCol2 = intCol2 + 2
            End If

        Next
    End With
    
    

End Sub

'
' END
'

Private Sub spdResult1_KeyPress(KeyAscii As Integer)
'
'    Dim aROW    As Integer, aCOL   As Integer
'    Dim varChk  As Variant, varBar As Variant, varNum As Variant
'    Dim iRow, iCol   As Integer, iCnt   As Integer
'    Dim varVal  As Variant
'    Dim varOldVal As Variant
'
'
'    If KeyAscii = vbKeyReturn Then
'        With spdResult1
'            aCOL = .ActiveCol
'            aROW = .ActiveRow
'            varVal = .text
'            If aCOL = .MaxCols Then
'
'                If IsNumeric(varVal) Then
'                    For iCol = 8 To .MaxCols - 1
'                        .Col = iCol
'                        varOldVal = .text
'                        If IsNumeric(varOldVal) Then
'                            .text = Round((.text * 58) / (100 - varVal))
'                            SendKeys "{TAB}"
'                        End If
'                    Next
'                Else
'                    MsgBox "숫자만 입력이 가능합니다."
'                End If
'            End If
'        End With
'    End If
    

    Dim aRow    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    Dim intCol As Integer
    Dim varTmp
    Dim itemX   As ListItem
    
    'Debug.Print Col & NewCol & Row & NewRow
       
    If KeyAscii = vbKeyReturn Then
        With spdResult1
            .Row = .ActiveRow: aRow = .ActiveRow
            .Col = .ActiveCol: aCOL = .ActiveCol
                               varBar = Trim(.text)
                               
            If aCOL = 5 And Len(varBar) = 10 Then
                '바코드번호로 오더찾아오기
                Set mAdoRs = f_subSet_WorkList_Barcode(varBar)
                
                If RecordChk = True Then
                    Do Until mAdoRs.EOF
                        .SetText 1, aRow, "1"
                        .SetText 2, aRow, mAdoRs("SCP42IDNOA") & ""
                        .SetText 3, aRow, mAdoRs("SCP41NAME") & ""
                        .SetText 4, aRow, mAdoRs("SCP41JDATE") & ""
                        '.SetText 5, aRow, varBar
                        .SetText 6, aRow, mAdoRs("SCP42SUGACD") & ""
                        
                        For intCol = 10 To .MaxCols
                            .GetText intCol, 0, varTmp
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                If mAdoRs("SCP42SUGACD") & "" = Trim(itemX.SubItems(1)) Then
                                    'strOrdBuffer = strOrdBuffer & "DSP|" & 29 + intOrdCnt & "||" & Trim(itemX.Tag) & "^^^|||" & vbCr + vbLf '검사채널(test id)
                                    'intOrdCnt = intOrdCnt + 1
                                    spdResult1.Col = itemX.Index + 9
                                    spdResult1.BackColor = &HC6FEFF   '&HC6FEFF

                                    Exit For
                                End If
                            End If
                        Next
                        
                        mAdoRs.MoveNext
                        
                    Loop
                Else
                    .SetText 1, aRow, "0"
                    .SetText 2, aRow, ""
                    .SetText 3, aRow, ""
                    .SetText 4, aRow, ""
                    .SetText 5, aRow, varBar
                    .SetText 6, aRow, ""
                    
                    lblStatus.Caption = "바코드 번호 " & varBar & " 는 검사대상이 아닙니다"
                End If
                                                                                                
                Set mAdoRs = Nothing
            End If
        End With
    End If
    
End Sub

Private Sub spdWorklist_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim varTmp  As Variant
    
    If Row = 0 Then
        If Col = 1 Then
            Col = 2
        End If
        
        If OrderSort_Flag = 1 Then
            Call SpreadSheetSort(spdWorklist, Col, 2)
            OrderSort_Flag = 2
        Else
            Call SpreadSheetSort(spdWorklist, Col, 1)
            OrderSort_Flag = 1
        End If
    Else
        With spdWorklist
            .GetText 2, Row, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
    
            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
            cmdWorkList_Click
        End With
    End If
    
End Sub

Private Sub SpreadSheetSort(ByRef Spread As vaSpread, ByVal Col As Integer, Optional ByVal SortType As Integer = 1)
    Dim intCount As Integer
    Dim strDataField As String
    'SortType
    ' 0 : none
    ' 1 : ascending
    ' 2 : descending

    With Spread
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .DataRowCnt
        .SortBy = 0
        .SortKey(1) = Col       '정렬키 열번호

        If SortType = 0 Then
            .SortKeyOrder(1) = SortKeyOrderNone
        ElseIf SortType = 1 Then
            .SortKeyOrder(1) = SortKeyOrderAscending
        ElseIf SortType = 2 Then
            .SortKeyOrder(1) = SortKeyOrderDescending
        Else
            .SortKeyOrder(1) = SortKeyOrderAscending
        End If

        .Action = ActionSort
    End With

End Sub


Private Sub Timer1_Timer()

'    Call COM_OUTPUT(ENQ)
'    Debug.Print ENQ

'    lblStatus.Caption = "Machine IP : " & gDMSIP & Space(3) & "DMS Port : " & gDMSPort
'    lblStatus.Caption = Winsock1.state & "Machine IP : " & "http://his012edu.cmcnu.or.kr/himed/webapps/com/commonweb/xrw/.live?" & Space(3) & "DMS Port : " & 80

'Debug.Print Winsock1.state
End Sub

'Private Sub Timer2_Timer()
'    comEQP.Output = ACK
'End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub

Private Sub txtBarCode_Change()

    If txtBarCode.SelStart = txtBarCode.MaxLength Then SendKeys "{TAB}"
    
End Sub

Private Sub txtBarCode_GotFocus()

    With txtBarCode
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
End Sub

Private Sub txtBarCode_KeyPress(KeyAscii As Integer)

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub txtBarCode_LostFocus()"
    
    Dim varTmp  As Variant, strEqpCd    As String
    Dim IntRow  As Integer, intCol  As Integer, blnFlag As Boolean
    Dim strOrdcd() As String, strPid()  As String, strPnm() As String
    Dim strPexzm() As String, strPeqpcd() As String
    Dim strEqcode() As String, strExamname() As String, strAcptno() As String

    Dim itemX   As ListItem
    
    If txtBarCode.text = "" Then Exit Sub
    
    blnFlag = False
    If KeyAscii = vbKeyReturn Then
        intCol = sl_examdata_select&(txtBarCode.text, INS_CODE, strEqcode, strExamname, strOrdcd, strPid, strPnm, strAcptno)
        
        For intCol = 0 To UBound(strOrdcd)
            If strOrdcd(intCol) <> "" Then
                strEqpCd = f_funGet_CODE(strOrdcd(intCol))
                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    If Not blnFlag Then
                        IntRow = f_funGet_SpreadRow(spdResult1, 2, txtBarCode.text)
                        If IntRow < 1 Then
                            IntRow = f_funGet_SpreadRow(spdResult1, 2, "")
                            If IntRow < 1 Then
                                spdResult1.maxrows = spdResult1.maxrows + 1
                                spdResult1.RowHeight(spdResult1.maxrows) = 13
                                IntRow = spdWorklist.maxrows
                            End If
                            spdResult1.SetText 2, IntRow, txtBarCode.text
                            spdResult1.SetText 3, IntRow, strPnm(0)
                            spdResult1.SetText 4, IntRow, strPid(0)
                        End If
                        spdResult1.SetText 1, IntRow, "1"
                    End If
                        
                    'spdResult1.SetText itemX.Index + 6, intRow, "V"
                    spdResult1.Col = itemX.Index + 6
                    spdResult1.Row = IntRow
                    spdResult1.BackColor = &HC6FEFF
                    
                    blnFlag = True
                End If
            End If
        Next
    
        If Not blnFlag Then MsgBox "해당 검사항목이 존재하지 않은 검체입니다.", vbInformation, App.Title
        
        txtBarCode.text = "":   txtBarCode.SetFocus
        Exit Sub
    
    End If
    
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Function psDataExists() As Boolean
Dim sCnt As Long
    
    psDataExists = False
    With spdWorklist
        For sCnt = 1 To .maxrows
            .Row = sCnt:    .Col = 2
            If Trim(.text) = Mid(txtBarCode.text, 1, 11) Then
                psDataExists = True
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub txtBarCode_LostFocus()

'    Dim intRow      As Integer
'    Dim strOrdcd(1 To 100) As String
'
'    Call sl_spcid_tstcd_select&(txtBarCode.Text, strOrdcd)
'    If strOrdcd(1) = "" Then
'        MsgBox "해당 검사항목이 존재하지 않은 검체입니다.", vbInformation, Me.Caption
'        Exit Sub
'    End If
'
'    intRow = f_funGet_SpreadRow(spdWorkList, 2, txtBarCode.Text)
'    If intRow < 1 Then
'        intRow = f_funGet_SpreadRow(spdWorkList, 2, "")
'        If intRow < 1 Then
'            spdWorkList.maxrows = spdWorkList.maxrows + 1
'            spdWorkList.RowHeight(spdWorkList.maxrows) = 13
'            intRow = spdWorkList.maxrows
'        End If
'        spdWorkList.SetText 2, intRow, txtBarCode.Text
'    End If
'    spdWorkList.SetText 1, intRow, "1"
    
End Sub

Private Sub txtChart_GotFocus()
'
' Focus 가졌을 경우
'
    txtChart.ForeColor = &HFF&
    txtChart.text = ""
End Sub

Private Sub txtChart_LostFocus()
'
' Focus 가 없을 경우
'
    txtChart.ForeColor = &HFFC0C0
    txtChart.text = "차트번호 입력"
End Sub

Private Sub txtChart_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim intRow2 As Integer
    
    Dim tBlood As Boolean
        
    If Len(Trim(txtChart)) > 0 Then
        If KeyCode = 13 Then
        
          tBlood = False
          
          Rem txtChart = Format(txtChart, "0000000")
          
          intRow2 = f_funGet_SpreadRow(spdWorklist, 5, txtChart)
          
          If intRow2 >= 1 Then
              
              With spdWorklist
                .SetText 1, intRow2, "1"
                cmdWorkList_Click
                txtChart.text = ""
                tBlood = True
              End With
          End If
          
          If tBlood = False Then
            MsgBox txtChart.text & " 해당 환자의 처방이 없습니다.     ", vbInformation + vbOKOnly, App.Title
            txtChart.text = ""
          End If
        
         End If
    End If

End Sub

' ------------------------------------------------------------------------
' 통신상태 확인 관련이벤트
' ------------------------------------------------------------------------
Private Sub txtCom_Change()
    txtCom.SelStart = Len(txtCom.text)
End Sub

Private Sub cmdCOMLoad_Click()
    Dim i               As Long
    Dim lngFIleNum      As Long
    Dim strTemp         As String
    Dim strTemp2        As String
    Dim bteBuffer()     As Byte
    
On Error GoTo ErrorRoutine
    
    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowOpen
        lngFIleNum = FreeFile
        
        Open .FileName For Binary Access Read As #lngFIleNum
        
        txtCOM2.text = ""
        ReDim bteBuffer(LOF(lngFIleNum))
        Get #lngFIleNum, , bteBuffer

        strTemp = StrConv(bteBuffer, vbUnicode)
        txtCOM2.text = strTemp
                
        Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum
        
End Sub

Private Sub cmdCOMSave_Click()
    Dim lngFIleNum      As Long
    
On Error GoTo ErrorRoutine

    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowSave
        lngFIleNum = FreeFile
        
        Open .FileName For Append As #lngFIleNum
        Print #lngFIleNum, _
              Format(Date, "YYYY년 MM월 DD일") & "  "; time & vbNewLine & _
              "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbNewLine & _
              txtCom.text & _
              "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbNewLine
    Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum

End Sub

Private Sub cmdCOMOutput_Click()
    'Call COM_OUTPUT(StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode))
    Call COM_OUTPUT(charCOM_Convert(txtCom.SelText))
End Sub

Private Sub cmdCOMClear_Click()
    mlngRecLen = 0
    txtCom.text = ""
End Sub

Private Sub cmdCOMClear2_Click()
    txtCOM2.text = ""
End Sub

Private Sub cmdCOMInput_Click()

    Dim bytTemp() As Byte
    
    bytTemp = StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode)

    Call ComReceive(txtCom.SelText)
    
End Sub

Private Sub cmdCOMInput2_Click()
    
    Dim bytTemp() As Byte
    
    If txtCOM2.SelLength = 0 Then
        bytTemp = StrConv(charCOM_Convert(txtCOM2.text), vbFromUnicode)
    Else
        bytTemp = StrConv(charCOM_Convert(txtCOM2.SelText), vbFromUnicode)
    End If

    Call ComReceive(txtCOM2.SelText)

End Sub

Private Sub cmdCOMOutput2_Click()
    
    If txtCOM2.SelLength = 0 Then
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.text))
    Else
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.SelText))
    End If
    
End Sub
' ------------------------------------------------------------------------
' 통신상태 확인 관련이벤트


Private Sub txtResult_DblClick()
    txtResult.text = ""
    'List1.text = ""
    
    If txtResult.Visible Then txtResult.Visible = False
    'List1.Visible = True
End Sub

'Private Sub Winsock1_Close()
'    If Winsock1.state <> sckClosed Then
'        Winsock1.Close
'    End If
'
'    Winsock1.LocalPort = 80 'gDMSPort '"5001" 'gSetup.gPort
'    Winsock1.Listen
'
'    lblStatus.Caption = "신호 대기중..."
'
'End Sub
'
'Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'
'    If Winsock1.state <> sckClosed Then
'        Winsock1.Close
'    End If
'
'
'    Winsock1.Accept requestID
'    lblStatus.Caption = "연결[" & requestID & "]" & Winsock1.RemoteHostIP
'
'End Sub
'
'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'    Dim strRcvBuffer As String
'    Dim strSndBuffer As String
'
'    imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
'    If tmrReceive.Enabled = False Then
'        tmrReceive.Enabled = True
'    Else
'        tmrReceive.Enabled = False
'        tmrReceive.Enabled = True
'    End If
'
'    Dim sTmp As String
'    Dim strSendData
'    Dim strResFlag
'
'    Winsock1.GetData sTmp
'
'    Debug.Print sTmp
'
'
'End Sub
