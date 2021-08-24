VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   11250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16875
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11250
   ScaleWidth      =   16875
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  '최대화
   Begin VB.ListBox lstStatus 
      Height          =   1500
      ItemData        =   "frmComm.frx":0000
      Left            =   150
      List            =   "frmComm.frx":0007
      TabIndex        =   51
      Top             =   7410
      Width           =   12705
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   8610
      Left            =   60
      TabIndex        =   7
      Top             =   360
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   15187
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   16711680
      TabCaption(0)   =   " ▒    WorkList     "
      TabPicture(0)   =   "frmComm.frx":0016
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "spdWorklist"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdStartNo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAppend(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSPanel1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdWorkList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdNext"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPrevious"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkAuto"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "spdResult1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "spdRstview"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   " ▒   받은 결과     "
      TabPicture(1)   =   "frmComm.frx":0032
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSel(3)"
      Tab(1).Control(1)=   "cmdSel(2)"
      Tab(1).Control(2)=   "chkExcel"
      Tab(1).Control(3)=   "cmdExcel"
      Tab(1).Control(4)=   "cmdAppend(1)"
      Tab(1).Control(5)=   "SSPanel2"
      Tab(1).Control(6)=   "spdResult2"
      Tab(1).ControlCount=   7
      Begin FPSpread.vaSpread spdRstview 
         Height          =   8115
         Left            =   12870
         TabIndex        =   49
         Top             =   390
         Width           =   2325
         _Version        =   196608
         _ExtentX        =   4101
         _ExtentY        =   14314
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
         SpreadDesigner  =   "frmComm.frx":004E
         UserResize      =   0
      End
      Begin VB.Frame Frame4 
         Caption         =   "hidden"
         Height          =   6225
         Left            =   2490
         TabIndex        =   33
         Top             =   1680
         Visible         =   0   'False
         Width           =   8055
         Begin BS400.sckStringData sck 
            Height          =   300
            Left            =   6480
            TabIndex        =   47
            Top             =   2370
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   529
         End
         Begin VB.ListBox List1 
            Height          =   2580
            ItemData        =   "frmComm.frx":0815
            Left            =   150
            List            =   "frmComm.frx":0817
            TabIndex        =   43
            Top             =   3480
            Width           =   7215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "TEST"
            Height          =   375
            Left            =   6030
            TabIndex        =   36
            Top             =   1290
            Width           =   1230
         End
         Begin VB.Timer tmrSend 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   6840
            Top             =   750
         End
         Begin VB.Timer tmrReceive 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   6390
            Top             =   750
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   5970
            Top             =   750
         End
         Begin MSCommLib.MSComm comEQP 
            Left            =   6690
            Top             =   150
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            Handshaking     =   1
            RThreshold      =   1
            SThreshold      =   1
         End
         Begin MSComctlLib.ImageList imlList 
            Left            =   6120
            Top             =   150
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
                  Picture         =   "frmComm.frx":0819
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":0DB3
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":134D
                  Key             =   "NOF"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":18E7
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":1E81
                  Key             =   "LSE"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":241B
                  Key             =   "LSN"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList imlStatus 
            Left            =   5550
            Top             =   150
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
                  Picture         =   "frmComm.frx":29B5
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":2F4F
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":34E9
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":3A83
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":4315
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":446F
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":45C9
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvwCuData 
            Height          =   3000
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   5292
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   5070
            Top             =   210
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSWinsockLib.Winsock Winsock1 
            Left            =   6810
            Top             =   1830
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
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
            Left            =   180
            TabIndex        =   44
            Top             =   3240
            Width           =   1755
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            DrawMode        =   5  '카피 펜이 아님
            X1              =   1620
            X2              =   7320
            Y1              =   3360
            Y2              =   3360
         End
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   6630
         Left            =   90
         TabIndex        =   48
         Top             =   420
         Width           =   12705
         _Version        =   196608
         _ExtentX        =   22410
         _ExtentY        =   11695
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
         SpreadDesigner  =   "frmComm.frx":4723
         UserResize      =   0
      End
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  '평면
         Caption         =   "Auto Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   12270
         TabIndex        =   35
         Top             =   30
         Value           =   1  '확인
         Width           =   1320
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   8
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":4BE8
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   9
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm.frx":506A
      End
      Begin BHButton.BHImageButton cmdPrevious 
         Height          =   330
         Left            =   90
         TabIndex        =   30
         Top             =   5640
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
         ForeColor       =   -2147483635
         BackColor       =   16711680
         AlphaColor      =   255
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdNext 
         Height          =   330
         Left            =   330
         TabIndex        =   31
         Top             =   5640
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
         TransparentPicture=   "frmComm.frx":54D8
         ForeColor       =   16711680
         BackColor       =   255
         AlphaColor      =   255
         ImgOutLineSize  =   3
      End
      Begin VB.Frame Frame3 
         Height          =   345
         Left            =   90
         TabIndex        =   27
         Top             =   900
         Width           =   555
         Begin Threed.SSCommand cmdSel 
            Height          =   345
            Index           =   1
            Left            =   270
            TabIndex        =   29
            Top             =   0
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   609
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "frmComm.frx":594A
         End
         Begin Threed.SSCommand cmdSel 
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   609
            _StockProps     =   78
            ForeColor       =   14735310
            BevelWidth      =   1
            Picture         =   "frmComm.frx":5DCC
         End
      End
      Begin VB.CheckBox chkExcel 
         Appearance      =   0  '평면
         BackColor       =   &H80000004&
         Caption         =   "Excel 생성"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -61080
         TabIndex        =   26
         Top             =   30
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1245
      End
      Begin BHButton.BHImageButton cmdExcel 
         Height          =   390
         Left            =   -66060
         TabIndex        =   25
         Top             =   330
         Visible         =   0   'False
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   688
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
      Begin BHButton.BHImageButton cmdWorkList 
         Height          =   435
         Left            =   90
         TabIndex        =   14
         Top             =   5160
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   767
         Caption         =   "WorkList 등록"
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
      Begin Threed.SSPanel SSPanel1 
         Height          =   525
         Left            =   90
         TabIndex        =   17
         Top             =   360
         Width           =   4755
         _Version        =   65536
         _ExtentX        =   8387
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
         Begin MSMask.MaskEdBox mskOrdDate1 
            Height          =   300
            Left            =   2415
            TabIndex        =   18
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskOrdDate 
            Height          =   300
            Left            =   1170
            TabIndex        =   19
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin BHButton.BHImageButton cmdSearch 
            Height          =   390
            Left            =   3570
            TabIndex        =   32
            Top             =   60
            Width           =   1110
            _ExtentX        =   1958
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
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   24
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
            Left            =   2280
            TabIndex        =   21
            Top             =   180
            Width           =   315
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            TabIndex        =   20
            Top             =   180
            Width           =   1095
         End
      End
      Begin BHButton.BHImageButton cmdAppend 
         Height          =   375
         Index           =   1
         Left            =   -62355
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
         Caption         =   "서버등록"
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
         Left            =   13710
         TabIndex        =   16
         Top             =   15
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         Caption         =   "서버등록"
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
      Begin BHButton.BHImageButton cmdStartNo 
         Height          =   420
         Left            =   4920
         TabIndex        =   22
         Top             =   405
         Width           =   1470
         _ExtentX        =   2593
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
      Begin FPSpread.vaSpread spdWorklist 
         Height          =   4230
         Left            =   90
         TabIndex        =   23
         Top             =   900
         Width           =   4755
         _Version        =   196608
         _ExtentX        =   8387
         _ExtentY        =   7461
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   1
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
         MaxCols         =   10
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":623A
         UserResize      =   2
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   525
         Left            =   -74910
         TabIndex        =   37
         Top             =   360
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
            Index           =   1
            ItemData        =   "frmComm.frx":6839
            Left            =   2235
            List            =   "frmComm.frx":6846
            Style           =   2  '드롭다운 목록
            TabIndex        =   40
            Top             =   135
            Width           =   1410
         End
         Begin MSMask.MaskEdBox mskRstDate 
            Height          =   300
            Left            =   1110
            TabIndex        =   41
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
            TabIndex        =   42
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
         Begin VB.Label Label12 
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
            TabIndex        =   39
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   38
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7350
         Left            =   -74910
         TabIndex        =   50
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
         SpreadDesigner  =   "frmComm.frx":6870
         UserResize      =   0
      End
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
      Left            =   30
      TabIndex        =   1
      Top             =   9000
      Width           =   15315
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   6615
         TabIndex        =   10
         Top             =   90
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
         TabIndex        =   11
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
         Left            =   9225
         TabIndex        =   12
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
         TabIndex        =   13
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
         TransparentPicture=   "frmComm.frx":6D10
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   345
         Index           =   5
         Left            =   12930
         TabIndex        =   45
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         Caption         =   "닫기"
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
         Height          =   345
         Index           =   4
         Left            =   12030
         TabIndex        =   46
         Top             =   90
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         Caption         =   "연결"
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
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16875
      _ExtentX        =   29766
      _ExtentY        =   609
      Border          =   1
      CaptionBackColor=   16777215
      Caption         =   " Communication"
      SubCaption      =   "검사 장비와 통신하여 결과를 저장 합니다."
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
         Left            =   14145
         TabIndex        =   4
         Top             =   75
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   13110
         TabIndex        =   3
         Top             =   75
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Left            =   12015
         TabIndex        =   2
         Top             =   75
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   15015
         Picture         =   "frmComm.frx":859A
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13725
         Picture         =   "frmComm.frx":8B24
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12525
         Picture         =   "frmComm.frx":90AE
         Top             =   45
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

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""
Const SB As String = ""  'Chr(11)
Const EB As String = ""   'Chr(28)

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


Private Const TEST_NM_EQP   As String = "EQP_NM"    '장비 코드
Private Const TEST_CD_LIS   As String = "LIS_CD"    '검사실 코드
Private Const TEST_NM_LIS   As String = "LIS_NM"    '검사실 이름
Private Const TEST_VALUES   As String = "VALUES"    '결과

Private mAdoRs          As ADODB.Recordset
Private mAdoRs1         As ADODB.Recordset
Private CallForm        As String
Private IS_SET          As Boolean
Private f_strBuffer     As String
Private f_strOrdList    As String

Dim fChannel()      As String
Dim brStr           As String
Dim fRcvString      As String
Dim sStxCheck       As Integer
Dim sEtxCheck       As Integer
Dim sLfCheck        As Integer
Dim sCrcheck        As Integer
Dim RecordChk       As Boolean
Dim strGumCd        As String
Dim strJinCd        As String
Dim OrderSort_Flag  As Integer
Dim gspdResultRow   As Integer

Private Type TYPE_CD
    strEqpCd        As String
    intCnt          As Integer
    strTestcd(50)   As String
End Type

Private f_typCode() As TYPE_CD

Private Type typeLabUReader
    SID             As String 'Sid
    TestId(50)      As String
    Result(50)      As String
End Type

Private LabUReader As typeLabUReader

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
    Dim intRow  As Integer
    
    f_funGet_SpreadRow = 0
    
    With objSpd
        For intRow = 1 To .MaxRows
            .GetText intCol, intRow, varTmp
            If Trim$(varTmp) = strPara Then
                f_funGet_SpreadRow = intRow
                Exit For
            End If
        Next
    End With
    
End Function

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
            Call .Add(, "REMARK", "검체코드", (lvwCuData.Width - 310) * 0.1)
        End With
        .HideColumnHeaders = False
    End With
    
   
End Sub

Private Function f_subSet_WorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String)
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim pFrDt As String
    Dim pToDt As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
   
    pFrDt = mskOrdDate.text
    pToDt = mskOrdDate1.text
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
                '-- 처방일자,처방일련번호,환자명,환자번호,입외구분,일련번호,성별,나이,내원번호,처방코드
             sqlDoc = "Select a.ORDT,a.ORQN,b.PANM,a.PAID,a.OIFL,a.SENO,b.SEXS,b.AGES,a.NWNO,a.ORCD "
    sqlDoc = sqlDoc & "  From LRESULT a, APATINF b"
'    sqlDoc = sqlDoc & " Where a.ORDT between  '" & mskOrdDate.text & "' and '" & mskOrdDate1.text & "'"
    sqlDoc = sqlDoc & " Where a.ETDT between TO_DATE(" & Format(pFrDt, "########") & ",'yyyymmdd') + 0.000000 "
    sqlDoc = sqlDoc & "    and TO_DATE(" & Format(pToDt, "########") & ",'yyyymmdd') + 0.999999 " & vbCrLf
    sqlDoc = sqlDoc & "   And a.PAID = b.PAID "
    sqlDoc = sqlDoc & "   And a.ORCD in (" & strGumCd & ")"
    sqlDoc = sqlDoc & "   And a.OKFL <> 'Y' "   '-- 결과확정유무
    sqlDoc = sqlDoc & " Order By a.ORDT,a.PAID,a.ORQN,b.PANM,a.SENO"
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE
    
    If AdoRs_ORACLE.RecordCount = 0 Then
        Set f_subSet_WorkList = Nothing
        RecordChk = False
        Set AdoRs_ORACLE = Nothing
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


Private Function f_subSet_WorkList_Barcode(ByVal strORDT As String, Optional ByVal strPAID As String, Optional ByVal strSENO As String)
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
        sqlDoc = sqlDoc & vbCrLf & "   AND b.SCP42SUGACD in (" & strGumCd & ")"
        
'        sqlDoc = sqlDoc & vbCrLf & "   AND (b.SCP42RSTCD <> 'N' OR b.SCP42RSTCD IS null)"
'        sqlDoc = sqlDoc & vbCrLf & "   AND (b.SCP42RSTCD <> 'N' OR b.SCP42RSTCD IS null OR b.SCP42PROFLG  <> 'M')"
        
        '-- 2012.04.13 수정
        'sqlDoc = sqlDoc & vbCrLf & "   AND a.SCP41SNDYN = 'N' "
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


Private Function f_subSet_RefVal(ByVal strORCD As String, Optional ByVal strRSLT As String, Optional ByVal strSex As String, Optional ByVal strAGE As String) As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim stryy, strmm, strdd, strDate  As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_RefVal() As ADODB.Recordset"
    
    strRSLT = Replace(strRSLT, "<", "")
    strRSLT = Replace(strRSLT, ">", "")
    f_subSet_RefVal = " "
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    f_subSet_RefVal = ""
    If strAGE <> "" Then
        If strAGE <= 7 Then
            sqlDoc = "Select YMAX as MAX, YMIN as MIN "
        Else
            If strSex = "M" Then
                     sqlDoc = "Select MMAX as MAX, MMIN as MIN "
            Else
                     sqlDoc = "Select WMAX as MAX, WMIN as MIN "
            End If
        End If
    Else
        sqlDoc = "Select MMAX as MAX, MMIN as MIN "
    End If
    
    sqlDoc = sqlDoc & "  From LABMAST"
    sqlDoc = sqlDoc & " Where ORCD =  '" & strORCD & "'"

    Set AdoRs_ORACLE = New ADODB.Recordset
    
    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE
    
    If AdoRs_ORACLE.RecordCount = 0 Then
        f_subSet_RefVal = " "
        Set AdoRs_ORACLE = Nothing
        Exit Function
    Else
        If IsNumeric(strRSLT) And IsNumeric(AdoRs_ORACLE.Fields("MAX")) And IsNumeric(AdoRs_ORACLE.Fields("MIN")) Then
            If Val(strRSLT) > Val(AdoRs_ORACLE.Fields("MAX")) Then
                f_subSet_RefVal = "H"
            ElseIf Val(strRSLT) < Val(AdoRs_ORACLE.Fields("MIN")) Then
                f_subSet_RefVal = "L"
            Else
                f_subSet_RefVal = " "
            End If
        Else
            f_subSet_RefVal = " "
        End If
    End If

    Set AdoRs_ORACLE = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    
    Call ErrMsgProc(CallForm)
     
End Function

Private Sub f_subSet_ItemList()

    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim AdoRs   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim strTest As String, intPos   As Integer
    Dim strTmp  As String, intCol   As Integer, intCol2   As Integer, intCnt  As Integer, intRow  As Integer
    
    Dim intPos1 As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 10
    intCol2 = 1
    intRow = 1
    With spdWorklist
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .MaxRows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .MaxRows = 1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .MaxRows = 1
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
             
    AdoRs.CursorLocation = adUseClient
    AdoRs.Open sqlDoc, AdoCn_Jet
    If AdoRs.RecordCount > 0 Then
        AdoRs.MoveFirst
        ReDim fChannel(AdoRs.RecordCount)
        strJinCd = ""
        strGumCd = ""
    End If
    
    Do While Not AdoRs.EOF
        If Trim(AdoRs.Fields("TESTCD")) <> "" Then
            intPos1 = InStr(Trim(AdoRs.Fields("TESTCD")), ",")
            If intPos1 = 0 Then
                strGumCd = strGumCd & "'" & Trim(AdoRs.Fields("TESTCD")) & "',"
            Else
                strGumCd = strGumCd & "'" & Mid(Trim(AdoRs.Fields("TESTCD")), 1, intPos1 - 1) & "',"
                strJinCd = strJinCd & "" & Mid(Trim(AdoRs.Fields("TESTCD")), intPos1 + 1) & ","
            End If
        End If
        
        Set itemX = lvwCuData.ListItems.Add(, , Trim(AdoRs.Fields("TEST_EQP") & ""), , "LST")
            itemX.SubItems(1) = Trim(AdoRs.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(AdoRs.Fields("TESTNM") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(AdoRs.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(AdoRs.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(AdoRs.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(AdoRs.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(AdoRs.Fields("REFL") & "")
            itemX.SubItems(9) = Trim(AdoRs.Fields("REFH") & "")
            itemX.SubItems(10) = Trim(AdoRs.Fields("AUTOVERIFY") & "")
            itemX.SubItems(11) = Trim(AdoRs.Fields("REMARK") & "")
            itemX.Tag = Trim(AdoRs.Fields("TEST_EQP") & "")
            itemX.text = Trim(AdoRs.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdWorklist
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(AdoRs("TESTNM") & "")
            .Col = intCol:  .ColHidden = True
        End With
        
        With spdResult1
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 6.5
            End If
            .SetText intCol, 0, Trim$(AdoRs("TESTNM") & "")
        End With
        
        With spdRstview
            If intRow > .MaxRows Then
                intRow = 1
                intCol2 = intCol2 + 2
            End If
            
            .SetText intCol2, intRow, Trim$(AdoRs("TESTNM") & "")
            intRow = intRow + 1
            
        End With
        
        With spdResult2
            If intCol > .MaxCols Then
                .MaxCols = .MaxCols + 1
                .ColWidth(intCol) = 6.5
            End If
            .SetText intCol, 0, Trim$(AdoRs("TESTNM") & "")
        End With
        
        fChannel(intCol - colNWNO) = AdoRs.Fields("TEST_EQP")
        
        intCnt = intCnt + 1
        ReDim Preserve f_typCode(1 To intCnt) As TYPE_CD
        
        f_typCode(intCnt).strEqpCd = Trim$(AdoRs.Fields("TEST_EQP"))
        f_typCode(intCnt).intCnt = 0
        
        strTmp = Trim$(AdoRs.Fields("TESTCD"))
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
        
        AdoRs.MoveNext
    Loop
    Set AdoRs = Nothing
    
    If Trim(strGumCd) <> "" Then strGumCd = Mid(strGumCd, 1, Len(strGumCd) - 1)
    If Trim(strJinCd) <> "" Then strJinCd = Mid(strJinCd, 1, Len(strJinCd) - 1)
    
    With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol:  .ColHidden = True
    End With

Exit Sub
ErrRoutine:
    Set AdoRs = Nothing
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

Private Sub cmdAppend_Click(Index As Integer)
   
    Dim AdoRs   As New ADODB.Recordset
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

    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
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
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"

On Error GoTo ErrorRoutine

    Me.MousePointer = 11

    If Index = 0 Then
        Set objSpd = spdResult1
    Else
        Set objSpd = spdResult2
    End If

    With objSpd
        For intRow = 1 To .MaxRows
            .GetText colORDT, intRow, varTmp:    strORDT = Trim$(varTmp)
            .GetText colORQN, intRow, varTmp:    strORQN = Trim$(varTmp)
            .GetText colPANM, intRow, varTmp:    strPANM = Trim$(varTmp)
            .GetText colPAID, intRow, varTmp:    strPAID = Trim$(varTmp): strBarno = strPAID
            .GetText colOIFL, intRow, varTmp:    strOIFL = Trim$(varTmp)
            .GetText colSENO, intRow, varTmp:    strSENO = Trim$(varTmp)
            .GetText colSEXS, intRow, varTmp:    strSEXS = Trim$(varTmp)
            .GetText colAGES, intRow, varTmp:    strAGES = Trim$(varTmp)
            .GetText colNWNO, intRow, varTmp:    strNWNO = Trim$(varTmp)

            .GetText colBANO, intRow, varTmp

            If strPAID = "" Then Exit For

            intCnt = 0: Erase strOrdcd ': Erase strRstval
            
            If Trim$(varTmp) = "1" Then
                For intCol = 10 To .MaxCols
                    strDate = Format$(Now, "YYYYMMDD"):    strTime = Format$(Now, "HHMMSS")
                    .GetText intCol, intRow, varTmp
                        If Trim$(varTmp) <> "" Then
                            .GetText intCol, 0, varTmp
                            Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                            If Not itemX Is Nothing Then
                                .GetText intCol, intRow, varTmp: strRstval = varTmp
                                strTestcd = itemX.ListSubItems(1)
                                intPos = InStr(strTestcd, ",")
                                strEqpCd = itemX.text
                                                    
                                If strEqpCd <> "" Then
                                    '-- 로컬저장
                                    sqlDoc = "Update INTERFACE003" & _
                                             "   set RSTVAL  = '" & strRstval & "', REFVAL = '" & strRefVal & "'" & _
                                             " where SPCNO   = '" & strBarno & "'" & _
                                             "   and EQPNUM  = '" & itemX.Tag & "'" & _
                                             "   and TRANSDT = '" & strDate & "'" & _
                                             "   and TRANSTM = '" & strTime & "'"
                                    AdoCn_Jet.Execute sqlDoc
                                    
                                    sqlDoc = "insert into INTERFACE003(" & _
                                             "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                             "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.Tag & "'," & _
                                             "            '" & strDate & "', '" & strTime & "'," & _
                                             "            '" & strRstval & "', '" & strRefVal & "'," & _
                                             "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                    AdoCn_Jet.Execute sqlDoc
                                    
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
                                                                                                                    
                                    .Row = intRow: .Col = colBANO: .Value = 0
                                                   .Col = colORDT: .BackColor = HNC_Cyan
                                                   .Col = colORQN: .BackColor = HNC_Cyan
                                                   .Col = colPANM: .BackColor = HNC_Cyan
                                                   .Col = colPAID: .BackColor = HNC_Cyan
                                                   .Col = colOIFL: .BackColor = HNC_Cyan
                                                   .Col = colSENO: .BackColor = HNC_Cyan
                                                   '.Col = colSEXS: .BackColor = HNC_Cyan
                                                   '.Col = colAGES: .BackColor = HNC_Cyan
                                                   '.Col = colNWNO: .BackColor = HNC_Cyan
                                                                        
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
        MsgBox "▒ EMR SERVER에 결과를 Upload 완료되었습니다. ▒      " & vbCrLf & vbCrLf & "     LIS 결과조회 화면에서 결과를 확인 하십시요..  ", vbInformation, App.Title
    End If
    
    Exit Sub
ErrorRoutine:

    Set AdoRs_SQL = Nothing

    Set itemX = Nothing

    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)
End Sub

Private Function SeqSearch_PAID(ByVal brspread As Object, ByVal brQry1 As String, ByVal brQry2 As String, ByVal brQry3 As String, ByVal brCol As Integer) As Long
Dim sCnt As Long
Dim sCnt1 As Long
Dim sCnt2 As Long
Dim sCnt3 As Long

    SeqSearch_PAID = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt1 = 1 To .MaxRows
            .Row = sCnt1
            .Col = 2
            If Trim(.text) = brQry1 Then
                .Col = 5
                If Trim(.text) = brQry3 Then
                    SeqSearch_PAID = sCnt1
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            End If
        Next sCnt1
    End With

End Function

Private Sub cmdSearch_Click()
    Dim strDoc As String
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim intRow      As Integer
    Dim pGrid_Point As Integer
    Dim intCnt      As Integer
    Dim strBarno As String
    Dim itemX As ListItem
    Dim strEqpCd As String
    Dim strBartmpNo As String
    Dim blt As Boolean
    
    With spdWorklist
        .MaxRows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    blt = True
    
On Error GoTo ErrorTrap

    CallForm = "clsCommon - Public Function cmdSearch_Click() As ADODB.Recordset"
        
    Set mAdoRs = f_subSet_WorkList(mskOrdDate.text, mskOrdDate1.text)
    
    If RecordChk = False Then
        MsgBox Format(mskOrdDate.text, "####-##-##") & "일 에서  " & Format(mskOrdDate1.text, "####-##-##") & "일까지의 검사 대상자가 없습니다.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    Else
        strBarno = ""
        mAdoRs.MoveFirst

        With spdWorklist
            For intCnt = 0 To mAdoRs.RecordCount - 1
                '-- 처방일자,처방일련번호,환자명,환자번호,입외구분,일련번호,성별,나이,내원번호,처방코드
                If strBarno <> Trim(mAdoRs.Fields("ORDT").Value & "") & Trim(mAdoRs.Fields("PAID").Value & "") Then
                    pGrid_Point = SeqSearch_PAID(spdWorklist, Trim(mAdoRs.Fields("ORDT").Value & ""), "", Trim(mAdoRs.Fields("PAID")), colORQN)

                    If pGrid_Point = 0 Then
                        pGrid_Point = SeqNullSearch(spdWorklist, Trim(mAdoRs.Fields("ORDT")), colORDT)
                        If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
                    End If
                    
                    .SetText colBANO, pGrid_Point, "0"
                    .SetText colORDT, pGrid_Point, mAdoRs("ORDT").Value & ""
                    .SetText colORQN, pGrid_Point, mAdoRs("ORQN").Value & ""
                    .SetText colPANM, pGrid_Point, Trim(mAdoRs("PANM").Value & "")
                    .SetText colPAID, pGrid_Point, Trim(mAdoRs("PAID").Value & "")
                    .SetText colOIFL, pGrid_Point, Trim(mAdoRs("OIFL").Value & "")
                    .SetText colSENO, pGrid_Point, Trim(mAdoRs("SENO").Value & "")
                    .SetText colSEXS, pGrid_Point, Trim(mAdoRs("SEXS").Value & "")
                    .SetText colAGES, pGrid_Point, Trim(mAdoRs("AGES").Value & "")
                    .SetText colNWNO, pGrid_Point, Trim(mAdoRs("NWNO").Value & "")
                    
                    .Row = pGrid_Point: .Col = colBANO: .ForeColor = HNC_Black
                                        .Col = colORDT: .ForeColor = HNC_Black
                                        .Col = colORQN: .ForeColor = HNC_Black
                                        .Col = colPANM: .ForeColor = HNC_Black
                                        .Col = colPAID: .ForeColor = HNC_Black
                                        .Col = colOIFL: .ForeColor = HNC_Black
                                        .Col = colSENO: .ForeColor = HNC_Black
                                        .Col = colSEXS: .ForeColor = HNC_Black
                                        .Col = colAGES: .ForeColor = HNC_Black
                                        .Col = colNWNO: .ForeColor = HNC_Black
                    
                End If

                strBarno = Trim(mAdoRs.Fields("ORDT").Value & "") & Trim(mAdoRs.Fields("PAID").Value & "")
                mAdoRs.MoveNext
            Next
            
            If blt = False Then
                .Row = pGrid_Point
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
            End If
        End With
    End If
    
    Set mAdoRs = Nothing
    
    spdWorklist.Row = 1
    spdWorklist.Col = 1
    spdWorklist.Action = ActionActiveCell
        
Exit Sub

ErrorTrap:
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)
    
End Sub


Private Sub cmdAction_Click(Index As Integer)
    Dim TxtIP As String
    
    Select Case Index
        Case 0:     Call cmdRun
        Case 1:     Call cmdStop
        Case 2:     Call cmdClear
        Case 3:     Call cmdExit
        Case 4:
            TxtIP = Winsock1.LocalIP
            Winsock1.LocalPort = CInt(5051)
            Winsock1.Listen
            cmdAction(4).Enabled = False
            cmdAction(5).Enabled = True
        Case 5
            Winsock1.Close
            cmdAction(4).Enabled = True
            cmdAction(5).Enabled = False
        Case Else
        
    End Select
    
End Sub

Private Sub cmdClear()
    
    List1.Clear
    
    With spdWorklist
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
        .MaxRows = 1
        
    End With
    
    With spdResult1
        .MaxRows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 13
    End With

    With spdResult2
        .MaxRows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 13
    End With
    
    Dim Rowcnt As Integer
    Dim Colcnt As Integer

    With spdRstview
        For Rowcnt = 1 To 8
            For Colcnt = 2 To 6 Step 2
                .Row = Rowcnt
                .Col = Colcnt
                .BackColor = &HFFFFFF
                .text = ""
            Next Colcnt
        Next Rowcnt
    End With

End Sub

Private Sub cmdExit()
    
    Unload Me

End Sub

Public Sub cmdRun()
    
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
'    If Not comEQP.PortOpen Then comEQP.PortOpen = True
'    If Winsock1.Connect Then
    If sck.state = "Connected" Then
'    If comEQP.PortOpen Then
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

Private Sub cmdRstQuery_Click()
'
'    Dim AdoRs   As New ADODB.Recordset
'    Dim sqlDoc  As String, intRet   As Integer
'
'    Dim strSpcno    As String
'    Dim intRow      As Integer, intCol  As Integer
'    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
'
'    Dim itemX       As ListItem
'
'    intRow = 0
'    With spdResult2
'        .MaxRows = 1
'        .Col = 1:   .Col2 = .MaxCols
'        .Row = 1:   .Row2 = .MaxRows
'        .BlockMode = True
'        .Action = ActionClearText
'        .BlockMode = False
'    End With
'
'    sqlDoc = "Select ORDT, ORQN, PAID, OIFL, SENO, NWNO, ORCD, TRANSDT, TRANSTM, EqpCD, RSTVAL, REFVAL, SERVERGBN, PANM, SEX, AGE" & _
'             "  From INTERFACE004" & _
'             " Where TRANSDT >= '" & mskRstDate.text & "'"
'
'    If cboRstgbn(1).ListIndex = 0 Then
'        sqlDoc = sqlDoc & "   And (SERVERGBN = '' or SERVERGBN = 'N')"
'    ElseIf cboRstgbn(1).ListIndex = 1 Then
'        sqlDoc = sqlDoc & "   And SERVERGBN = 'Y'"
'    End If
'    sqlDoc = sqlDoc & " Order By ORDT, TRANSDT,TRANSTM"
'
'    AdoRs.CursorLocation = adUseClient
'    AdoRs.Open sqlDoc, AdoCn_Jet
'    If AdoRs.RecordCount > 0 Then AdoRs.MoveFirst
'    Do While Not AdoRs.EOF
'        With spdResult2
'            If strSpcno <> Trim$(AdoRs("ORDT") & "") & Trim$(AdoRs("ORQN") & "") & Trim$(AdoRs("PAID") & "") Then
'                intRow = intRow + 1
'                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
'                .SetText 1, intRow, "1"
'                .SetText 2, intRow, AdoRs("ORDT").Value & ""
'                .SetText 3, intRow, AdoRs("ORQN").Value & ""
'                .SetText 4, intRow, Trim(AdoRs("PANM").Value & "")
'                .SetText 5, intRow, Trim(AdoRs("PAID").Value & "")
'                .SetText 6, intRow, Trim(AdoRs("OIFL").Value & "")
'                .SetText 7, intRow, Trim(AdoRs("SENO").Value & "")
'                .SetText 8, intRow, Trim(AdoRs("SEX").Value & "")
'                .SetText 9, intRow, Trim(AdoRs("AGE").Value & "")
'                .SetText 10, intRow, Trim(AdoRs("NWNO").Value & "")
'                .SetText 11, intRow, Format(Trim(AdoRs("TRANSDT").Value & ""), "####-##-##")
'            End If
'            strSpcno = Trim$(AdoRs("ORDT") & "") & Trim$(AdoRs("ORQN") & "") & Trim$(AdoRs("PAID") & "")
'            Set itemX = lvwCuData.FindItem(Trim$(AdoRs("EQPCD") & ""), lvwTag, , lvwWhole)
'            If Not itemX Is Nothing Then
'                intCol = itemX.Index + 11
'                .SetText intCol, intRow, Trim$(AdoRs("RSTVAL")) & ""
'                .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(AdoRs("REFVAL") & "") <> "", vbRed, vbBlack)
'            End If
'        End With
'        AdoRs.MoveNext
'    Loop
'    AdoRs.Close:    Set AdoRs = Nothing
'



    Dim AdoRs   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    Dim itemX       As ListItem

    intRow = 0
    With spdResult2
        .MaxRows = 25
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
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
    
    AdoRs.CursorLocation = adUseClient
    AdoRs.Open sqlDoc, AdoCn_Jet
    If AdoRs.RecordCount > 0 Then AdoRs.MoveFirst
    Do While Not AdoRs.EOF
        With spdResult2
        If strSpcno <> Trim$(AdoRs(0) & "") + Trim$(AdoRs(6) & "") Then
                intRow = intRow + 1
                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(AdoRs(3) & "")
                .SetText 3, intRow, Trim$(AdoRs(0) & "")
                .SetText 6, intRow, Trim$(AdoRs(8) & "")
                .SetText 7, intRow, Trim$(AdoRs(9) & "")
                '.SetText .MaxCols, intRow, Trim$(adoRS(6) & "")
            End If
                strSpcno = Trim$(AdoRs(0) & "") + Trim$(AdoRs(6) & "")
                Set itemX = lvwCuData.FindItem(Trim$(AdoRs(7) & ""), lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    intCol = itemX.Index + 8
                    .SetText intCol, intRow, Trim$(AdoRs(4)) & ""
                    .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(AdoRs(5) & "") <> "", vbRed, vbBlack)
                End If
        End With
        AdoRs.MoveNext
    Loop
'    spdResult2.MaxCols = spdResult2.MaxCols - 1
    AdoRs.Close:    Set AdoRs = Nothing
    
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    If Index = 2 Or Index = 3 Then
        With spdResult1
            For intRow = 1 To .MaxRows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
    Else
        With spdWorklist
            For intRow = 1 To .MaxRows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
    End If
    
End Sub

Private Sub cmdStartNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    
    sNo = InputBox("시작 번호를 입력하세요 !")
    If Len(sNo) > 0 And spdResult1.MaxRows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "숫자만 입력하세요.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .MaxRows
                .Row = sCnt
                .Col = colSQNO:       .text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If

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
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD, strSQNO As String
    
    blnFlag = False
    
    With spdWorklist
        For intRow1 = 1 To .MaxRows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                    '-- 처방일자,처방일련번호,환자명,환자번호,입외구분,일련번호,성별,나이,내원번호,처방코드
                .GetText colORDT, intRow1, varTmp:    strORDT = Trim$(varTmp)
                .GetText colORQN, intRow1, varTmp:    strORQN = Trim$(varTmp)
                .GetText colPANM, intRow1, varTmp:    strPANM = Trim$(varTmp)
                .GetText colPAID, intRow1, varTmp:    strPAID = Trim$(varTmp)
                .GetText colOIFL, intRow1, varTmp:    strOIFL = Trim$(varTmp)
                .GetText colSENO, intRow1, varTmp:    strSENO = Trim$(varTmp)
                .GetText colSEXS, intRow1, varTmp:    strSEXS = Trim$(varTmp)
                .GetText colAGES, intRow1, varTmp:    strAGES = Trim$(varTmp)
                .GetText colNWNO, intRow1, varTmp:    strNWNO = Trim$(varTmp)
                .GetText colSQNO, intRow1, varTmp:    strSQNO = Trim$(varTmp)
                
                .Row = intRow1: .Col = colBANO: .ForeColor = HNC_Red
                                .Col = colORDT: .ForeColor = HNC_Red
                                .Col = colORQN: .ForeColor = HNC_Red
                                .Col = colPANM: .ForeColor = HNC_Red
                                .Col = colPAID: .ForeColor = HNC_Red
                                .Col = colOIFL: .ForeColor = HNC_Red
                                .Col = colSENO: .ForeColor = HNC_Red
                                .Col = colSEXS: .ForeColor = HNC_Red
                                .Col = colAGES: .ForeColor = HNC_Red
                                .Col = colNWNO: .ForeColor = HNC_Red
                
                intRow2 = f_funGet_SpreadRow_PAID(spdResult1, colORQN, strORDT, strORQN, strPAID)
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, colORDT, "")
                    If intRow2 < 1 Then
                        spdResult1.MaxRows = spdResult1.MaxRows + 1
                        spdResult1.RowHeight(spdResult1.MaxRows) = 13
                        intRow2 = spdResult1.MaxRows
                    End If

                    blnFlag = False
                    
                    Set mAdoRs = f_subSet_WorkList_Barcode(strORDT, strPAID, strSENO)

                    If Len(strPAID) > 0 And Not mAdoRs Is Nothing Then
                        Do Until mAdoRs.EOF
                            strEqpCd = f_funGet_CODE(Trim(mAdoRs("ORCD").Value))
                            
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                blnFlag = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + colSQNO
                                spdResult1.BackColor = &HC6FEFF '&H80C0FF
                                spdResult1.ForeColor = vbWhite '&H80C0FF
                                spdResult1.text = Trim(mAdoRs("ORQN").Value) & ""
                                DoEvents
                            End If
                            mAdoRs.MoveNext
                        Loop
                    End If
                    If blnFlag = True Then
                        spdResult1.SetText colORDT, intRow2, strORDT
                        spdResult1.SetText colORQN, intRow2, strORQN
                        spdResult1.SetText colPANM, intRow2, strPANM
                        spdResult1.SetText colPAID, intRow2, strPAID
                        spdResult1.SetText colOIFL, intRow2, strOIFL
                        spdResult1.SetText colSENO, intRow2, strSENO
                        spdResult1.SetText colSEXS, intRow2, strSEXS
                        spdResult1.SetText colAGES, intRow2, strAGES
                        spdResult1.SetText colNWNO, intRow2, strNWNO
                        spdResult1.SetText colSQNO, intRow2, strSQNO
                        
                        spdResult1.Row = intRow2:
                        spdResult1.Col = colSQNO:
                        spdResult1.ForeColor = HNC_Red
                    Else
                        spdResult1.MaxRows = spdResult1.MaxRows - 1
                    End If
                End If
                
                .SetText 1, intRow1, ""
            End If
        Next
    End With
                
End Sub


Private Function f_funGet_SpreadRow_PAID(ByVal objSpd As vaSpread, ByVal intCol As Integer, _
                                    ByVal strPara1 As String, ByVal strPara2 As String, ByVal strPara3 As String) As Integer

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    f_funGet_SpreadRow_PAID = 0
    
    With objSpd
        For intRow = 1 To .MaxRows
            .GetText 2, intRow, varTmp
            If Trim$(varTmp) = strPara1 Then
                .GetText 5, intRow, varTmp
                If Trim$(varTmp) = strPara3 Then
                    f_funGet_SpreadRow_PAID = intRow
                    Exit For
                End If
            End If
        Next
    End With
    
End Function


'Private Sub spdResult1_DblClick(ByVal Col As Long, ByVal Row As Long)
'Dim TmpYesno As String
'Dim Tmpptno, TmpPtnm As String
'
'    If Row = 0 Then
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
'            .Col = colPAID
'            Tmpptno = .text
'
'            ' 환자이름 불러오기
'            .Col = colPANM
'            TmpPtnm = .text
'        End With
'
'        If Len(Trim(Tmpptno)) >= 1 And Len(Trim(TmpPtnm)) >= 1 Then
'             TmpYesno = MsgBox(Tmpptno & " (  " & TmpPtnm & "  ) " & " 환자를 선택 하셨습니다..    " & vbCrLf & vbCrLf & "검사를 제외 하시겠습니까..??", vbCritical + vbYesNo, App.Title)
'
'             If TmpYesno = vbYes Then
'                spdResult1.Action = ActionDeleteRow
'                spdResult1.MaxRows = spdResult1.MaxRows - 1
'             End If
'        End If
'    End If
'
'End Sub

'Private Sub spdResult1_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim aCOL, aRow As Integer
'    If KeyCode = vbKeyInsert Then
'        With spdResult1
'            .MaxRows = .MaxRows + 1
'            aCOL = .ActiveCol
'            aRow = .ActiveRow
'            .Action = ActionInsertRow
'
'        End With
'    End If
'
'    If KeyCode = vbKeyDelete Then
'        With spdResult1
'
'            aCOL = .ActiveCol
'            aRow = .ActiveRow
'            .Action = ActionDeleteRow
'            .MaxRows = .MaxRows - 1
'
'        End With
'    End If
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
'                .MaxRows = .MaxRows + 1
'                .Col = Col
'                .Row = Row
'                .Action = ActionInsertRow
'            End With
'        Case 3
'            With spdResult1
'                .Col = Col
'                .Row = Row
'                .Action = ActionDeleteRow
'                .MaxRows = .MaxRows - 1
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
    Dim strDta      As String
    Dim Arr()       As Byte
    Dim sStxCheck As Integer, sEnqCheck As Integer, sEtxCheck As Integer
    Dim sLfCheck As Integer, sCrcheck As Integer, ii As Integer
    Dim MHead As String, Pinfo As String, OutPutData As String, com_sTemp As String
    Dim strRec  As String, strBuff  As String
    
    Dim intIdx     As Integer, intIdx1    As Integer, intIdx2     As Integer
    Dim strTmp1     As String, strTmp2      As String
    Dim intPos1     As Integer, intPos2     As Integer
    Dim intCnt       As Integer
   
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            brStr = ""
            brStr = comEQP.Input
            
            Call ComReceive(brStr)
            
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

Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)

    Dim sTemp      As String
    Dim Channel_No As String       ' 검사항목 번호 : Channel No
    Dim pGrid_Point As Integer
    Dim pDoCount   As Integer
    Dim pDoCount1   As Integer
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
    Dim sSeq, strTmp, varTmp, strBarno, strDate, strTime As String
    Dim sCol As Integer
    Dim sDecnt As Integer
    Dim Float_rate1 As String
    Dim Float_rate2 As String
    Dim Float_rate  As String
    Dim intRow, intIdx As Integer
    Dim i     As Integer
    Dim RcvBuffer As String
    Dim strSPnm As String
    Dim strSPid  As String
    Dim UpdateYN  As Integer
    Dim intpCnt As Integer
    Dim pName   As String
    Dim pNo     As String
    Dim strSeqno As String
    Dim strResult As String
    Dim tmpResult As Variant
    Dim strChannel
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD As String
    Dim strEqpCd As String
    Dim strOrdBuffer As String
    Dim strSndBuffer As String
    Dim strMsgType As String
    Dim strRCQType As String   '0:일반검사,1:CAL, 2:QC
    Dim intOrdCnt As Integer
    Dim strSampleID As String
    Dim blnSend As Boolean
    
    On Error Resume Next
    
    CallForm = "frmInterface - Privete sub psDataDefine()"
    
    RcvBuffer = Mid(strdata, 2)
    
    strTmp = Split(RcvBuffer, vbCr)
    
    pGrid_Point = 0
    f_strBuffer = ""
    
    For intpCnt = 0 To UBound(strTmp)
        Select Case Mid(strTmp(intpCnt), 1, 3)
            
            '검사결과 = ORU^R01   'MSH|^~\&|Mindray|BS-380|||20111215100123||ORU^R01|1|P|2.3.1||||0||ASCII|||
            '오더요청 = QRY^Q02   'MSH|^~\&|Mindray|BS-380|||20111215100123||QRY^Q02|1|P|2.3.1||||0||ASCII|||
            Case "MSH"
                
                strMsgType = mGetP(strTmp(intpCnt), 9, "|")
                
                '-- 0:일반검사,1:CAL, 2:QC
                strRCQType = mGetP(strTmp(intpCnt), 16, "|")
                
        
            '수검자상세정보
            Case "PID"
        
            '오더요청-바코드(123456789) = QRD|2007301193232|R|D|1|||RD|123456789|OTH|||T|
            Case "QRD"
                strSampleID = mGetP(strTmp(intpCnt), 5, "|")
                strBarno = mGetP(strTmp(intpCnt), 9, "|")
                intOrdCnt = 0
                strOrdBuffer = ""
                    
                  '바코드번호로 오더찾아오기
                  Set mAdoRs = f_subSet_WorkList_Barcode(strBarno)
                  
                  If RecordChk = True Then
                      Do Until mAdoRs.EOF
                          intIdx = 0
                          With spdResult1
                              sCol = 5
                              pGrid_Point = SeqSearch(spdResult1, strBarno, sCol)
                              If pGrid_Point = 0 Then
                                  pGrid_Point = SeqNullSearch(spdResult1, strBarno, sCol)
                                  If pGrid_Point = 0 Then
                                      .MaxRows = .MaxRows + 1
                                      .RowHeight(.MaxRows) = 13
                                  End If
                                  pGrid_Point = .MaxRows
                              End If
                              
                              If intOrdCnt = 0 Then
                                  .SetText 1, pGrid_Point, "1"
                                  strSeqno = mAdoRs("SCP42IDNOA")
                                  .SetText 2, pGrid_Point, strSeqno 'mAdoRs("SCP42IDNOA") & ""
                                  .SetText 3, pGrid_Point, mAdoRs("SCP41NAME") & ""
                                  .SetText 4, pGrid_Point, mAdoRs("SCP41JDATE") & ""
                                  .SetText 5, pGrid_Point, strBarno
                                  .SetText 6, pGrid_Point, mAdoRs("SCP42SUGACD") & ""
                              End If
                              
                              For intCol = 10 To .MaxCols
                                  .GetText intCol, 0, varTmp
                                  Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                  If Not itemX Is Nothing Then
                                      If mAdoRs("SCP42SUGACD") & "" = Trim(itemX.SubItems(1)) Then
                                          strOrdBuffer = strOrdBuffer & "DSP|" & 29 + intOrdCnt & "||" & Trim(itemX.Tag) & "^^^|||" & vbCr + vbLf '검사채널(test id)
                                          intOrdCnt = intOrdCnt + 1
                                          spdResult1.Col = itemX.Index + 9
                                          spdResult1.BackColor = &HC6FEFF   '&HC6FEFF

                                          Exit For
                                      End If
                                  End If
                              Next
                              
                          End With
                          mAdoRs.MoveNext
                      Loop
                  Else
                      lblStatus.Caption = "바코드 번호 " & strBarno & " 는 검사대상이 아닙니다"
                      lstStatus.AddItem "[검사오더] 바코드 번호 " & strBarno & " 는 검사대상이 아닙니다"
                  End If
                                                                                                  
                  Set mAdoRs = Nothing
                                 
'                                 strSndBuffer = SB & "MSH|^~\&|Mindray|BS-400|||20100608142704||QCK^Q02|1|P|2.3.1||||0||ASCII|||" & vbCr + vbLf
'                  strSndBuffer = strSndBuffer & "MSA|AA|1|Message accepted|||0|" & vbCr + vbLf
'                  strSndBuffer = strSndBuffer & "ERR|0|" & vbCr + vbLf
                                 strSndBuffer = SB & "MSH|^~\&|Mindray|BS-400|||20100608142704||QCK^Q02|" & strSampleID & "|P|2.3.1||||0||ASCII|||" & vbCr + vbLf
                  strSndBuffer = strSndBuffer & "MSA|AA|" & strSampleID & "|Message accepted|||0|" & vbCr + vbLf
                  strSndBuffer = strSndBuffer & "ERR|0|" & vbCr + vbLf
                  If strOrdBuffer = "" Then
                      strSndBuffer = strSndBuffer & "QAK|SR|NF|" & vbCr + vbLf '-- 오더없을때
                      blnSend = False
                  Else
                      strSndBuffer = strSndBuffer & "QAK|SR|OK|" & vbCr + vbLf '-- 오더있을때
                      blnSend = True
                  End If
                  strSndBuffer = strSndBuffer & Chr(28) & vbCr + vbLf
                  
                  sck.ProcSendMessage strSndBuffer
                  Debug.Print "[오더수신SEND_1]" & strSndBuffer
                  Print #1, vbNewLine & "[오더수신SEND_1]" & strSndBuffer;
                  
                  
'function TfrmMain.Create_DSR_Message(sDateTime, sControlID, sSPECNO : string;
'  var sDsrMsg : string) : Integer;
'Var
'  sQry, sDeviceSPECNO : string;
'  i : Integer;
'begin
'  sDeviceSPECNO := Copy(sSPECNO, 3, 10);
'
'  sDsrMsg :=
'  #$B +
  'MSH|^~\&|||Mindray|BS-400|' + sDateTime + '||DSR^Q03|' + sControlID + '|P|2.3.1||||||ASCII|||' + #$D#$A +
  'MSA|AA|' + sControlID + '|Message accepted|||0|' + #$D#$A +
  'ERR|0|' + #$D#$A +
  'QAK|SR|OK|' + #$D#$A +
  'QRD|' + sDateTime + '|R|D|1|||RD|' + sDeviceSPECNO + '|OTH|||T|' + #$D#$A +
  'QRF|BS-400|' + sDateTime + '|' + sDateTime + '|||RCT|COL|ALL||' + #$D#$A +
  'DSP|1|||||' + #$D#$A +
  'DSP|2|||||' + #$D#$A +
  'DSP|3|||||' + #$D#$A +
  'DSP|4|||||' + #$D#$A +
  'DSP|5|||||' + #$D#$A +
  'DSP|6|||||' + #$D#$A +
  'DSP|7|||||' + #$D#$A +
  'DSP|8|||||' + #$D#$A +
  'DSP|9|||||' + #$D#$A +
  'DSP|10|||||' + #$D#$A +
  'DSP|11|||||' + #$D#$A +
  'DSP|12|||||' + #$D#$A +
  'DSP|13|||||' + #$D#$A +
  'DSP|14|||||' + #$D#$A +
  'DSP|15|||||' + #$D#$A +
  'DSP|16|||||' + #$D#$A +
  'DSP|17|||||' + #$D#$A +
  'DSP|18|||||' + #$D#$A +
  'DSP|19|||||' + #$D#$A +
  'DSP|20|||||' + #$D#$A +
  'DSP|21||' + sDeviceSPECNO + '|||' + #$D#$A +
  'DSP|22|||||' + #$D#$A +
  'DSP|23||' + sDateTime + '|||' + #$D#$A +
  'DSP|24|||||' + #$D#$A +
  'DSP|25|||||' + #$D#$A +
  'DSP|26|||||' + #$D#$A +
  'DSP|27|||||' + #$D#$A +
  'DSP|28|||||' + #$D#$A;
                  
                If blnSend = True Then
                    '                                 strSndBuffer = SB & "MSH|^~\&|||Mindray|BS-400|" & Format(Now, "yyyymmddhhmmss") & "||DSR^Q03|1|P|2.3.1||||||ASCII|||" & vbCr + vbLf
                    '                  strSndBuffer = strSndBuffer & "MSA|AA|1|Message accepted|||0|" & vbCr + vbLf
                                                                      
                                   strSndBuffer = SB & "MSH|^~\&|||Mindray|BS-400|" & Format(Now, "yyyymmddhhmmss") & "||DSR^Q03|" & strSampleID & "|P|2.3.1||||0||ASCII|||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "MSA|AA|" & strSampleID & "|Message accepted|||0|" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "ERR|0|" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "QAK|SR|OK|" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "QRD|" & Format(Now, "yyyymmdd") & "|R|D|1|||RD||OTH|||T|" & "" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "QRF|BS-400|" & Format(Now, "yyyymmddhhmmss") & "|" & Format(Now, "yyyymmddhhmmss") & "|||RCT|COR|ALL||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|1|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|2|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|3|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|4|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|5|||||" & vbCr + vbLf   '성별"
                    strSndBuffer = strSndBuffer & "DSP|6|||||" & vbCr + vbLf  'Blood Type
                    strSndBuffer = strSndBuffer & "DSP|7|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|8|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|9|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|10|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|11|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|12|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|13|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|14|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|15|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|16|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|17|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|18|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|19|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|20|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|21||" & strBarno & "|||" & vbCr + vbLf '바코드
                    '                  strSndBuffer = strSndBuffer & "DSP|22||" & strSampleID & "|||" & vbCr + vbLf   'sample id
                    'strSndBuffer = strSndBuffer & "DSP|22||" & "52" & "|||" & vbCr + vbLf   'sample id
                    strSndBuffer = strSndBuffer & "DSP|22|||||" & vbCr + vbLf   'sample id
                    strSndBuffer = strSndBuffer & "DSP|23||" & Format(Now, "yyyymmddhhmmss") & "|||" & vbCr + vbLf 'sending time
                    strSndBuffer = strSndBuffer & "DSP|24|||||" & vbCr + vbLf  'emergency
                    strSndBuffer = strSndBuffer & "DSP|25|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|26|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|27|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSP|28|||||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & strOrdBuffer
                    'strSndBuffer = strSndBuffer & "DSC||" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & "DSC|" & intOrdCnt - 1 & "|" & vbCr + vbLf
                    strSndBuffer = strSndBuffer & Chr(28) & vbCr + vbLf
                    
                    sck.ProcSendMessage strSndBuffer
                    Debug.Print "[오더수신SEND_2]" & strSndBuffer
                    Print #1, "[오더수신SEND_2]" & strSndBuffer;
                End If
            
            
            '검사환자정보-바코드(123456789) = OBR|1|12345678|2|Mindray^BS-380|N||20111214104203|||||||20111214101918|serum|||||||||||||||||||||||||||||||||
            Case "OBR"
                'Debug.Print strTmp(intpCnt)
                'Debug.Print strRCQType   '0:일반검사,1:CAL, 2:QC
                
                               strSndBuffer = SB & "MSH|^~\&|Mindray|BS-400|||" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|1|P|2.3.1||||0||ASCII|||" & vbCr + vbLf
                strSndBuffer = strSndBuffer & "MSA|AA|1|Message accepted|||0|" & vbCr + vbLf
                strSndBuffer = strSndBuffer & Chr(28) & vbCr + vbLf
'                strSndBuffer = strSndBuffer & "ERR|0|" & vbCr + vbLf

'''                       strSndBuffer = Chr(11) & "MSH|^~\&|Mindray|BS-400|||20100608142704||ACK^R01|1|P|2.3.1||||0||ASCII|||" & Chr(13)
'''        strSndBuffer = strSndBuffer & "MSA|AA|1|Message accepted|||0|" & Chr(13)
'''        strSndBuffer = strSndBuffer & Chr(28) & vbCr

                sck.ProcSendMessage strSndBuffer
                Debug.Print "[결과수신SEND]" & strSndBuffer
                Print #1, vbNewLine & "[결과수신SEND]" & strSndBuffer;
                lstStatus.AddItem "[OBR 결과수신]"
                
                strBarno = mGetP(strTmp(intpCnt), 3, "|")
                strSeqno = mGetP(strTmp(intpCnt), 4, "|")
                'strSampleID = mGetP(strTmp(intpCnt), 4, "|")
                
                If strBarno = "" Then
                    strBarno = strSeqno
                End If
                
                If strRCQType = "2" Then
                    strBarno = ""
                    'strBarno = "QC_" & strBarno & "_" & strSeqno
                End If
                
                'strBarno = "1012006793"
                If strBarno <> "" Then
                    pGrid_Point = SeqSearch(spdResult1, strBarno, colPAID)
                    If pGrid_Point = 0 Then
                        pGrid_Point = SeqNullSearch(spdResult1, strBarno, colPAID)
                    End If
                Else
                    pGrid_Point = SeqNullSearch(spdResult1, strBarno, colPAID)
                End If
                
                If pGrid_Point = 0 Then
                    spdResult1.MaxRows = spdResult1.MaxRows + 1
                    pGrid_Point = spdResult1.MaxRows
                End If
                
                '바코드번호로 오더찾아오기
                Set mAdoRs = f_subSet_WorkList_Barcode(strBarno)
                
                If RecordChk = True Then
                    Do Until mAdoRs.EOF
                        intIdx = 0
                        With spdResult1
                            sCol = 5
                            pGrid_Point = SeqSearch(spdResult1, strBarno, colPAID)
                            If pGrid_Point = 0 Then
                                pGrid_Point = SeqNullSearch(spdResult1, strBarno, colPAID)
                                If pGrid_Point = 0 Then
                                    .MaxRows = .MaxRows + 1
                                    .RowHeight(.MaxRows) = 13
                                End If
                                pGrid_Point = .MaxRows
                            End If
                            
                            .SetText 1, pGrid_Point, "1"
                            .SetText 2, pGrid_Point, mAdoRs("SCP42IDNOA") & ""
                            .SetText 3, pGrid_Point, mAdoRs("SCP41NAME") & ""
                            .SetText 4, pGrid_Point, mAdoRs("SCP41JDATE") & ""
                            .SetText 5, pGrid_Point, strBarno
                            .SetText 6, pGrid_Point, mAdoRs("SCP42SUGACD") & ""
                        End With
                        
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
                    lstStatus.AddItem "[검사결과] 바코드 번호 " & strBarno & " 는 검사대상이 아닙니다"
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
        
            '검사결과정보
    '        OBX|1|NM|TP|TP|7.625771|g/dl|6.500000-8.000000|N|||F||7.625771|20111214104203||||
    '        OBX|2|NM|ALB|ALB|4.548543|g/dl|3.800000-5.300000|N|||F||4.548543|20111214104203||||
    '        OBX|3|NM|GOT|GOT|8.429273|g/L|12.000000-33.000000|L|||F||8.429273|20111214104203||||
    '        OBX|4|NM|GPT|GPT|6.162995|U/L|5.000000-35.000000|N|||F||6.162995|20111214104203||||
    '        OBX|5|NM|r-GTP|r-GTP|123.405985|U/L|11.000000-73.000000|H|||F||123.405985|20111214104203||||
            
            Case "OBX"
                Channel_No = mGetP(strTmp(intpCnt), 4, "|")
                strResult = mGetP(strTmp(intpCnt), 6, "|")
            
                If pGrid_Point > 0 And RecordChk = True And Len(strBarno) = 10 Then
                    For intCol = 10 To spdResult1.MaxCols
                        spdResult1.GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            spdResult1.Row = pGrid_Point
                            spdResult1.Col = intCol
                            If Channel_No = Trim(itemX.Tag) And spdResult1.BackColor = &HC6FEFF And spdResult1.BackColor <> vbMagenta Then
                                Select Case Channel_No
                                   Case "105", "106", "107", "108", "109", "111", "120", "121", "122"
                                       strRstval = Format(strResult, "###,##0.0")
                                   Case "119"
                                       strRstval = Format(strResult, "###,##0.00")
                                   Case Else
                                       strRstval = Format(strResult, "##0")
                                End Select

                                strEqpCd = itemX.text
                                
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
                                         "   and EQPNUM  = '" & itemX.Tag & "'" & _
                                         "   and TRANSDT = '" & strDate & "'" & _
                                         "   and TRANSTM = '" & strTime & "'"
                                AdoCn_Jet.Execute sqlDoc
                                
                                sqlDoc = "insert into INTERFACE003(" & _
                                         "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN, NAME, PNO)" & _
                                         "    values( '" & strBarno & "', '" & strEqpCd & "', '" & itemX.Tag & "'," & _
                                         "            '" & strDate & "', '" & strTime & "'," & _
                                         "            '" & strRstval & "', '" & strRefVal & "'," & _
                                         "            '" & INS_CODE & "', '', '" & pName & "', '" & pNo & "')"
                                AdoCn_Jet.Execute sqlDoc
                                
                                If chkAuto.Value = "1" And Len(strEqpCd) <> 0 Then
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
                                    
                                    spdResult1.Row = pGrid_Point
                                    
                                    spdResult1.Col = colORDT:   spdResult1.BackColor = vbCyan
                                    spdResult1.Col = colORQN:   spdResult1.BackColor = vbCyan
                                    spdResult1.Col = colPANM:   spdResult1.BackColor = vbCyan
                                    spdResult1.Col = colPAID:   spdResult1.BackColor = vbCyan
                                    spdResult1.Col = colOIFL:   spdResult1.BackColor = vbCyan
                                    spdResult1.Col = colSENO:   spdResult1.BackColor = vbCyan
                                    spdResult1.Col = colSEXS:   spdResult1.BackColor = vbCyan
                                    spdResult1.Col = colSENO:   spdResult1.BackColor = vbCyan
                                    'spdResult1.Col = colAGES:   spdResult1.BackColor = vbCyan
                                    'spdResult1.Col = colNWNO:   spdResult1.BackColor = vbCyan
                                    spdResult1.Col = colBANO:   spdResult1.Value = 0
                                    
                                    spdResult1.Col = intCol
                                    spdResult1.BackColor = vbMagenta
                                    
                                    lstStatus.AddItem "[OBX 결과수신] " & strBarno & "-" & strEqpCd & "-" & strRstval

                                End If
                                Set itemX = Nothing
                                
                                Exit For

                            End If
                        End If
                        Set itemX = Nothing
                    Next
                    
                    
                End If

                
        End Select
    Next
        
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Public Sub ComReceive(ByRef RecData As String)
                
    
    Print #1, RecData;
    
    Call psDataDefine(RecData, fChannel(), spdResult1)
    
    
End Sub

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
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

Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Val(.text) = brSeq Then
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub Command1_Click()
    If sck.state <> "Connected" Then Exit Sub
    
    sck.ProcSendMessage "테스트 데이타입니다."
    
End Sub

Private Sub Form_Activate()

    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
    
'    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    
    Call cmdClear               ' 초기화
    Call f_subSet_ItemHeader    ' 리스트해더
    Call f_subSet_ItemList      ' 검사항목
    Call f_subGet_Setting       ' 통신설정
    
    Call cmdRun                 ' 실행
    
    mskRstDate.text = Format$(Now, "YYYYMMDD")
    mskOrdDate.text = Format$(Now - 1, "YYYYMMDD")
    mskOrdDate1.text = Format$(Now, "YYYYMMDD")
    
    Open App.Path + "\Log\" + REG_INSNAME + "_" + Format(Now, "YYYYMMDD") + ".Log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    Open App.Path + "\ErrorLog\" + REG_INSNAME + "_" + Format(Now, "YYYYMMDD") + ".sql" For Append As #2

    Print #2, Chr(13) + Chr(10);
   
    tabWork.Tab = 0
    
    Call cmdAction_Click(4)
    
End Sub

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
    
    Close #1
    Close #2
    
End Sub


Private Sub Label6_DblClick()
    If Frame4.Visible = False Then
        Frame4.Visible = True
    Else
        Frame4.Visible = False
    End If
End Sub

Private Sub mskOrdDate_GotFocus()

    With mskOrdDate
        .SelStart = 8
        .SelLength = Len(.text)
    End With
    
End Sub

Private Sub mskOrdDate_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskOrdDate.SelLength = 1
    
End Sub

Private Sub mskRstDate_GotFocus()

    With mskRstDate
        .SelStart = 0
        .SelLength = Len(.text) + 2
    End With '
    
End Sub

Private Sub mskRstDate_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskRstDate.SelLength = 1
    
End Sub

Private Sub spdResult1_KeyPress(KeyAscii As Integer)

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


Private Sub spdRstview_Click(ByVal Col As Long, ByVal Row As Long)

Dim iCnt, rCnt As Integer
Dim intCol, intRow As Integer
Dim tCol As Integer
Dim iresult As String
'
' 결과 시작 Position
'
Const sResultPos As Integer = 8

    With spdRstview
        For iCnt = 2 To .MaxCols Step 2
            For rCnt = 1 To .MaxRows
                .Row = rCnt: .Col = iCnt
                iresult = Trim(.text)
                
                With spdResult1
                    .Row = gspdResultRow:  .Col = sResultPos + tCol
                    If Len(Trim(iresult)) <> 0 Then
                        .text = iresult
                    End If
                    DoEvents
                End With
                tCol = tCol + 1
                
            Next rCnt
            rCnt = 0
        Next iCnt
    End With

End Sub

Private Sub spdRstview_EnterRow(ByVal Row As Long, ByVal RowIsLast As Long)
    Call spdRstview_Click(Row, RowIsLast)
End Sub

Private Sub spdRstview_KeyPress(KeyAscii As Integer)

Dim iCnt, rCnt As Integer
Dim intCol, intRow As Integer
Dim tCol As Integer
Dim iresult As String

'
' 결과 시작 Position
'
Const sResultPos As Integer = 8
     
    ' 처방 존재 유무 확인..
    With spdRstview
        .Row = .ActiveRow: .Col = .ActiveCol
        If .BackColor <> &HC6FEFF And Len(.text) >= 1 Then
            .text = ""
            MsgBox "▒ OCS/EMR의 검사 처방이 없는 항목 입니다.." & Space(5), vbOKOnly + vbInformation, App.Title
            spdRstview.SetFocus
            Exit Sub
        End If
    End With
    
    ' Enter Key 유무..
    If KeyAscii = vbKeyReturn Then
    
        If gspdResultRow < 1 Then
            With spdRstview
                .Row = .ActiveRow:  .Col = .ActiveCol
                .text = ""
            End With
            
            MsgBox "▒ 수정을 원하는 검사 Sample을 선택 후 수정 하십시요.." & Space(5), vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        
        ' 수정된 결과 본 Spread로 옮기기..
        With spdRstview
            For iCnt = 2 To .MaxCols Step 2
                For rCnt = 1 To .MaxRows
                    .Row = rCnt: .Col = iCnt
                    iresult = Trim(.text)
                    
                    With spdResult1
                        .Row = gspdResultRow:  .Col = sResultPos + tCol
                        If Len(Trim(iresult)) <> 0 Then
                            .text = iresult
                        End If
                        DoEvents
                    End With
                    tCol = tCol + 1
                    
                Next rCnt
                rCnt = 0
            Next iCnt
        End With
    End If

End Sub


Private Sub cmdNext_Click()
Dim Col, Row As Integer
    
    With spdResult1
        Col = .ActiveCol
        Row = .ActiveRow
    End With
    
    If gspdResultRow < spdResult1.MaxRows Then
        Call spdResult1_Click(Col, gspdResultRow + 1)
    ElseIf gspdResultRow = spdResult1.MaxRows Then
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
    
    If Row = 0 Then
        gspdResultRow = 0:        Exit Sub
    Else
        gspdResultRow = Row
    End If
    
    intCol1 = 10
    intCol2 = 2
    intRow1 = 1
    
    With spdResult1
        For iCnt = intCol1 To .MaxCols
            .Row = Row
            .Col = intCol1
            
            spdRstview.Row = intRow1
            spdRstview.Col = intCol2
            spdRstview.BackColor = vbWhite
            
            If .BackColor = &HC6FEFF Then
                spdRstview.BackColor = &HC6FEFF
            Else
                spdRstview.BackColor = &H80000005
            End If
            
            spdRstview.text = .text
            
            intRow1 = intRow1 + 1
            intCol1 = intCol1 + 1
            
            If intRow1 > spdRstview.MaxRows Then
                intRow1 = 1
                intCol2 = intCol2 + 2
            End If

        Next
    End With
    

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

    lblStatus.Caption = "Socket :: " & sck.state & Space(1) & "IP :: " & sck.StateConnIP & Space(1) & "Port :: " & sck.StateConnPort
       
End Sub

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
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    sck.Accept requestID
    Winsock1.Close
    Winsock1.Listen
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strRcvBuffer As String
    Dim strSndBuffer As String
   
    imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If tmrReceive.Enabled = False Then
        tmrReceive.Enabled = True
    Else
        tmrReceive.Enabled = False
        tmrReceive.Enabled = True
    End If
    
    Winsock1.GetData strRcvBuffer
    Debug.Print strRcvBuffer


    strSndBuffer = "ORDER"
    Winsock1.SendData (strSndBuffer)


End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Number & " >> " & Description
    Winsock1.Close
End Sub

