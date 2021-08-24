VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
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
      TabPicture(0)   =   "frmComm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "spdWorklist"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "spdRstview"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdStartNo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAppend(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SSPanel1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdWorkList"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdNext"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdPrevious"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "spdResult1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkAuto"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtSeqNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   " ▒   받은 결과     "
      TabPicture(1)   =   "frmComm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPanel2"
      Tab(1).Control(1)=   "cmdAppend(1)"
      Tab(1).Control(2)=   "cmdExcel"
      Tab(1).Control(3)=   "chkExcel"
      Tab(1).Control(4)=   "spdResult2"
      Tab(1).Control(5)=   "cmdSel(2)"
      Tab(1).Control(6)=   "cmdSel(3)"
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtSeqNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   270
         Left            =   7800
         TabIndex        =   48
         Top             =   510
         Width           =   825
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
         Left            =   13710
         TabIndex        =   37
         Top             =   30
         Value           =   1  '확인
         Width           =   1320
      End
      Begin VB.Frame Frame4 
         Caption         =   "hidden"
         Height          =   6225
         Left            =   7170
         TabIndex        =   35
         Top             =   1950
         Visible         =   0   'False
         Width           =   8055
         Begin VB.Timer Timer2 
            Left            =   6510
            Top             =   1860
         End
         Begin VB.ListBox List1 
            Height          =   2580
            ItemData        =   "frmComm.frx":0038
            Left            =   150
            List            =   "frmComm.frx":003A
            TabIndex        =   45
            Top             =   3480
            Width           =   7215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "TEST"
            Height          =   375
            Left            =   6030
            TabIndex        =   38
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
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   5940
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
                  Picture         =   "frmComm.frx":003C
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":05D6
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":0B70
                  Key             =   "NOF"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":110A
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":16A4
                  Key             =   "LSE"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":1C3E
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
                  Picture         =   "frmComm.frx":21D8
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":2772
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":2D0C
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":32A6
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":3B38
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":3C92
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmComm.frx":3DEC
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvwCuData 
            Height          =   3000
            Left            =   150
            TabIndex        =   36
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
            TabIndex        =   46
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
         Picture         =   "frmComm.frx":3F46
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
         Picture         =   "frmComm.frx":43C8
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   7320
         Left            =   -74910
         TabIndex        =   33
         Top             =   900
         Width           =   15045
         _Version        =   196608
         _ExtentX        =   26538
         _ExtentY        =   12912
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
         MaxCols         =   11
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":4836
         UserResize      =   2
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   7620
         Left            =   4920
         TabIndex        =   32
         Top             =   900
         Width           =   10245
         _Version        =   196608
         _ExtentX        =   18071
         _ExtentY        =   13441
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   11
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":4EB1
         UserResize      =   2
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
         TransparentPicture=   "frmComm.frx":5539
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
            Picture         =   "frmComm.frx":59AB
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
            Picture         =   "frmComm.frx":5E2D
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
            TabIndex        =   34
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
            Caption         =   "처방일자 :"
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
         Height          =   420
         Index           =   0
         Left            =   13650
         TabIndex        =   16
         Top             =   400
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   741
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
      Begin FPSpread.vaSpread spdRstview 
         Height          =   2865
         Left            =   90
         TabIndex        =   23
         Top             =   5640
         Width           =   4725
         _Version        =   196608
         _ExtentX        =   8334
         _ExtentY        =   5054
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
         MaxRows         =   8
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   0
         ShadowColor     =   14735310
         SpreadDesigner  =   "frmComm.frx":629B
         UserResize      =   0
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   525
         Left            =   -74910
         TabIndex        =   39
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
            ItemData        =   "frmComm.frx":6D74
            Left            =   2235
            List            =   "frmComm.frx":6D81
            Style           =   2  '드롭다운 목록
            TabIndex        =   42
            Top             =   135
            Width           =   1410
         End
         Begin MSMask.MaskEdBox mskRstDate 
            Height          =   300
            Left            =   1110
            TabIndex        =   43
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
            TabIndex        =   44
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
            TabIndex        =   41
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "분 접수까지."
            Height          =   255
            Left            =   5520
            TabIndex        =   40
            Top             =   840
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin FPSpread.vaSpread spdWorklist 
         Height          =   4230
         Left            =   90
         TabIndex        =   47
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
         SpreadDesigner  =   "frmComm.frx":6DAB
         UserResize      =   2
      End
      Begin VB.Label Label4 
         Caption         =   "Last Seq :"
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
         Left            =   6660
         TabIndex        =   49
         Top             =   540
         Width           =   1095
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
         TransparentPicture=   "frmComm.frx":73AA
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
         Picture         =   "frmComm.frx":8C34
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13725
         Picture         =   "frmComm.frx":91BE
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12525
         Picture         =   "frmComm.frx":9748
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
Dim strBuffer As String

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
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList() As ADODB.Recordset"
   
    Set AdoRs_ORACLE = New ADODB.Recordset
    
                '-- 처방일자,처방일련번호,환자명,환자번호,입외구분,일련번호,성별,나이,내원번호,처방코드
             sqlDoc = "Select a.ORDT,a.ORQN,b.PANM,a.PAID,a.OIFL,a.SENO,b.SEXS,b.AGES,a.NWNO,a.ORCD "
    sqlDoc = sqlDoc & "  From LRESULT a, APATINF b"
    sqlDoc = sqlDoc & " Where a.ORDT between  '" & mskOrdDate.text & "' and '" & mskOrdDate1.text & "'"
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
    Dim stryy, strmm, strdd, strDate  As String
    
On Error GoTo ErrorTrap
    CallForm = "clsCommon - Public Function f_subSet_WorkList_Barcode() As ADODB.Recordset"
    
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
                '-- 처방일자,처방일련번호,환자명,환자번호,입외구분,일련번호,성별,나이,내원번호,처방코드
             sqlDoc = "Select a.ORDT,a.ORQN, b.PANM,a.PAID,a.OIFL,a.SENO,b.SEXS,b.AGES,a.NWNO,a.ORCD "
    sqlDoc = sqlDoc & "  From LRESULT a, APATINF b"
    sqlDoc = sqlDoc & " Where a.ORDT =  '" & strORDT & "'"
    sqlDoc = sqlDoc & "   And a.PAID =  '" & strPAID & "'"
    sqlDoc = sqlDoc & "   And a.PAID = b.PAID "
    sqlDoc = sqlDoc & "   And a.ORCD in (" & strGumCd & ")"
    sqlDoc = sqlDoc & "   And a.OKFL <> 'Y' "   '-- 결과확정유무

    Set AdoRs_ORACLE = New ADODB.Recordset
    
    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open sqlDoc, AdoCn_ORACLE
    
    If AdoRs_ORACLE.RecordCount = 0 Then
        Set f_subSet_WorkList_Barcode = Nothing
        RecordChk = False
        Set AdoRs_ORACLE = Nothing
        Exit Function
    Else
        Set f_subSet_WorkList_Barcode = AdoRs_ORACLE
        RecordChk = True
    End If

    Set AdoRs_ORACLE = Nothing
    
Exit Function

ErrorTrap:
    Set AdoRs_ORACLE = Nothing
    
    Call ErrMsgProc(CallForm)
     
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
    If strAGE = "소아" Then
        sqlDoc = "Select YMAX as MAX, YMIN as MIN "
    Else
        If strSex = "M" Then
                 sqlDoc = "Select MMAX as MAX, MMIN as MIN "
        Else
                 sqlDoc = "Select WMAX as MAX, WMIN as MIN "
        End If
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
        If strRSLT > AdoRs_ORACLE.Fields("MAX") Then
            f_subSet_RefVal = "H"
        ElseIf strRSLT < AdoRs_ORACLE.Fields("MIN") Then
            f_subSet_RefVal = "L"
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
    
    intCol = 12
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
        
        fChannel(intCol - colSQNO) = AdoRs.Fields("TEST_EQP")
        
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
    Dim strBarno As String
    Dim strSPnm As String
    Dim strSPid As String
    Dim strChartNo As String
    Dim strEqpCd As String
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD As String
    Dim strRefVal As String
    
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
            .GetText colPAID, intRow, varTmp:    strPAID = Trim$(varTmp)
            .GetText colOIFL, intRow, varTmp:    strOIFL = Trim$(varTmp)
            .GetText colSENO, intRow, varTmp:    strSENO = Trim$(varTmp)
            .GetText colSEXS, intRow, varTmp:    strSEXS = Trim$(varTmp)
            .GetText colAGES, intRow, varTmp:    strAGES = Trim$(varTmp)
            .GetText colNWNO, intRow, varTmp:    strNWNO = Trim$(varTmp)

            .GetText colBANO, intRow, varTmp

            If strORDT = "" Then Exit For

            intCnt = 0: Erase strOrdcd ': Erase strRstval
            
            If Trim$(varTmp) = "1" Then
                For intCol = colSQNO + 1 To .MaxCols
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
                                    '-- H/L 판정
                                    'strRefVal = f_subSet_RefVal(strEqpCd, strRstval, strSEXS, strAGES)
                                    strRefVal = " "
                                    
                                    '-- 최근 접수번호[ORQN] 찾기
                                    sqlDoc = "Select ORQN,TRANSDT,TRANSTM from INTERFACE004"
                                    sqlDoc = sqlDoc & " Where ORDT = '" & strORDT & "'"
                                    sqlDoc = sqlDoc & "   And PAID = '" & strPAID & "'"
                                    sqlDoc = sqlDoc & "   And OIFL = '" & strOIFL & "'"
                                    sqlDoc = sqlDoc & "   And EQPCD = '" & strEqpCd & "'"
                                    sqlDoc = sqlDoc & " Order By TRANSDT, TRANSTM desc "
                                    
                                    AdoRs.CursorLocation = adUseClient
                                    AdoRs.Open sqlDoc, AdoCn_Jet
                                    If AdoRs.RecordCount > 0 Then
                                        AdoRs.MoveFirst
                                    End If
                                    
                                    Do While Not AdoRs.EOF
                                        Debug.Print AdoRs.Fields("TRANSDT") & AdoRs.Fields("TRANSTM")
                                        If Trim(AdoRs.Fields("ORQN")) <> "" Then
                                            strORQN = Trim(AdoRs.Fields("ORQN"))
                                            Exit Do
                                        End If
                                        AdoRs.MoveNext
                                    Loop
                                    
                                    Set AdoRs = Nothing
                                    
                                    '-- 서버저장
                                    sqlDoc = " Update LRESULT"
                                    sqlDoc = sqlDoc & "   Set RSFL = 'Y',"
                                    sqlDoc = sqlDoc & "       RSLT = '" & strRstval & "',"
                                    sqlDoc = sqlDoc & "       HLFL = '" & strRefVal & "',"
                                    sqlDoc = sqlDoc & "       RSDT = '" & Format(Now, "YYYYMMDD") & "',"
                                    sqlDoc = sqlDoc & "       RSID = '" & CurrUser.CuUserID & "'"
                                    sqlDoc = sqlDoc & " Where PAID = '" & strPAID & "'"
                                    sqlDoc = sqlDoc & "   And NWNO = " & strNWNO
                                    sqlDoc = sqlDoc & "   And ORDT = '" & strORDT & "'"
                                    sqlDoc = sqlDoc & "   And ORQN = " & strORQN
                                    sqlDoc = sqlDoc & "   And OIFL = '" & strOIFL & "'"
                                    sqlDoc = sqlDoc & "   And ORCD = '" & strEqpCd & "'"
                                    sqlDoc = sqlDoc & "   And OKFL <> 'Y' "   '-- 결과확정유무
                                    
                                    AdoCn_ORACLE.Execute sqlDoc
                                    
                                    '-- 로컬 업데이트
                                    sqlDoc = "Update INTERFACE004" & _
                                             "   Set SERVERGBN = 'Y'" & _
                                             " Where PAID = '" & strPAID & "'" & _
                                             "   And NWNO = " & strNWNO & _
                                             "   And ORDT = '" & strORDT & "'" & _
                                             "   And ORQN = " & strORQN & _
                                             "   And OIFL = '" & strOIFL & "'" & _
                                             "   And EQPCD = '" & strEqpCd & "'"
                                                                                                     
                                    AdoCn_Jet.Execute sqlDoc
                                    
                                    Debug.Print sqlDoc
                                    
                                    
                                    lblStatus.Caption = "저장 성공!!"
                                                                                                                    
                                    .Row = intRow: .Col = colBANO: .Value = 0
                                                   .Col = colORDT: .BackColor = HNC_Cyan
                                                   .Col = colORQN: .BackColor = HNC_Cyan
                                                   .Col = colPANM: .BackColor = HNC_Cyan
                                                   .Col = colPAID: .BackColor = HNC_Cyan
                                                   .Col = colOIFL: .BackColor = HNC_Cyan
                                                   .Col = colSENO: .BackColor = HNC_Cyan
                                                   .Col = colSEXS: .BackColor = HNC_Cyan
                                                   .Col = colAGES: .BackColor = HNC_Cyan
                                                   .Col = colNWNO: .BackColor = HNC_Cyan
                                                                        
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
    
    Select Case Index
        Case 0:     Call cmdRun
        Case 1:     Call cmdStop
        Case 2:     Call cmdClear
        Case 3:     Call cmdExit
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

    Dim AdoRs   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    
    Dim itemX       As ListItem

    intRow = 0
    With spdResult2
        .MaxRows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
                                                             
    sqlDoc = "Select ORDT, ORQN, PAID, OIFL, SENO, NWNO, ORCD, TRANSDT, TRANSTM, EqpCD, RSTVAL, REFVAL, SERVERGBN, PANM, SEX, AGE,EQPNUM " & _
             "  From INTERFACE004" & _
             " Where TRANSDT >= '" & mskRstDate.text & "'"
             
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   And (SERVERGBN = '' or SERVERGBN = 'N')"
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   And SERVERGBN = 'Y'"
    End If
    sqlDoc = sqlDoc & " Order By ORDT, TRANSDT,TRANSTM"
    
    AdoRs.CursorLocation = adUseClient
    AdoRs.Open sqlDoc, AdoCn_Jet
    If AdoRs.RecordCount > 0 Then AdoRs.MoveFirst
    Do While Not AdoRs.EOF
        With spdResult2
            If strSpcno <> Trim$(AdoRs("ORDT") & "") & Trim$(AdoRs("ORQN") & "") & Trim$(AdoRs("PAID") & "") Then
                intRow = intRow + 1
                If intRow > .MaxRows Then .MaxRows = .MaxRows + 1:  .RowHeight(.MaxRows) = 13
                .SetText 1, intRow, "1"
                .SetText 2, intRow, AdoRs("ORDT").Value & ""
                .SetText 3, intRow, AdoRs("ORQN").Value & ""
                .SetText 4, intRow, Trim(AdoRs("PANM").Value & "")
                .SetText 5, intRow, Trim(AdoRs("PAID").Value & "")
                .SetText 6, intRow, Trim(AdoRs("OIFL").Value & "")
                .SetText 7, intRow, Trim(AdoRs("SENO").Value & "")
                .SetText 8, intRow, Trim(AdoRs("SEX").Value & "")
                .SetText 9, intRow, Trim(AdoRs("AGE").Value & "")
                .SetText 10, intRow, Trim(AdoRs("NWNO").Value & "")
                .SetText 11, intRow, Format(Trim(AdoRs("TRANSDT").Value & ""), "####-##-##")
            End If
            strSpcno = Trim$(AdoRs("ORDT") & "") & Trim$(AdoRs("ORQN") & "") & Trim$(AdoRs("PAID") & "")
            Set itemX = lvwCuData.FindItem(Trim$(AdoRs("EQPNUM") & ""), lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then
                intCol = itemX.Index + 11
                .SetText intCol, intRow, Trim$(AdoRs("RSTVAL")) & ""
                .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(AdoRs("REFVAL") & "") <> "", vbRed, vbBlack)
            End If
        End With
        AdoRs.MoveNext
    Loop
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
                txtSeqNo.text = .text
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
            
        End With
    End If

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
                .GetText 3, intRow, varTmp
                If Trim$(varTmp) = strPara2 Then
                    .GetText 5, intRow, varTmp
                    If Trim$(varTmp) = strPara3 Then
                        f_funGet_SpreadRow_PAID = intRow
                        Exit For
                    End If
                End If
            End If
        Next
    End With
    
End Function

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

Private Sub spdResult1_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim TmpYesno As String
Dim Tmpptno, TmpPtnm As String

    If Row = 0 Then
        If Col = 1 Then
            Col = 2
        End If
        
        If OrderSort_Flag = 1 Then
            Call SpreadSheetSort(spdResult1, Col, 2)
            OrderSort_Flag = 2
        Else
            Call SpreadSheetSort(spdResult1, Col, 1)
            OrderSort_Flag = 1
        End If
        
        Exit Sub
    End If


    If Col = 4 Or Col = 6 Then
        With spdResult1
            .Row = Row
            
            ' 병록번호 불러오기
            .Col = colPAID
            Tmpptno = .text
            
            ' 환자이름 불러오기
            .Col = colPANM
            TmpPtnm = .text
        End With
        
        If Len(Trim(Tmpptno)) >= 1 And Len(Trim(TmpPtnm)) >= 1 Then
             TmpYesno = MsgBox(Tmpptno & " (  " & TmpPtnm & "  ) " & " 환자를 선택 하셨습니다..    " & vbCrLf & vbCrLf & "검사를 제외 하시겠습니까..??", vbCritical + vbYesNo, App.Title)
        
             If TmpYesno = vbYes Then
                spdResult1.Action = ActionDeleteRow
                spdResult1.MaxRows = spdResult1.MaxRows - 1
             End If
        End If
    End If
        
End Sub

Private Sub spdResult1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim aCOL, arow As Integer
    If KeyCode = vbKeyInsert Then
        With spdResult1
            .MaxRows = .MaxRows + 1
            aCOL = .ActiveCol
            arow = .ActiveRow
            .Action = ActionInsertRow
            
        End With
    End If
    
    If KeyCode = vbKeyDelete Then
        With spdResult1

            aCOL = .ActiveCol
            arow = .ActiveRow
            .Action = ActionDeleteRow
            .MaxRows = .MaxRows - 1
            
        End With
    End If
End Sub

Private Sub spdResult1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
    
    Set oMenu = New cPopupMenu
    
    lMenuChosen = oMenu.Popup(" ▒ 검사자 추가", "-", " ▒ 검사자 삭제", "-", " ▒ 시작번호수정", "-", " ▒ 서버 저장")

    Select Case lMenuChosen
        Case 1
            With spdResult1
                .MaxRows = .MaxRows + 1
                .Col = Col
                .Row = Row
                .Action = ActionInsertRow
            End With
        Case 3
            With spdResult1
                .Col = Col
                .Row = Row
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
            End With
        Case 5
            Call cmdStartNo_Click
        Case 7
            Call cmdAppend_Click(0)
    End Select
End Sub

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
'    Dim intIdx As Long
    
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            'brStr = ""
            brStr = comEQP.Input
'            brStr = strBuffer
            
'            Call ComReceive(brStr)
            
            For intIdx = 1 To Len(brStr)
                strBuff = Mid$(brStr, intIdx, 1)
                'strBuffLen = Mid$(strRec, intIdx, 2)
                Select Case strBuff
                    Case "~" '-- STX & ETX
                        If f_strBuffer <> "" And Len(f_strBuffer) > 50 Then
                            Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
                            f_strBuffer = ""
                            Timer2.Interval = 5000
                            Timer2.Enabled = True
                            Exit For
                        End If
                        f_strBuffer = f_strBuffer + strBuff
                    Case Else
                        f_strBuffer = f_strBuffer + strBuff
                End Select
             Next
        
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


Private Sub psDataDefine(ByVal strdata As String, ByRef brChannel() As String, ByVal brspread As Object) ', ByVal brOst As String) ' ByRef brItemdeci() As String)
    Dim strEqpCd As String
    Dim strOrderMsg As String
    Dim itemX   As ListItem
    Dim pGrid_Point As Integer
    Dim varTmp
    Dim strBarno As String
    Dim pName As String
    Dim pNo As String
    Dim intCol0 As Integer
    Dim intCol As Integer
    Dim strRstval As String
    Dim intIdx As Integer
    Dim TestId As String
    Dim Channel_No As String
    Dim strRefVal As String
    Dim strDate As String
    Dim strTime As String
    Dim sqlDoc As String
    Dim sSeq As String
    Dim intCnt As Integer
    Dim intOrdCnt As Integer
    Dim strTmpBar1 As String
    Dim sCol As Integer
    Dim strTmpDate As String
    Dim ReceiveData As String
    
    Dim strTmpBar As String
    
    Dim sqlRet   As Integer
    Dim stryy, strmm, strdd As String
    Dim mResult As Variant
    Dim mIcount As Integer
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD As String
    
    On Error GoTo ErrReceive
    
     ReceiveData = strdata
     mResult = Split(strdata, vbCrLf)
     
'     For mIcount = 0 To UBound(mResult)
'        Debug.Print mIcount, " : ", mResult(mIcount)
'     Next mIcount
     
'     If InStr(mResult(0), "F") = 0 Then
        LabUReader.SID = Mid(Trim(mResult(3)), 8)
        LabUReader.SID = Replace(LabUReader.SID, "(", "")
        LabUReader.SID = Replace(LabUReader.SID, ")", "")
'     Else
'         LabUReader.SID = Mid(Trim(mResult(4)), 8)
'         LabUReader.SID = Replace(LabUReader.SID, "(", "")
'         LabUReader.SID = Replace(LabUReader.SID, ")", "")
'     End If
                     
     List1.AddItem ("▒ LabUReader.Sample Position Number : " & Val(Mid(LabUReader.SID, 1)))
    
     strTmpDate = Format(Now, "YYYY")
     ReceiveData = Mid(ReceiveData, 50)
     intOrdCnt = 0
     
     For intCnt = 4 To UBound(mResult)
         LabUReader.TestId(intCnt) = Mid(mResult(intCnt), 1, 3)
        If UCase(Trim(LabUReader.TestId(intCnt))) = "PH" Or UCase(LabUReader.TestId(intCnt)) = "S.G" Then
            LabUReader.Result(intCnt) = Trim(Mid(mResult(intCnt), 12))
        Else
            LabUReader.Result(intCnt) = Trim(Mid(mResult(intCnt), 4, 10))
        End If
     Next
     
     If LabUReader.SID <> "" Then
         With spdResult1
             pGrid_Point = SeqSearch(spdResult1, Val(Trim(LabUReader.SID)), colSQNO)
             If pGrid_Point = 0 Then
                 List1.AddItem "▒ 대상 번호 없음"
                Exit Sub
             End If
             .GetText colORDT, pGrid_Point, varTmp:    strORDT = Trim$(varTmp)
             .GetText colORQN, pGrid_Point, varTmp:    strORQN = Trim$(varTmp)
             .GetText colPANM, pGrid_Point, varTmp:    strPANM = Trim$(varTmp)
             .GetText colPAID, pGrid_Point, varTmp:    strPAID = Trim$(varTmp)
             .GetText colOIFL, pGrid_Point, varTmp:    strOIFL = Trim$(varTmp)
             .GetText colSENO, pGrid_Point, varTmp:    strSENO = Trim$(varTmp)
             .GetText colSEXS, pGrid_Point, varTmp:    strSEXS = Trim$(varTmp)
             .GetText colAGES, pGrid_Point, varTmp:    strAGES = Trim$(varTmp)
             .GetText colNWNO, pGrid_Point, varTmp:    strNWNO = Trim$(varTmp)
             
             List1.AddItem ("▒ " & strORQN & " | " & strPANM)
             List1.AddItem ("----------------------------------------")
             DoEvents

             If pGrid_Point > 0 Then
                 For intCol = 12 To .MaxCols
                     strRstval = ""
                     .GetText intCol, 0, varTmp
                     Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                     If Not itemX Is Nothing Then
                         For intIdx = 4 To UBound(mResult)
                             If Trim(LabUReader.TestId(intIdx)) = itemX.Tag Then
                                 Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                 If Not itemX Is Nothing Then
                                     strEqpCd = itemX.text
                                     strRstval = Trim(LabUReader.Result(intIdx))
                                     strRefVal = ""
                                     
                                     If UCase(itemX.Tag) = "PH" Then
                                         strRstval = Format(strRstval, "##.0")
                                     End If
                                     
                                     Select Case Trim(strRstval)
                                         Case "neg", "norm", "-"
                                             strRstval = "음성(-)"
                                         Case "pos"
                                             strRstval = "양성(+)"
                                         Case "+"
                                             strRstval = "양성(+)"
                                         Case "++"
                                             strRstval = "양성(++)"
                                         Case "+++"
                                             strRstval = "양성(+++)"
                                         Case "++++"
                                             strRstval = "양성(++++)"
                                         Case "+++++"
                                             strRstval = "양성(+++++)"
                                         Case "+/-"
                                             strRstval = "약양성(+/-)"
                                         Case Else
                                             strRstval = Trim(strRstval)
                                     End Select

                                     strDate = Format$(Now, "YYYYMMDD"):    strTime = Format$(Now, "HHMMSS")
                                     
                                    '-- 처방번호 찾기
                                    spdResult1.GetText intCol, pGrid_Point, varTmp: strORQN = varTmp
                                    spdResult1.Col = intCol
                                    spdResult1.ForeColor = vbBlack
                                     
                                     .SetText intCol, pGrid_Point, strRstval
                                     .Col = colSQNO: .ForeColor = vbRed: .BackColor = vbCyan ': .text = "받음"
                                     .SetText colBANO, pGrid_Point, "1"
                                     
                                     If Len(strEqpCd) <> 0 And strORQN <> "" And strORQN <> "" Then
                                         '-- H/L 판정
                                         'strRefVal = f_subSet_RefVal(strEqpCd, strRstval, strSEXS, strAGES)
                                         strRefVal = " "
                                         
                                        strSENO = 0
                                        'strNWNO = 0

                                         '-- 로컬저장
                                         sqlDoc = "insert into INTERFACE004(" & _
                                                  " ORDT,  ORQN,  PAID,     OIFL,   SENO," & _
                                                  " NWNO,  ORCD,  TRANSDT,  TRANSTM,EQPCD," & _
                                                  " RSTVAL,REFVAL,SERVERGBN,PANM,   SEX,  AGE,EQPNUM)" & _
                                                  " values( '" & strORDT & "'," & strORQN & ", '" & strPAID & "','" & strOIFL & "',  " & strSENO & ",  " & _
                                                                 strNWNO & ",'" & strORCD & "','" & strDate & "','" & strTime & "','" & strEqpCd & "','" & _
                                                                 strRstval & "','" & strRefVal & "','N','" & strPANM & "','" & strSEXS & "','" & strAGES & "','" & itemX.Tag & "')"
                                             
                                         AdoCn_Jet.Execute sqlDoc
                                         
                                         If chkAuto.Value = "1" And Len(strEqpCd) <> 0 Then
                                             '-- 서버저장
                                            sqlDoc = " Update LRESULT"
                                            sqlDoc = sqlDoc & "   Set RSFL = 'Y',"
                                            sqlDoc = sqlDoc & "       RSLT = '" & strRstval & "',"
                                            sqlDoc = sqlDoc & "       HLFL = '" & strRefVal & "',"
                                            sqlDoc = sqlDoc & "       RSDT = '" & Format(Now, "YYYYMMDD") & "',"
                                            sqlDoc = sqlDoc & "       RSID = '" & CurrUser.CuUserID & "'"
                                            sqlDoc = sqlDoc & " Where PAID = '" & strPAID & "'"
                                            sqlDoc = sqlDoc & "   And NWNO = " & strNWNO
                                            sqlDoc = sqlDoc & "   And ORDT = '" & strORDT & "'"
                                            sqlDoc = sqlDoc & "   And ORQN = " & strORQN
                                            sqlDoc = sqlDoc & "   And OIFL = '" & strOIFL & "'"
                                            sqlDoc = sqlDoc & "   And ORCD = '" & strEqpCd & "'"
                                            sqlDoc = sqlDoc & "   And OKFL <> 'Y'"
                                            
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
                                            spdResult1.Col = colAGES:   spdResult1.BackColor = vbCyan
                                            spdResult1.Col = colNWNO:   spdResult1.BackColor = vbCyan
                                            spdResult1.Col = colBANO:   spdResult1.Value = 0
                                                                                                    
                                            '-- 로컬 업데이트
                                            sqlDoc = "Update INTERFACE004 "
                                            sqlDoc = sqlDoc & "   Set SERVERGBN = 'Y' "
                                            sqlDoc = sqlDoc & " Where PAID = '" & strPAID & "'"
                                            sqlDoc = sqlDoc & "   And NWNO = " & strNWNO
                                            sqlDoc = sqlDoc & "   And ORDT = '" & strORDT & "'"
                                            sqlDoc = sqlDoc & "   And ORQN = " & strORQN
                                            sqlDoc = sqlDoc & "   And OIFL = '" & strOIFL & "'"
                                            sqlDoc = sqlDoc & "   And EQPCD = '" & strEqpCd & "'"
                                                                                                             
                                            AdoCn_Jet.Execute sqlDoc
                                         End If
                                     End If
                                 
                                 Exit For
                                 Set itemX = Nothing
                             End If
                            End If
                         Next intIdx
                     End If
                 Next
             End If
         End With
     End If
     
     ReceiveData = ""
    
    Exit Sub
    
ErrReceive:

    Call ErrMsgProc(CallForm)

End Sub


Private Sub ComReceive(ByRef RecData As String)
                
    Dim strRec  As String, strBuff, strBuffLen As String
    Dim strTmp  As String, intIdx   As Integer
    Dim intPos1 As Integer, intPos2 As Integer
    
    Dim strdata()   As String, intCnt   As Integer
    
    Static OrgMsg As String
    strRec = RecData
    Debug.Print strRec
    
    Print #1, strRec;
    
    'strTmp = strRec
    
    For intIdx = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx, 1)
        'strBuffLen = Mid$(strRec, intIdx, 2)
        Select Case strBuff
            Case "~" '-- STX & ETX
                If f_strBuffer <> "" And Len(f_strBuffer) > 50 Then
                    Call psDataDefine(f_strBuffer, fChannel(), spdResult1)
                    f_strBuffer = ""
                    Exit For
                End If
                f_strBuffer = f_strBuffer + strBuff
            Case Else
                f_strBuffer = f_strBuffer + strBuff
        End Select
     Next
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
   
    Dim Arr()   As Byte
    Dim strTmp  As String
   

    ' UroMeter 720
    strTmp = strTmp & "~NOV/28/2009 11:12:10" & vbCrLf
    strTmp = strTmp & "Name:             Sex:" & vbCrLf
    strTmp = strTmp & "Ward:             Age:" & vbCrLf
    strTmp = strTmp & "SD (10)      (1051)" & vbCrLf
    strTmp = strTmp & "BLD       -        neg" & vbCrLf
    strTmp = strTmp & "BLD     +++   250RBC/ul" & vbCrLf
    strTmp = strTmp & "BIL       -        neg" & vbCrLf
    strTmp = strTmp & "URO    norm    0.1mg/dl" & vbCrLf
    strTmp = strTmp & "KET       -        neg" & vbCrLf
    strTmp = strTmp & "PRO       -        neg" & vbCrLf
    strTmp = strTmp & "NIT       -        neg" & vbCrLf
    strTmp = strTmp & "GLU       -        neg" & vbCrLf
    strTmp = strTmp & "pH                 7.0" & vbCrLf
    strTmp = strTmp & "S.G              1.015" & vbCrLf
    strTmp = strTmp & "LEU       -        neg" & vbCrLf
    strTmp = strTmp & "(SN=60158391)" & vbCrLf
    strTmp = strTmp & "ID(             )" & vbCrLf
    strTmp = strTmp & "OP(00000000) LOT(000000) ~" & vbCrLf
    'strTmp = strTmp & "~FinDb~" & vbCrLf

strTmp = ""
    strTmp = strTmp & "~DEC/06/2011 09:58:46" & vbCrLf
    strTmp = strTmp & "Name:             Sex:" & vbCrLf
    strTmp = strTmp & "Ward:             Age:" & vbCrLf
    strTmp = strTmp & "SD (10)      (1046)" & vbCrLf
    strTmp = strTmp & "BLD     +++   250RBC/ul" & vbCrLf
    strTmp = strTmp & "BIL -neg" & vbCrLf
    strTmp = strTmp & "URO       +    1.0mg/dl" & vbCrLf
    strTmp = strTmp & "KET -neg" & vbCrLf
    strTmp = strTmp & "PRO       +     30mg/dl" & vbCrLf
    strTmp = strTmp & "NIT     pos    0.1mg/dl" & vbCrLf
    strTmp = strTmp & "GLU     +/-    100mg/dl" & vbCrLf
    strTmp = strTmp & "pH 6.5" & vbCrLf
    strTmp = strTmp & "S.G 1.01" & vbCrLf
    strTmp = strTmp & "LEU     +++   500WBC/ul" & vbCrLf
    strTmp = strTmp & "(SN=60158713)" & vbCrLf
    strTmp = strTmp & "ID (6622)" & vbCrLf
    strTmp = strTmp & "OP(00000000) LOT(000000) ~" & vbCrLf
    
    
    strTmp = ""
    strTmp = strTmp & "~DEC/10/2011 21:46:03     " & vbCrLf
    strTmp = strTmp & "Name:             Sex:    " & vbCrLf
    strTmp = strTmp & "Ward:             Age:    " & vbCrLf
    strTmp = strTmp & "SD (10)      (1100)       " & vbCrLf
    strTmp = strTmp & "BLD     +++   250RBC/ul   " & vbCrLf
    strTmp = strTmp & "BIL       -        neg    " & vbCrLf
    strTmp = strTmp & "URO    norm    0.1mg/dl   " & vbCrLf
    strTmp = strTmp & "KET       -        neg    " & vbCrLf
    strTmp = strTmp & "PRO       -        neg    " & vbCrLf
    strTmp = strTmp & "NIT       -        neg    " & vbCrLf
    strTmp = strTmp & "GLU     +/-    100mg/dl   " & vbCrLf
    strTmp = strTmp & "pH                 6.5    " & vbCrLf
    strTmp = strTmp & "S.G              1.010    " & vbCrLf
    strTmp = strTmp & "LEU     +/-    10WBC/ul   " & vbCrLf
    strTmp = strTmp & "(SN=60158713)             " & vbCrLf
    strTmp = strTmp & "ID(        20643)         " & vbCrLf
    strTmp = strTmp & "OP(00000000) LOT(000000) ~" & vbCrLf
    
    strTmp = ""
    strTmp = strTmp & "~DEC/11/2011 01:18:15     " & vbCrLf
    strTmp = strTmp & "Name:             Sex:    " & vbCrLf
    strTmp = strTmp & "Ward:             Age:    " & vbCrLf
    strTmp = strTmp & "SD (10)      (1101)       " & vbCrLf
    strTmp = strTmp & "BLD       +    10RBC/ul   " & vbCrLf
    strTmp = strTmp & "BIL       -        neg    " & vbCrLf
    strTmp = strTmp & "URO    norm    0.1mg/dl   " & vbCrLf
    strTmp = strTmp & "KET       +     10mg/dl   " & vbCrLf
    strTmp = strTmp & "PRO       -        neg    " & vbCrLf
    strTmp = strTmp & "NIT       -        neg    " & vbCrLf
    strTmp = strTmp & "GLU     +/-    100mg/dl   " & vbCrLf
    strTmp = strTmp & "pH                 7.5    " & vbCrLf
    strTmp = strTmp & "S.G              1.025    " & vbCrLf
    strTmp = strTmp & "LEU       -        neg    " & vbCrLf
    strTmp = strTmp & "(SN=60158713)             " & vbCrLf
    strTmp = strTmp & "ID(        20644)         " & vbCrLf
    strTmp = strTmp & "OP(00000000) LOT(000000) ~" & vbCrLf
    
    
    strBuffer = strTmp
    Call comEQP_OnComm
    
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
    mskOrdDate.text = Format$(Now, "YYYYMMDD")
    mskOrdDate1.text = Format$(Now, "YYYYMMDD")
    
    Open App.Path + "\" + REG_INSNAME + ".Log" For Append As #1

    Print #1, Chr(13) + Chr(10);
    
    Open App.Path + "\ErrorLog\" + REG_INSNAME + "_" + Format(Now, "YYYYMMDD") + ".sql" For Append As #2

    Print #2, Chr(13) + Chr(10);
   
    tabWork.Tab = 0
    
    Timer1.Interval = 60000
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    comEQP.Output = ACK
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

    Dim arow    As Integer, aCOL   As Integer
    Dim varChk  As Variant, varBar As Variant, varNum As Variant
    Dim iRow    As Integer, iCnt   As Integer
    
    'Debug.Print Col & NewCol & Row & NewRow
       
    If KeyAscii = vbKeyReturn Then
        With spdResult1
            aCOL = .ActiveCol
            arow = .ActiveRow
            If aCOL = 4 Then
                iCnt = 0
                For iRow = arow To .MaxRows
                    .GetText 1, iRow, varChk
                    .GetText 3, iRow, varBar
                    .GetText aCOL, arow, varNum
                    If Trim(varChk) = "1" And Trim(varBar) <> "" Then
                        .SetText aCOL, iRow, varNum
                        .SetText aCOL + 1, iRow, ((iCnt Mod 40) + 1) + (40 * (varNum - 1))
                        iCnt = iCnt + 1
                        If (iCnt Mod 40) = 0 Then varNum = varNum + 1
                    End If
                Next
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
    
    intCol1 = 12
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

Private Sub Timer2_Timer()
    f_strBuffer = ""
    Timer2.Enabled = False
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

