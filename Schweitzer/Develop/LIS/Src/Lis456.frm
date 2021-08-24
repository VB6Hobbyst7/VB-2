VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm456SuscTrand 
   BackColor       =   &H00DBE6E6&
   Caption         =   "감수성 추이"
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   420
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14655
   Tag             =   "Antibiotic Susceptibility Report"
   WindowState     =   2  '최대화
   Begin TabDlg.SSTab sstResult 
      Height          =   375
      Left            =   75
      TabIndex        =   6
      Top             =   8100
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   661
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tab             =   2
      TabHeight       =   511
      BackColor       =   14411494
      TabCaption(0)   =   "Result  "
      TabPicture(0)   =   "Lis456.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Define Condition  "
      TabPicture(1)   =   "Lis456.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "SuscTrand Statics"
      TabPicture(2)   =   "Lis456.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
   End
   Begin VB.CommandButton cmdPopupList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10725
      MousePointer    =   14  '화살표와 물음표
      Picture         =   "Lis456.frx":0054
      Style           =   1  '그래픽
      TabIndex        =   22
      Top             =   75
      Width           =   315
   End
   Begin VB.TextBox txtSpcCd 
      Height          =   345
      Left            =   11055
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   75
      Width           =   1545
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  '없음
      Caption         =   "Frame5"
      Height          =   1515
      Left            =   75
      TabIndex        =   11
      Top             =   510
      Width           =   14370
      Begin VB.TextBox txtAnti 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "AMP, AZP, BLA, CAZ, CRO, GEN"
         Top             =   555
         Width           =   12585
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F7FFFF&
         Height          =   510
         Left            =   1005
         TabIndex        =   14
         Top             =   870
         Width           =   12585
         Begin VB.OptionButton optSen 
            BackColor       =   &H00F7FFFF&
            Caption         =   "N"
            Height          =   180
            Index           =   4
            Left            =   3375
            TabIndex        =   19
            Top             =   210
            Width           =   600
         End
         Begin VB.OptionButton optSen 
            BackColor       =   &H00F7FFFF&
            Caption         =   "P"
            Height          =   180
            Index           =   3
            Left            =   2625
            TabIndex        =   18
            Top             =   210
            Width           =   600
         End
         Begin VB.OptionButton optSen 
            BackColor       =   &H00F7FFFF&
            Caption         =   "R"
            Height          =   180
            Index           =   2
            Left            =   1845
            TabIndex        =   17
            Top             =   210
            Width           =   600
         End
         Begin VB.OptionButton optSen 
            BackColor       =   &H00F7FFFF&
            Caption         =   "I"
            Height          =   180
            Index           =   1
            Left            =   1095
            TabIndex        =   16
            Top             =   210
            Width           =   600
         End
         Begin VB.OptionButton optSen 
            BackColor       =   &H00F7FFFF&
            Caption         =   "S"
            Height          =   180
            Index           =   0
            Left            =   300
            TabIndex        =   15
            Top             =   210
            Value           =   -1  'True
            Width           =   600
         End
      End
      Begin VB.TextBox txtMic 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "ACT, AEH, CAM, CLA, EAG, PAE"
         Top             =   165
         Width           =   12585
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   375
         Index           =   3
         Left            =   45
         TabIndex        =   55
         Top             =   150
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "미생물"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   405
         Index           =   4
         Left            =   45
         TabIndex        =   56
         Top             =   540
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   714
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "항생제"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   405
         Index           =   5
         Left            =   45
         TabIndex        =   57
         Top             =   960
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   714
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "감수성"
         Appearance      =   0
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F7FFFF&
         BackStyle       =   1  '투명하지 않음
         Height          =   1515
         Index           =   0
         Left            =   15
         Top             =   0
         Width           =   14355
      End
   End
   Begin VB.CommandButton cmdWardList 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8190
      MousePointer    =   14  '화살표와 물음표
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   60
      Width           =   285
   End
   Begin VB.TextBox txtWardId 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7350
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   60
      Width           =   825
   End
   Begin VB.ComboBox cboDept 
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   8445
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FCEFE9&
      Caption         =   "검색시작(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "158"
      Top             =   8535
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker txtDate2 
      Height          =   360
      Left            =   2625
      TabIndex        =   1
      Top             =   60
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyyMMdd"
      Format          =   65273859
      CurrentDate     =   36392
   End
   Begin MSComCtl2.DTPicker txtDate1 
      Height          =   375
      Left            =   1020
      TabIndex        =   0
      Top             =   45
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyyMMdd"
      Format          =   65273859
      CurrentDate     =   36392
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   5295
      Top             =   8190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   5190
      TabIndex        =   20
      Top             =   8220
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
      SpreadDesigner  =   "Lis456.frx":05DE
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Index           =   0
      Left            =   75
      TabIndex        =   48
      Top             =   45
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   661
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "조회기간"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   330
      Index           =   1
      Left            =   9765
      TabIndex        =   49
      Top             =   75
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   582
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "검체코드"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   345
      Index           =   2
      Left            =   4035
      TabIndex        =   51
      Top             =   60
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "조회유형"
      Appearance      =   0
   End
   Begin VB.Frame fraInOut 
      BackColor       =   &H00DBE6E6&
      Height          =   450
      Left            =   4980
      TabIndex        =   50
      Top             =   -30
      Width           =   2340
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "외래"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   0
         Left            =   1530
         TabIndex        =   53
         Top             =   165
         Width           =   750
      End
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "병동"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   1
         Left            =   780
         TabIndex        =   54
         Top             =   165
         Width           =   750
      End
      Begin VB.OptionButton optBussDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B679D&
         Height          =   240
         Index           =   2
         Left            =   30
         TabIndex        =   52
         Top             =   165
         Value           =   -1  'True
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   6030
      Left            =   75
      TabIndex        =   24
      Top             =   2070
      Width           =   14385
      Begin FPSpread.vaSpread tblSus 
         Height          =   5115
         Left            =   210
         TabIndex        =   29
         Top             =   765
         Width           =   13035
         _Version        =   196608
         _ExtentX        =   22992
         _ExtentY        =   9022
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   14737632
         MaxCols         =   6
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis456.frx":0788
         UserResize      =   0
      End
      Begin VB.CommandButton cmdSusPrint 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   11925
         Style           =   1  '그래픽
         TabIndex        =   28
         Tag             =   "132"
         Top             =   195
         Width           =   1320
      End
      Begin VB.CommandButton cmdSusExcel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "To Excel"
         Height          =   510
         Left            =   10605
         Style           =   1  '그래픽
         TabIndex        =   27
         Tag             =   "127"
         Top             =   195
         Width           =   1320
      End
      Begin VB.CommandButton cmdStatics 
         BackColor       =   &H00FCEFE9&
         Caption         =   "실행(&S)"
         Height          =   510
         Left            =   9285
         Style           =   1  '그래픽
         TabIndex        =   26
         Tag             =   "158"
         Top             =   195
         Width           =   1320
      End
      Begin VB.PictureBox pic 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   585
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   12720
         TabIndex        =   25
         Top             =   780
         Width           =   12720
      End
      Begin VB.Label Label11 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "♣ 균별 항생제 통계"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   47
         Top             =   270
         Width           =   2685
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   420
         Index           =   2
         Left            =   210
         Shape           =   4  '둥근 사각형
         Top             =   195
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   6030
      Left            =   75
      TabIndex        =   30
      Top             =   2055
      Width           =   14385
      Begin VB.ComboBox cboMKind 
         Height          =   300
         Left            =   1605
         TabIndex        =   37
         Text            =   "Combo1"
         Top             =   345
         Width           =   1665
      End
      Begin VB.CommandButton cmdClsMic 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Clear"
         Height          =   510
         Left            =   3825
         Style           =   1  '그래픽
         TabIndex        =   36
         Tag             =   "45609"
         Top             =   165
         Width           =   1320
      End
      Begin VB.ListBox lstMic 
         BackColor       =   &H00EEE9E6&
         CausesValidation=   0   'False
         Columns         =   3
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         Left            =   120
         Style           =   1  '확인란
         TabIndex        =   35
         Top             =   720
         Width           =   6360
      End
      Begin VB.ListBox lstAnti 
         BackColor       =   &H00E0E3DB&
         Columns         =   3
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         Left            =   6930
         Style           =   1  '확인란
         TabIndex        =   34
         Top             =   720
         Width           =   6450
      End
      Begin VB.CommandButton cmdDetMic 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Determine"
         Height          =   510
         Left            =   5145
         Style           =   1  '그래픽
         TabIndex        =   33
         Tag             =   "45609"
         Top             =   165
         Width           =   1320
      End
      Begin VB.CommandButton cmdClsAnti 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Clear"
         Height          =   510
         Left            =   10725
         Style           =   1  '그래픽
         TabIndex        =   32
         Tag             =   "45609"
         Top             =   165
         Width           =   1320
      End
      Begin VB.CommandButton cmdDetAnti 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Determine"
         Height          =   510
         Left            =   12045
         Style           =   1  '그래픽
         TabIndex        =   31
         Top             =   165
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   405
         Index           =   6
         Left            =   120
         TabIndex        =   58
         Top             =   225
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   714
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "균종,균 선택"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   405
         Index           =   7
         Left            =   6930
         TabIndex        =   59
         Top             =   225
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   714
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "항생제 선택"
         Appearance      =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   6705
         X2              =   6705
         Y1              =   720
         Y2              =   5685
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   6690
         X2              =   6690
         Y1              =   720
         Y2              =   5670
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   6030
      Left            =   75
      TabIndex        =   38
      Top             =   2070
      Width           =   14385
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   12015
         Style           =   1  '그래픽
         TabIndex        =   44
         Tag             =   "132"
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "To &Excel"
         Height          =   510
         Left            =   10695
         Style           =   1  '그래픽
         TabIndex        =   43
         Tag             =   "127"
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00DBE6E6&
         Caption         =   "S&ort"
         Height          =   510
         Left            =   9375
         Style           =   1  '그래픽
         TabIndex        =   42
         Tag             =   "45610"
         Top             =   135
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CheckBox chkOField 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Total Count"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   41
         Tag             =   "45611"
         Top             =   300
         Width           =   1350
      End
      Begin VB.CheckBox chkOField 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Accepted Count"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   1
         Left            =   1725
         TabIndex        =   40
         Tag             =   "45612"
         Top             =   285
         Width           =   1710
      End
      Begin VB.CheckBox chkOField 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Percentage"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   2
         Left            =   3735
         TabIndex        =   39
         Tag             =   "45613"
         Top             =   270
         Value           =   1  '확인
         Width           =   1350
      End
      Begin FPSpread.vaSpread ssResult 
         Height          =   4815
         Left            =   105
         TabIndex        =   45
         Top             =   675
         Width           =   13230
         _Version        =   196608
         _ExtentX        =   23336
         _ExtentY        =   8493
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   14737632
         MaxCols         =   15
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis456.frx":0EDD
         UserResize      =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "미생물이나 항생제의 나열 순서를 바꾸시려면 해당 아이템의 헤더부분을 더블클릭하세요. 선택된 아이템이 맨 뒤로 옮겨집니다."
         ForeColor       =   &H0062394A&
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   5700
         Width           =   13155
      End
   End
   Begin VB.Label lblSpcNm 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   1  '단일 고정
      Caption         =   "lblSpcNm"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   12615
      TabIndex        =   23
      Top             =   75
      Width           =   1830
   End
   Begin VB.Label lblWardNm 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   1  '단일 고정
      Caption         =   "WardNm"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8505
      TabIndex        =   10
      Top             =   75
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "병동/부서"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Tag             =   "45601"
      Top             =   8490
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Caption         =   "-"
      Height          =   240
      Index           =   0
      Left            =   2475
      TabIndex        =   3
      Top             =   135
      Width           =   270
   End
End
Attribute VB_Name = "frm456SuscTrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cAllMic As String = "(전체)"

Dim aMic(0 To 1000) As String
Dim aAnti(0 To 400) As String

Dim aM() As String
Dim aMnm() As String
Dim aA() As String
Dim aR() As Single
Dim aSen As String

Public Event LastFormUnload()
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private objDic                  As clsDictionary
Private objMn                   As clsDictionary

Private Sub cboMKind_Click()

'    Call cmdClsMic_Click
'
'    Dim sMKindCd As String
'
'    sMKindCd = cboMKind.List(cboMKind.ListIndex)
'    Call LoadMicData(sMKindCd)
    Call cmdClsMic_Click

    Dim sMKindCd As String

    sMKindCd = cboMKind.List(cboMKind.ListIndex)
    Call LoadMicData(sMKindCd)
    Call ChkMicAnti
End Sub


Private Sub ChkMicAnti()
    Dim RS      As Recordset
    Dim objAnti As New clsDictionary
    Dim strTmp() As String
    
    Dim sSusCd  As String   '균종코드
    Dim SSQL    As String
    Dim ii      As Integer
    
    objAnti.Clear
    objAnti.FieldInialize "anticd", "antinm"
    
    
    sSusCd = cboMKind.List(cboMKind.ListIndex)
    
    SSQL = " SELECT cdval2 as MsGs , text1 as Anti " & _
           " FROM " & T_LAB031 & _
           " WHERE " & DBW("cdindex = ", "C108") & _
           " AND " & DBW("cdval1 =", sSusCd)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            strTmp() = Split(RS.Fields("anti").Value & "", ";")
            For ii = LBound(strTmp()) + 1 To UBound(strTmp())
                If objAnti.Exists(strTmp(ii)) = False Then
                    objAnti.AddNew strTmp(ii), strTmp(ii)
                End If
            Next
            RS.MoveNext
        Loop
    End If
    
    If objAnti.RecordCount > 0 Then
        For ii = 1 To lstAnti.ListCount
            lstAnti.Selected(ii - 1) = False
            If objAnti.Exists(aAnti(ii)) = True Then
                lstAnti.Selected(ii - 1) = True
            End If
        Next
    Else
        For ii = 1 To lstAnti.ListCount
            lstAnti.Selected(ii - 1) = False
        Next
        
    End If
    txtAnti.Text = ""
    Call cmdDetAnti_Click
    Set RS = Nothing
    Set objAnti = Nothing
    
End Sub

Private Sub chkOField_Click(Index As Integer)

    Dim i As Integer, sWhat As Integer

    For i = 1 To ssResult.MaxCols
        sWhat = IIf((i Mod 3) = 0, 3, (i Mod 3))
        If chkOField(sWhat - 1).Value = 1 Then
            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = False
        Else
            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = True
        End If
    Next i
    ssResult.LeftCol = 1
    
End Sub

Private Sub cmdClsAnti_Click()
    
    Dim i As Integer
    For i = 0 To lstAnti.ListCount - 1
        'If lstAnti.Selected(i) Then lstAnti.Selected(i) = False
        lstAnti.Selected(i) = False
    Next i
    txtAnti.Text = ""
    Erase aA

End Sub

Private Sub cmdClsMic_Click()

    Dim i As Integer
    For i = 0 To lstMic.ListCount - 1
        'If lstMic.Selected(i) Then lstMic.Selected(i) = False
        lstMic.Selected(i) = False
    Next i
    txtMic.Tag = ""
    txtMic.Text = ""
    Erase aM: Erase aMnm

End Sub

Private Sub cmdDetAnti_Click()
    Dim sAntiBuf As String

    sAntiBuf = ""
    
    Dim i As Integer, iA As Integer, sTmp As String
    iA = 0: sTmp = ""
    For i = 0 To lstAnti.ListCount - 1
    
        If lstAnti.Selected(i) Then
            iA = iA + 1
            sTmp = sTmp & Trim(aAnti(i + 1)) & ";"
            'ReDim Preserve aA(1 To iA)
            'aA(iA) = Trim(aAnti(i + 1))
            
            If sAntiBuf <> "" Then sAntiBuf = sAntiBuf & ","
            sAntiBuf = sAntiBuf & aAnti(i + 1)
        End If
        
    Next i
    
    If iA > 0 Then
        ReDim aA(1 To iA)
        For i = 1 To iA
            aA(i) = medGetP(sTmp, i, ";")
        Next i
    End If
    txtAnti.Text = sAntiBuf
    
End Sub

Private Sub cmdDetMic_Click()
    Dim sMicBuf As String
    Dim sMnName As String
    Dim sMnTmp  As String
    
    sMicBuf = ""
    
    Dim i As Integer, iM As Integer, sTmp As String, sTnm As String
    iM = 0: sTmp = "": sTnm = ""
    For i = 0 To lstMic.ListCount - 1
    
        If lstMic.Selected(i) Then
            iM = iM + 1
            sTmp = sTmp & Trim(aMic(i + 1)) & ";"
            sTnm = sTnm & Trim(lstMic.List(i)) & ";"
                        
            If sMicBuf <> "" Then sMicBuf = sMicBuf & ","
            sMicBuf = sMicBuf & "'" & Trim(aMic(i + 1)) & "'"
            sMnTmp = BacteriaName(aMic(i + 1))
            If sMnName <> "" Then sMnName = sMnName & ","
            sMnName = sMnName & "'" & sMnTmp & "'"
        End If
        
    Next i
    
    If iM > 0 Then
        ReDim aM(1 To iM): ReDim aMnm(1 To iM)
        For i = 1 To iM
            aM(i) = medGetP(sTmp, i, ";")
            aMnm(i) = medGetP(sTnm, i, ";")
        Next i
    End If
    txtMic.Tag = sMicBuf
    txtMic.Text = sMnName

End Sub

Private Function BacteriaName(ByVal MnCd As String) As String
    Dim SSQL As String
    Dim RS    As Recordset
    
    
    SSQL = " SELECT * FROM " & T_LAB032 & _
           " WHERE " & DBW("cdindex=", LC3_Microbe) & _
           " AND   " & DBW("field2=", cboMKind.List(cboMKind.ListIndex)) & _
           " AND   " & DBW("cdval1=", MnCd)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        BacteriaName = RS.Fields("field1").Value & ""
    End If
    Set RS = Nothing
End Function



Private Sub cmdExcel_Click()
    
    Dim tmpStr As String
    
    With ssResult
        
        .ReDraw = False
        
        .MaxRows = .MaxRows + 1
        .MaxCols = .MaxCols + 1
        
        .Row = 1: .Col = 1
        .Action = ActionInsertRow
        .Action = ActionInsertCol
        
        .Row = 1: .Row2 = .MaxRows
        .Col = 0: .COL2 = 0
        .BlockMode = True
        tmpStr = .Clip
        .BlockMode = False
        .Col = 1: .COL2 = 1
        .BlockMode = True
        .Clip = tmpStr
        .BlockMode = False
        
        .Row = 0: .Row2 = 0
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        tmpStr = .Clip
        .BlockMode = False
        .Row = 1: .Row2 = 1
        .BlockMode = True
        .Clip = tmpStr
        .BlockMode = False
        
        DlgSave.InitDir = "C:\"
        DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
        DlgSave.ShowSave
         
        .SaveTabFile (DlgSave.FileName)
        
        .Row = 1: .Action = ActionDeleteRow
        .Col = 1: .Action = ActionDeleteCol
        
        .MaxRows = .MaxRows - 1
        .MaxCols = .MaxCols - 1
        
        .ReDraw = True
        
    End With
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
'    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdPopupList_Click()
  lblSpcNm.Caption = ""
  txtSpcCd.Text = ""
  Set objCodeList = New clsPopUpList
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Specimen Code List.."
        .Tag = "Specimen"
        .ColumnHeaderText = "검체코드;검체명"
        Call .LoadPopUp(SqlLAB032CodeList(LC3_Specimen, "cdval1, field3")) ', 2100, 8700)
        
        If .SelectedString <> "" Then
            txtSpcCd.Text = medGetP(.SelectedString, 1, ";")
            lblSpcNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
        
    End With
    Set objCodeList = Nothing
End Sub
Private Function SqlLAB032CodeList(ByVal Cdindex As String, ByVal Fields As String, _
                                                   Optional ByVal CdVal1 As Variant, _
                                                   Optional ByVal Orderby As Variant) As String
   SqlLAB032CodeList = "SELECT " & Fields & " " & _
                       "FROM " & T_LAB032 & " " & _
                       "WHERE cdindex = '" & Cdindex & "' "
   If IsMissing(CdVal1) = False Then
      SqlLAB032CodeList = SqlLAB032CodeList & _
                        "AND     cdval1 = '" & CdVal1 & "' "
   End If
   If IsMissing(Orderby) = False Then
      SqlLAB032CodeList = SqlLAB032CodeList & Orderby
   End If

End Function
Private Sub cmdPrint_Click()
    
    With ssResult
    
        .PrintOrientation = PrintOrientationLandscape
        
        .PrintJobName = "미생물 항균제 감수성 추이 출력"
        .PrintAbortMsg = "미생물 항균제 감수성 추이 분석 테이블을 출력중입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1
       
        .PrintHeader = "/n/n/l/fb1 " & "♧ 항균제 감수성 추이 분석 (" & Format(txtDate1.Value, CS_DateDbFormat) & " 부터 " & Format(txtDate2.Value, CS_DateDbFormat) & " 까지 ) /c/fb1/n/n"
        .PrintFooter = "/c/p/fb1"
        
        .PrintGrid = False
        .PrintMarginBottom = 100
        .PrintMarginLeft = 200
        .PrintMarginRight = 100
        .PrintShadows = False
        .PrintMarginTop = 300
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintPageEnd = 2
        .PrintRowHeaders = True
        .PrintColHeaders = True
        .PrintBorder = True
        '.PrintGrid = True
        .PrintGrid = False
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint
        .GridSolid = True
    
    End With

End Sub


Private Sub cmdStart_Click()

    Dim sMsg As String
    Dim objPrgBar As New jProgressBar.clsProgress
    
    With objPrgBar
        .Container = Me
        .Width = sstResult.Width
        .Left = sstResult.Left
        .Top = sstResult.Top
        .Height = 280
        .Message = "자료를 검색중입니다. 데이타량과 기간에 따라서 몇 분이 소요될 수도 있습니다."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = sstResult.Width
'        .XPos = sstResult.Left
'        .YPos = sstResult.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "자료를 검색중입니다. 데이타량과 기간에 따라서 몇 분이 소요될 수도 있습니다."
'        .Value = 1
    End With

    If Trim(txtDate1.Value) = "" Or Trim(txtDate2.Value) = "" Then sMsg = "날짜를 선택하지 않았습니다": GoTo ErrMsg
    If txtDate1.Value > txtDate2.Value Then sMsg = "기간 설정이 잘못되었습니다": GoTo ErrMsg

    If Trim(txtMic.Tag) = "" Then sMsg = "적용 미생물을 선택하지 않았습니다": GoTo ErrMsg
    If Trim(txtAnti) = "" Then sMsg = "적용 항생제를 선택하지 않았습니다": GoTo ErrMsg
    
    objPrgBar.Max = DateDiff("d", txtDate1.Value, txtDate2.Value) + 1
    
    ReDim aR(1 To UBound(aM), 1 To UBound(aA), 1 To 3)
    
    Dim dtTmp As Date, ibar As Integer
    dtTmp = txtDate1.Value: ibar = 0
    Do While dtTmp <= txtDate2.Value
    
        ibar = ibar + 1
        objPrgBar.Message = Format(dtTmp, "dddddd") & " 자료를 검색중입니다"
        objPrgBar.Value = ibar
        DoEvents
        Call SearchDataADay(Format(dtTmp, "yyyymmdd"))
    
        dtTmp = DateAdd("d", 1, dtTmp)
    
    Loop
    Call DisplayResult
    DoEvents
    
    Exit Sub

ErrMsg:
    MsgBox sMsg
    Exit Sub

End Sub

Private Sub SearchDataADay(ByVal pDay As String)
    Dim dsRst       As New Recordset
    Dim objsSQL     As clsLISSqlStatistic
    Dim SSQL        As String
    Dim iRCol       As Integer
    Dim sDept   As String
    Dim sOptDept    As String
    
    
    
    
    Set objsSQL = New clsLISSqlStatistic
    
    sDept = ""
    If optBussDiv(0).Value Then
    '외래
        sOptDept = "2"
        sDept = txtWardId.Text
    ElseIf optBussDiv(1).Value Then
    '병동
        sOptDept = "1"
        sDept = txtWardId.Text
    ElseIf optBussDiv(2).Value Then
    '전체
        sOptDept = "0"
        sDept = txtWardId.Text
    End If
    
'    sSQL = objsSQL.GetSensiResult(pDay, txtMic.Tag)

    If txtSpcCd.Text <> "" Then
        SSQL = " SELECT b.* " & _
               "  FROM " & T_LAB201 & " c," & T_LAB404 & " a," & T_LAB405 & " b" & _
               " WHERE " & _
                           DBW("a.vfydt=", pDay) & _
               "   AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
               "   AND " & DBW("a.senfg=", "Y") & _
               "   AND b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq AND b.testcd=a.testcd AND b.mfyseq=a.mfyseq " & _
               "   AND b.workarea=c.workarea AND b.accdt=c.accdt AND b.accseq=c.accseq" & _
               "   AND " & DBW("c.spccd=", Trim(txtSpcCd.Text)) & _
               "   AND b.mnmcd in (" & txtMic.Tag & ")"
    Else
        SSQL = " SELECT b.* " & _
               "  FROM " & T_LAB404 & " a," & T_LAB405 & " b" & _
               " WHERE " & _
                         DBW("a.vfydt=", pDay) & _
               " AND " & DBW("a.stscd>=", enStsCd.StsCd_LIS_FinRst) & _
               " AND " & DBW("a.senfg=", "Y") & _
               " AND b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
               " AND b.testcd=a.testcd AND b.mfyseq=a.mfyseq " & _
               " AND b.mnmcd in (" & txtMic.Tag & ")"
    End If
    
    Select Case sOptDept
        Case "1"    '병동
            SSQL = SSQL & objsSQL.GetSensiData(True, pDay, txtMic.Tag, sDept)
        Case "2"    '진료과
            SSQL = SSQL & objsSQL.GetSensiData(False, pDay, txtMic.Tag, sDept)
    End Select
    
'    iRCol = dsRst.OpenCursor(dbconn, SSQL)
    dsRst.Open SSQL, DBConn
    
    
    Do Until dsRst.EOF
            
        Dim sM As String, sC As Integer
        sM = Trim("" & dsRst.Fields("mnmcd").Value)
        sC = Val(Trim("" & dsRst.Fields("scnt").Value))
        
        Dim i As Integer
        Dim iSX As Integer, iSY As Integer, iSZ As Integer
        iSX = -1: iSY = -1: iSZ = -1
        
        For i = 1 To UBound(aM)
            If aM(i) = sM Then
                iSX = i
                Exit For
            End If
        Next i
        
        If iSX > 0 Then
            For i = 1 To sC
                Dim sTmp As String, sA As String, sS As String
                sTmp = Trim("" & dsRst.Fields(i + 9).Value)
                sA = Trim(medGetP(sTmp, 1, ";"))
                sS = Trim(medGetP(sTmp, 2, ";"))
        
                Dim j As Integer
                For j = 1 To UBound(aA)
                    If aA(j) = sA Then
                        iSY = j
                        If sS = aSen Then
                            Call ApplyRst(iSX, iSY, "O")
                        Else
                            Call ApplyRst(iSX, iSY, "X")
                        End If
                        Exit For
                    End If
                Next j
        
            Next i
            
        End If
        dsRst.MoveNext
    Loop
    
    Set dsRst = Nothing
    Set objsSQL = Nothing

End Sub

Private Sub ApplyRst(ByVal pX As Integer, ByVal pY As Integer, ByVal pC As String)

    aR(pX, pY, 1) = aR(pX, pY, 1) + 1
    If pC = "O" Then aR(pX, pY, 2) = aR(pX, pY, 2) + 1
    
    'Debug.Print pX & ":" & pY & ":" & aR(pX, pY, 1) & " - " & aR(pX, pY, 2)

End Sub

Private Sub DisplayResult()

    Dim i As Integer, j As Integer
    For i = 1 To UBound(aM)
        For j = 1 To UBound(aA)
            If aR(i, j, 1) = 0 Then
                aR(i, j, 3) = 0#
            Else
                aR(i, j, 3) = Format((aR(i, j, 2) / aR(i, j, 1)) * 100, "##0")
            End If
        Next j
    Next i

    ssResult.MaxRows = 0

    ssResult.MaxCols = UBound(aA) * 3
    ssResult.MaxRows = UBound(aM)

    Call SetTableAttr
    Dim blnFirst As Boolean
    
    For i = 1 To UBound(aM)
        blnFirst = False
        For j = 1 To UBound(aA)
            ssResult.Row = i
            ssResult.Col = ((j - 1) * 3) + 1:
            ssResult.Value = Format(aR(i, j, 1), "#####0")
            If ssResult.Value = "0" Then
                ssResult.ForeColor = DCM_Gray
            Else
                ssResult.Col = 0:
                If blnFirst = False Then
                    ssResult.Value = ssResult.Value & Space(23 - Len(ssResult.Value)) & " ( " & Format(aR(i, j, 1), "#####0") & " )"
                    blnFirst = True
                End If
            End If
            ssResult.Col = ((j - 1) * 3) + 2: ssResult.Value = Format(aR(i, j, 2), "#####0")
            If ssResult.Value = "0" Then ssResult.ForeColor = DCM_Gray
            ssResult.Col = ((j - 1) * 3) + 3: ssResult.Value = Format(aR(i, j, 3), "#####0")
            If ssResult.Value = "0" Then ssResult.ForeColor = DCM_Gray
'            Call ssResult.SetText(((j - 1) * 3) + 1, I, Format(aR(I, j, 1), "#####0"))
'            Call ssResult.SetText(((j - 1) * 3) + 2, I, Format(aR(I, j, 2), "#####0"))
'            Call ssResult.SetText(((j - 1) * 3) + 3, I, Format(aR(I, j, 3), "#####0"))
        Next j
    Next i

    sstResult.Tab = 0
    ssResult.LeftCol = 1

End Sub

Private Sub SetTableAttr()

    Dim i As Integer, j As Integer, sWhat As Integer
    For i = 1 To ssResult.MaxCols
        If (i Mod 3) = 1 Then
            Call ssResult.SetText(i, 0, aA(((i - 1) \ 3) + 1) & " (T)")
        ElseIf (i Mod 3) = 2 Then
            Call ssResult.SetText(i, 0, aA(((i - 1) \ 3) + 1) & " (" & aSen & ")")
        Else
            Call ssResult.SetText(i, 0, aA(((i - 1) \ 3) + 1) & " (%)")
        End If
        
        sWhat = i Mod 3
        If sWhat = 0 Then
            ssResult.Col = i: ssResult.Row = -1
            ssResult.ForeColor = DCM_LightBlue  'RGB(0, 0, 255)
        End If
        If sWhat = 0 Then sWhat = 3
        If chkOField(sWhat - 1).Value = 1 Then
            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = False
        Else
            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = True
        End If
        
        '
        'SELECT Case sWhat
        '    Case 0:
        '        Call ssResult.SetText(i, 0, aA(((i - 1) \ 3) + 1) & "(T)")
        '        ssResult.Col = i: ssResult.Row = -1: ssResult.ForeColor = RGB(0, 0, 255)
        '        If chkOField(0).Value = 1 Then
        '            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = False
        '        Else
        '            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = true
        '        End If
        '    Case 1:
        '        Call ssResult.SetText(i, 0, aA(((i - 1) \ 3) + 1) & "(A)")
        '        If chkOField(1).Value = 1 Then
        '            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = False
        '        Else
        '            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = true
        '        End If
        '    Case 2:
        '        Call ssResult.SetText(i, 0, aA(((i - 1) \ 3) + 1) & "(%)")
        '        If chkOField(2).Value = 1 Then
        '            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = False
        '        Else
        '            ssResult.Col = i: ssResult.Row = -1: ssResult.ColHidden = true
        '        End If
        'End SELECT
        
        If (i Mod 6) >= 1 And (i Mod 6) <= 3 Then
            ssResult.Col = i: ssResult.Row = -1
            ssResult.BackColor = RGB(238, 254, 237)
        End If
    Next i
    
    For j = 1 To ssResult.MaxRows
        Call ssResult.SetText(0, j, aMnm(j))
        ssResult.ColWidth(0) = 24
        
        ssResult.Row = j
        ssResult.Col = 0
        
        ssResult.TypeHAlign = TypeHAlignLeft
        
    Next j

End Sub

Private Sub cmdSusExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblSus.DataRowCnt = 0 And tblSus.DataRowCnt = 0 Then Exit Sub
    
    With tblSus
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblSus.MaxRows + 1
        .MaxCols = tblSus.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "항생제통계"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdWardList_Click()
    Dim objMyList As clsPopUpList
    Dim strCaption As String
    Dim strHead As String
'    Dim objData As clsBasisData
    
    If optBussDiv(2).Value Then Exit Sub
    
    Set objMyList = New clsPopUpList
'    Set objData = New clsBasisData
    
    If optBussDiv(0).Value Then
        strCaption = "진료과 조회"
        strHead = "부서코드;부서명"
    Else
        strCaption = "병동 조회"
        strHead = "병동코드;병동명"
    End If
    
    With objMyList
        .Connection = DBConn
        .FormCaption = strCaption
        .ColumnHeaderText = strHead
        .Tag = "WardID"
        Me.ScaleMode = 1
        If optBussDiv(0).Value Then
            Call .LoadPopUp(GetSQLDeptList) ', 3950, 6300) ', ObjLISComCode.DeptCd)
        Else
            Call .LoadPopUp(GetSQLWardList) ', 3950, 6300) ', ObjLISComCode.WardId)
        End If
        txtWardId.Text = medGetP(.SelectedString, 1, ";")
        lblWardNm.Caption = medGetP(.SelectedString, 2, ";")

    End With
    
'    Set objData = Nothing
    Set objMyList = Nothing
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()

'    ClearForm
'
'    SetInitDate
'    LoadDeptList
'
'    LoadMKindData
'    cboMKind.ListIndex = 0
'
'    LoadMicData cAllMic
'    LoadAntiData
'
'    sstResult.Tab = 1
'    cboDept.ListIndex = 0
    
    Call ClearForm

    Call SetInitDate
    Call LoadDeptList
    
    '항생제 리스트 가지고 오기
    Call LoadAntiData
    
    '균종 로드
    Call LoadMKindData
    
    '첫번째 균종코드에 해당하는 균 가지고 오기
    cboMKind.ListIndex = 1
        
'    LoadMicData cAllMic

    sstResult.Tab = 1
End Sub

Private Sub ClearForm()
    
    txtSpcCd.Text = ""
    txtMic.Text = ""
    txtAnti.Text = ""
    txtWardId.Text = ""
    
    lblWardNm.Caption = ""
    lblSpcNm.Caption = ""
    optSen(0).Value = True: aSen = "S"
    
    chkOField(0).Value = 0: chkOField(1).Value = 0: chkOField(2).Value = 1
    Call ClearResult
    
    Frame1.ZOrder 0
    
End Sub

Private Sub ClearResult()
    
    ssResult.Col = 0: ssResult.COL2 = ssResult.MaxCols
    ssResult.Row = 0: ssResult.Row2 = ssResult.MaxRows
    ssResult.BlockMode = True
    ssResult.Action = ActionClear
    ssResult.BlockMode = False
    
End Sub

Private Sub SetInitDate()

    txtDate1.Value = Format(Now, "yyyy-mm-") & "01"
    txtDate2.Value = Format(Now, "yyyy-mm-dd")

End Sub

Private Sub LoadMKindData()
    Dim objsSQL As clsLISSqlStatistic
    
    Dim sqlMKind As String, dsMKind As New Recordset, iMKindCol As Integer
    
    Set objsSQL = New clsLISSqlStatistic
        
    sqlMKind = objsSQL.GetSpecies
    
'    sqlMKind = " SELECT cdval1 mkind FROM " & T_LAB032 & _
'               " WHERE " & DBW("cdindex=", LC3_Species) & " " & _
'               " ORDER BY cdval1 asc"
'    iMKindCol = dsMKind.OpenCursor(dbconn, sqlMKind)
    
    dsMKind.Open sqlMKind, DBConn

    cboMKind.Clear
    cboMKind.AddItem cAllMic

    Do Until dsMKind.EOF
        cboMKind.AddItem "" & dsMKind.Fields("mkind").Value
        dsMKind.MoveNext
    Loop
    
    Set dsMKind = Nothing
    Set objsSQL = Nothing

End Sub

Private Sub LoadMicData(ByVal pMKind As String)
    Dim objsSQL As clsLISSqlStatistic
    Dim sqlMic As String, dsMic As New Recordset, iMicCol As Integer

    'lblMName.Caption = ""
    
    Set objsSQL = New clsLISSqlStatistic
    
    If pMKind = cAllMic Then
        sqlMic = objsSQL.GetMicrobe(True)
    Else
        sqlMic = objsSQL.GetMicrobe(False, pMKind)
    End If
    
'    If pMKind = cAllMic Then
'        sqlMic = " SELECT cdval1 mic, text1 micnm FROM " & T_LAB032 & _
'                 " WHERE " & DBW("cdindex=", LC3_Microbe) & " " & _
'                 " ORDER BY cdval1 asc"
'    Else
'        sqlMic = " SELECT cdval1 mic, text1 micnm FROM " & T_LAB032 & _
'                 " WHERE " & DBW("cdindex=", LC3_Microbe) & _
'                 " AND " & DBW("field2=", pMKind) & " " & _
'                 " ORDER BY cdval1 asc"
'    End If
    
    dsMic.Open sqlMic, DBConn
    
    Dim iCnt As Integer
    iCnt = 0
    lstMic.Clear

    Do Until dsMic.EOF
        iCnt = iCnt + 1
        'lstMic.AddItem "" & dsMic.GetValue("micnm") & vbtab & vbtab & vbtab & dsMic.GetValue("mic")
        lstMic.AddItem Mid(Trim("" & dsMic.Fields("micnm").Value), 1, 20)
        aMic(iCnt) = Trim("" & dsMic.Fields("mic").Value)
        dsMic.MoveNext
    Loop
    aMic(0) = iCnt
    
    Set dsMic = Nothing
    Set objsSQL = Nothing
End Sub

Private Sub LoadAntiData()
    Dim objsSQL As clsLISSqlStatistic
    Dim sqlAnti As String, dsAnti As New Recordset, iAntiCol As Integer

    'lblAName.Caption = ""
       
    Set objsSQL = New clsLISSqlStatistic
    sqlAnti = objsSQL.GetAntiBiotic(CS_AllCaption)
    
'    sqlAnti = " SELECT cdval1 anti, text1 antinm FROM " & T_LAB032 & _
'              " WHERE " & DBW("cdindex=", LC3_AntiBiotic) & " " & _
'              " ORDER BY cdval1 asc"
    dsAnti.Open sqlAnti, DBConn

    Dim iCnt As Integer
    iCnt = 0
    lstAnti.Clear

    Do Until dsAnti.EOF
        iCnt = iCnt + 1
        'lstAnti.AddItem "" & dsAnti.GetValue("antinm") & vbtab & vbtab & vbtab & dsAnti.GetValue("anti")
        lstAnti.AddItem Mid(Trim("" & dsAnti.Fields("antinm").Value), 1, 20)
        aAnti(iCnt) = Trim("" & dsAnti.Fields("anti").Value)
        dsAnti.MoveNext
    Loop
    aAnti(0) = iCnt
    
    Set dsAnti = Nothing
    Set objsSQL = Nothing
End Sub

Private Sub LisLabel1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objDic = Nothing
    Set objMn = Nothing
End Sub

Private Sub lstAnti_Click()
'Dim sAntiCd As String, sAntiNm As String
'Dim sqlAnti As String, dsAnti As New Recordset, iAntiCol As Integer
'
'    sAntiCd = medGetP(lstAnti.List(lstAnti.ListIndex), 4, vbtab)
'
'    sqlAnti = "SELECT text1 Antinm FROM " & T_LAB032 & _
'             " WHERE cdindex='" & LC3_AntiBiotic & "' AND cdval1='" & sAntiCd & "'"
'    iAntiCol = dsAnti.OpenCursor(DbConn, sqlAnti)
'
'    If dsAnti.FetchCursor(iAntiCol) Then
'        lblAName.Caption = "" & dsAnti.GetValue("Antinm")
'    Else
'        lblAName.Caption = ""
'    End If
'
'    dsAnti.CloseCursor: Set dsAnti = Nothing
'
End Sub

Private Sub lstMic_Click()
'Dim sMicCd As String, sMicNm As String
'Dim sqlMic As String, dsMic As New Recordset, iMicCol As Integer
'
'    sMicCd = medGetP(lstMic.List(lstMic.ListIndex), 4, vbtab)
'
'    sqlMic = "SELECT text1 micnm FROM " & T_LAB032 & _
'             " WHERE cdindex='" & LC3_Microbe & "' AND cdval1='" & sMicCd & "'"
'    iMicCol = dsMic.OpenCursor(DbConn, sqlMic)
'
'    If dsMic.FetchCursor(iMicCol) Then
'        lblMName.Caption = "" & dsMic.GetValue("micnm")
'    Else
'        lblMName.Caption = ""
'    End If
'
'    dsMic.CloseCursor: Set dsMic = Nothing
'
End Sub

Private Sub cboMKind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub optBussDiv_Click(Index As Integer)
    txtWardId.Text = "": lblWardNm.Caption = ""
End Sub

Private Sub optSen_Click(Index As Integer)
    
    Select Case Index
    
        Case 0: aSen = "S"
        Case 1: aSen = "I"
        Case 2: aSen = "R"
        Case 3: aSen = "P"
        Case 4: aSen = "N"
        Case Else: aSen = ""
    
    End Select
    
End Sub

Private Sub ssResult_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Col <= 0 And Row <= 0 Then Exit Sub
    
    If Col = 0 And Row > 0 Then       ' 미생물 정렬
        
        ssResult.MaxRows = ssResult.MaxRows + 1
        
        ssResult.Col = 0: ssResult.COL2 = ssResult.MaxCols
        ssResult.Row = Row: ssResult.Row2 = Row
        ssResult.DestCol = 0: ssResult.DestRow = ssResult.MaxRows
        

        'ssResult.Action = ActionMoveRange
        '컴파일 뜨다가 변수정의 되지 않아서 바꿈..정규
        ssResult.Action = 20
        
        ssResult.Row = Row
        ssResult.Action = ActionDeleteRow
        
        ssResult.MaxRows = ssResult.MaxRows - 1
        
    End If
    
    If Row = 0 And Col > 0 Then       ' 항생제 정렬
        
        ssResult.MaxCols = ssResult.MaxCols + 3
        
        Dim iWhere As Integer, iC1 As Integer
        iWhere = Col Mod 3
        If iWhere = 1 Then
            iC1 = Col
        ElseIf iWhere = 2 Then
            iC1 = Col - 1
        Else
            iC1 = Col - 2
        End If
        
        ssResult.Col = iC1: ssResult.COL2 = iC1 + 2
        ssResult.Row = 0: ssResult.Row2 = ssResult.MaxRows
        ssResult.DestCol = ssResult.MaxCols - 2: ssResult.DestRow = 0
        
        'ssResult.Action = ActionMoveRange
        ssResult.Action = 20
        
        ssResult.Col = iC1 + 3: ssResult.COL2 = ssResult.MaxCols
        ssResult.Row = 0: ssResult.Row2 = ssResult.MaxRows
        ssResult.DestCol = iC1: ssResult.DestRow = 0
        'ssResult.Action = ActionMoveRange
        ssResult.Action = 20
        
'        ssResult.Col = iC1 + 2: ssResult.Action = ActionDeleteCol
'        ssResult.Col = iC1 + 1: ssResult.Action = ActionDeleteCol
'        ssResult.Col = iC1: ssResult.Action = ActionDeleteCol
        
        ssResult.MaxCols = ssResult.MaxCols - 3
        
    End If

End Sub


Private Sub LoadDeptList()

    Dim MySql As New clsLISSqlStatement
    Dim SqlStmt As String
    Dim strTmp As String
    Dim RS As Recordset
'    Dim objData As clsBasisData
    
'    Set objData = New clsBasisData
    
    cboDept.Clear
    cboDept.AddItem "<ALL>"
    SqlStmt = GetSQLWardList  'getwardlistsql
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    While (Not RS.EOF)
        strTmp = "" & RS.Fields("WardId").Value
        strTmp = strTmp & Space(7 - Len(Mid(strTmp, 1, 5))) & "" & RS.Fields("WardNm").Value
        cboDept.AddItem strTmp & vbTab & "1"
        RS.MoveNext
    Wend
    SqlStmt = GetSQLDeptList 'getdeptlistsql
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    While (Not RS.EOF)
        strTmp = "" & RS.Fields("DeptCd").Value
        strTmp = strTmp & Space(7 - Len(Mid(strTmp, 1, 5))) & "" & RS.Fields("DeptNm").Value
        cboDept.AddItem strTmp & vbTab & "2"
        RS.MoveNext
    Wend
    Set RS = Nothing
    Set MySql = Nothing
'    Set objData = Nothing
    
End Sub

'====================================================================================
'2003년 1월 21일 추가 부분(S,R,I)공통부분:
'Coding by KJG
'====================================================================================

Private Sub SusTrandQuery()
    Dim i           As Integer
    Dim objPro      As jProgressBar.clsProgress
    Dim dtTmp       As Date
    Dim ibar        As Integer
    
    Set objDic = New clsDictionary
    
    objDic.Clear
    objDic.FieldInialize "mncd,suscd", "scnt,rcnt,icnt,spccnt"
    objDic.Sort = False
    
    Set objMn = New clsDictionary
    objMn.Clear
    objMn.FieldInialize "mncd", "mncnt"
    objMn.Sort = False
    
    Set objPro = Nothing
    Set objPro = New jProgressBar.clsProgress
    
    Call medClearTable(tblSus)
    Me.MousePointer = 11
    
    With objPro
        .Container = Me
        .Left = Frame4.Left + pic.Left
        .Top = Frame4.Top + pic.Top
        .Width = pic.Width
        .Height = pic.Height
        .Max = DateDiff("D", txtDate1.Value, txtDate2.Value)
'        .Choice = True
'        .SetMyForm Me
'        .XPos = Frame4.Left + pic.Left
'        .YPos = Frame4.Top + pic.Top
'        .XWidth = pic.Width
'        .YHeight = pic.Height
'        .Appearance = aPlate
'        .Max = DateDiff("D", txtDate1.Value, txtDate2.Value)
    End With
    pic.ZOrder 0
    
    dtTmp = txtDate1.Value: ibar = 0
    
'    Do While dtTmp <= txtDate2.Value
        DoEvents
        ibar = ibar + 1
'        objPro.Message = Format(dtTmp, "dddddd") & " 자료를 검색중입니다"
        objPro.Message = " 자료를 검색중입니다"
        objPro.Value = ibar

        '데이터 조회및 담아주기(objdic)
        Call SusTrandSearchDataADay(Format(dtTmp, "yyyymmdd"))
        '데이터 보여주기
        
        dtTmp = DateAdd("d", 1, dtTmp)
'    Loop
    Set objPro = Nothing
    
    tblSus.ZOrder 0
    
    If objDic.RecordCount < 1 Then
        MsgBox "해당조건의 자료가 없습니다.", vbInformation + vbOKOnly
    Else
        objDic.Sort = True
        Call SusctrandDisplay
    End If
    Me.MousePointer = 0
End Sub
Private Sub SusTrandSearchDataADay(ByVal pDay As String)
    Dim dsRst       As New Recordset
    Dim objsSQL     As clsLISSqlStatistic
    Dim SSQL        As String
    
    Dim sDept       As String
    Dim sOptDept    As String
    Dim sMnmCD      As String        '균코드
    
    Dim sSusCd      As String        '항생제코드
    Dim sSusRe      As String        '결과값
    Dim strTmp      As String        'srst의 결과필드 저장값
    Dim sRCnt       As Long          '결과 카운트
    Dim iRCol       As Integer
    Dim i As Long
    
    Set objsSQL = New clsLISSqlStatistic
    
'    sDept = ""
'    If optBussDiv(0).Value Then     '외래
'        sOptDept = "2":        sDept = txtWardId.Text
'    ElseIf optBussDiv(1).Value Then '병동
'        sOptDept = "1":        sDept = txtWardId.Text
'    ElseIf optBussDiv(2).Value Then '전체
'        sOptDept = "0":        sDept = txtWardId.Text
'    End If
    
    sDept = txtWardId.Text
    sOptDept = IIf(optBussDiv(0).Value, "2", IIf(optBussDiv(1).Value, "1", IIf(optBussDiv(2).Value, "0", "0")))

'    SSQL = objsSQL.GetSensiResult(pDay, "")
'
'    Select Case sOptDept
'        Case "1"    '병동
'            SSQL = SSQL & objsSQL.GetSensiData(True, pDay, "", sDept)
'        Case "2"    '진료과
'            SSQL = SSQL & objsSQL.GetSensiData(False, pDay, "", sDept)
'    End Select
    
    SSQL = objsSQL.GetSensiResult_New(pDay, Format(txtDate2.Value, "yyyymmdd"), "", sOptDept, sDept)
    
    '균코드,항생제코드
    
    dsRst.Open SSQL, DBConn
    
    Do Until dsRst.EOF
        sMnmCD = Trim("" & dsRst.Fields("mnmcd").Value)             '균코드
        sRCnt = Val(Trim("" & dsRst.Fields("scnt").Value))          '항생제 갯수
                
        If objMn.Exists(sMnmCD) Then
            objMn.KeyChange sMnmCD
            objMn.Fields("mncnt") = Val(objMn.Fields("mncnt")) + 1
        Else
            objMn.AddNew sMnmCD, "1"
        End If
        
        If sRCnt > 0 Then
            For i = 1 To sRCnt
                strTmp = Trim("" & dsRst.Fields("srst" & i))
                
                sSusCd = Trim(medGetP(strTmp, 1, ";"))          '항생제코드
                sSusRe = Trim(medGetP(strTmp, 2, ";"))          '항생제결과
                If objDic.Exists(sMnmCD & COL_DIV & sSusCd) = False Then objDic.AddNew sMnmCD & COL_DIV & sSusCd, Join(Array("0", "0", "0", "0"), COL_DIV)
                objDic.KeyChange sMnmCD & COL_DIV & sSusCd
                Select Case sSusRe
                    Case "S"
                        objDic.Fields("scnt") = Val(objDic.Fields("scnt")) + 1
                    Case "R"
                        objDic.Fields("rcnt") = Val(objDic.Fields("rcnt")) + 1
                    Case "I"
                        objDic.Fields("icnt") = Val(objDic.Fields("icnt")) + 1
                End Select
                objDic.Fields("spccnt") = Val(objDic.Fields("spccnt")) + 1
            Next i
        End If
        dsRst.MoveNext
    Loop
    Set dsRst = Nothing
    Set objsSQL = Nothing
End Sub
Private Sub SusctrandDisplay()
    Dim objPro  As jProgressBar.clsProgress
    Dim strTmp  As String
    
    Dim lngStot     As Long
    Dim lngRtot     As Long
    Dim lngItot     As Long
    Dim lngSpcTot   As Long
    Dim lngTotal    As Long
    
    Dim lngCnt      As Long
    
    Dim blnTotal    As Boolean
    
    Dim ii      As Long
    Dim jj      As Long
    
    Set objPro = Nothing
    Set objPro = New jProgressBar.clsProgress
    
    With objPro
        .Container = Me
        .Left = Frame4.Left + pic.Left
        .Top = Frame4.Top + pic.Top
        .Width = pic.Width
        .Height = pic.Height
        .Max = objDic.RecordCount
'        .Choice = True
'        .SetMyForm Me
'        .XPos = Frame4.Left + pic.Left
'        .YPos = Frame4.Top + pic.Top
'        .XWidth = pic.Width
'        .YHeight = pic.Height
'        .Appearance = aPlate
'        .Max = objDic.RecordCount
    End With
    
    pic.ZOrder 0
    objDic.MoveFirst
    
    With tblSus
        .ReDraw = False
        Do Until objDic.EOF
            If .DataRowCnt >= .MaxRows Then .MaxRows = .MaxRows + 1
            .Row = .DataRowCnt + 1
            If strTmp <> objDic.Fields("mncd") Then
                If blnTotal = True Then
                    .Col = 1: .Value = "        균 별 합 계     ": .FontBold = True
'                    .Col = 3: .Value = lngStot & " 건 /  " & Format((lngStot / lngTotal) * 100, "00.0") & "%": .FontBold = True
'                    .Col = 4: .Value = lngRtot & " 건 /  " & Format((lngRtot / lngTotal) * 100, "00.0") & "%": .FontBold = True
'                    .Col = 5: .Value = lngItot & " 건 /  " & Format((lngItot / lngTotal) * 100, "00.0") & "%": .FontBold = True
'                    .Col = 6: .Value = lngTotal & " 건": .FontBold = True
                    .Col = 6: .Value = objMn.Fields("mncnt") & " 건"
                    If .DataRowCnt >= .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = .DataRowCnt + 1
                    .FontBold = False
                End If
                .Col = 1: .Value = GetName("C219", objDic.Fields("mncd"), "") ': .FontBold = False
                lngStot = 0: lngRtot = 0: lngItot = 0: lngTotal = 0
            End If
            
            
            
            .Col = 2: .Value = GetNameP("C221", objDic.Fields("suscd"), "S")

            objMn.KeyChange objDic.Fields("mncd")
            lngCnt = Val(objDic.Fields("scnt")) + Val(objDic.Fields("rcnt")) + Val(objDic.Fields("icnt"))
            If lngCnt > 0 Then
                .Col = 3: .Value = objDic.Fields("scnt") & " 건 /  " & Format((objDic.Fields("scnt") / lngCnt) * 100, "00.0") & "%"
                .Col = 4: .Value = objDic.Fields("rcnt") & " 건 /  " & Format((objDic.Fields("rcnt") / lngCnt) * 100, "00.0") & "%"
                .Col = 5: .Value = objDic.Fields("icnt") & " 건 /  " & Format((objDic.Fields("icnt") / lngCnt) * 100, "00.0") & "%"
            Else
                .Col = 3: .Value = objDic.Fields("scnt") & " 건 /  " & "00.0" & "%"
                .Col = 4: .Value = objDic.Fields("rcnt") & " 건 /  " & "00.0" & "%"
                .Col = 5: .Value = objDic.Fields("icnt") & " 건 /  " & "00.0" & "%"
            End If
'                .Col = 6: .Value = objDic.Fields("mncnt") & " 건"
            
'            .Col = 3: .Value = objDic.Fields("scnt") & " 건 /  " & Format((objDic.Fields("scnt") / objDic.Fields("spccnt")) * 100, "00.0") & "%"
'            .Col = 4: .Value = objDic.Fields("rcnt") & " 건 /  " & Format((objDic.Fields("rcnt") / objDic.Fields("spccnt")) * 100, "00.0") & "%"
'            .Col = 5: .Value = objDic.Fields("icnt") & " 건 /  " & Format((objDic.Fields("icnt") / objDic.Fields("spccnt")) * 100, "00.0") & "%"
'
'            .Col = 6: .Value = objDic.Fields("spccnt") & " 건"
            
            lngStot = lngStot + Val(objDic.Fields("scnt"))
            lngRtot = lngRtot + Val(objDic.Fields("rcnt"))
            lngItot = lngItot + Val(objDic.Fields("icnt"))
            
            lngTotal = lngTotal + Val(objDic.Fields("spccnt"))
            
            blnTotal = True
            
            ii = ii + 1
            If ii = objDic.RecordCount Then
                If .DataRowCnt >= .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .Value = "        균 별 합 계     ": .FontBold = True
'                .Col = 3: .Value = lngStot & " 건 /  " & Format((lngStot / lngTotal) * 100, "00.0") & "%": .FontBold = True
'                .Col = 4: .Value = lngRtot & " 건 /  " & Format((lngRtot / lngTotal) * 100, "00.0") & "%": .FontBold = True
'                .Col = 5: .Value = lngItot & " 건 /  " & Format((lngItot / lngTotal) * 100, "00.0") & "%": .FontBold = True
'                .Col = 6: .Value = lngTotal & " 건": .FontBold = True
                .Col = 6: .Value = objMn.Fields("mncnt") & " 건"
                .Row = .Row + 1
            End If
            strTmp = objDic.Fields("mncd")
            objPro.Value = ii
            objPro.Message = ii & " 번째 Display 중입니다."
            objDic.MoveNext
        Loop
        .ReDraw = True
    End With
    
    tblSus.ZOrder 0
    Set objPro = Nothing

End Sub

Private Sub sstResult_Click(PreviousTab As Integer)
    LisLabel3(1).Visible = True: cmdPopupList.Visible = True: txtSpcCd.Visible = True: lblSpcNm.Visible = True
    Call medClearTable(tblSus)
    Select Case sstResult.Tab
        Case 0: Frame3.ZOrder 0
        Case 1: Frame1.ZOrder 0
        Case 2: Frame4.ZOrder 0: txtAnti.Text = "": txtMic.Text = ""
                LisLabel3(1).Visible = False: cmdPopupList.Visible = False: txtSpcCd.Visible = False: lblSpcNm.Visible = False
    End Select
End Sub

Private Sub cmdStatics_Click()
    Call SusTrandQuery
End Sub

Private Function GetName(ByVal Cdindex As String, ByVal CdVal1 As String, ByVal Sus As String) As String
    Dim SSQL As String
    Dim RS   As Recordset
    
    SSQL = " SELECT field1,text1 FROM " & T_LAB032 & _
           " WHERE " & DBW("cdindex =", Cdindex) & " AND " & DBW("cdval1=", CdVal1)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    GetName = IIf(Not RS.EOF, IIf(Sus <> "", CdVal1 & Space(8 - Len(CdVal1)) & " " & RS.Fields("text1").Value & "", CdVal1 & Space(8 - Len(CdVal1)) & " " & RS.Fields("field1").Value & ""), CdVal1)

    Set RS = Nothing
End Function

Private Function GetNameP(ByVal Cdindex As String, ByVal CdVal1 As String, ByVal Sus As String) As String
    Dim SSQL As String
    Dim RS   As Recordset

    SSQL = " SELECT field1,text1 FROM " & T_LAB032 & " WHERE " & DBW("cdindex =", Cdindex) & " AND " & DBW("cdval1=", CdVal1)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    GetNameP = IIf(Not RS.EOF, IIf(Sus <> "", CdVal1 & Space(8 - Len(CdVal1)) & " " & RS.Fields("text1").Value & "", CdVal1 & Space(8 - Len(CdVal1)) & " " & RS.Fields("field1").Value & ""), CdVal1)
    Set RS = Nothing
End Function

Private Sub cmdSusPrint_Click()
    Call SuscTrand
End Sub

Private Sub SuscTrandHead()
    Dim strTmp  As String
    Dim ii      As Integer
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    lngCurYPos = 10

    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("항생제 감수성추이", PrtLeft, LineSpace * 3, Printer.ScaleWidth - PrtLeft, "C", "C", True)
    Printer.FontSize = 9: Printer.FontBold = False
    
    strTmp = Format(txtDate1.Value, "YYYY년 MM월 DD일") & " ~ " & Format(txtDate2.Value, "YYYY년 MM월 DD일")
    
    Call Print_Setting("조회기간 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    If optBussDiv(2).Value Then
        strTmp = "전 체 "
    ElseIf optBussDiv(1).Value Then
        strTmp = "병 동 [ " & txtWardId.Text & " - " & lblWardNm.Caption & " ]"
    Else
        strTmp = "외래 [ " & txtWardId.Text & " - " & lblWardNm.Caption & " ]"
    End If
    
    Call Print_Setting("조회유형 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")
    strTmp = Format(GetSystemDate, "YYYY년 MM월 DD일")
    Call Print_Setting("출 력 일 : " & strTmp, PrtLeft, LineSpace, Printer.ScaleWidth, "L", "C")

    
    
    
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
    
    Call SuscTrandBody("균명", "항생제명", "세부결과(S)", "세부결과(R)", "세부결과(I)", "검체수")
    
    Printer.DrawStyle = 0: Printer.DrawWidth = 6
    Printer.Line (PrtLeft, lngCurYPos)-(Printer.Width - PrtLeft, lngCurYPos)
End Sub
Private Sub SuscTrandBody(ByVal sSUS As String, ByVal sMic As String, ByVal sCnt1 As String, _
                          ByVal sCnt2 As String, ByVal sCnt3 As String, ByVal sCnt As String)
                           
    If lngCurYPos > Printer.ScaleHeight - 6 Then
        Printer.NewPage
        Call SuscTrandHead
    End If
   
    Call Print_Setting(sSUS, 5, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sMic, 35, LineSpace, 50, "L", "C", False)
    Call Print_Setting(sCnt1, 85, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sCnt2, 115, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sCnt3, 145, LineSpace, 30, "L", "C", False)
    Call Print_Setting(sCnt, 175, LineSpace, 35, "L", "C")
End Sub

Private Sub SuscTrand()
    Dim sSUS    As String
    Dim sMic    As String
    Dim sCnt1   As String
    Dim sCnt2   As String
    Dim sCnt3   As String
    Dim sCnt    As String
    
    
    Dim ii          As Integer
    
    If tblSus.DataRowCnt < 1 Then Exit Sub
    
'    Call P_PrtSet
    Call SuscTrandHead
    
    With tblSus
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 1:   sSUS = .Value
            .Col = 2:   sMic = .Value
            .Col = 3:   sCnt1 = .Value
            .Col = 4:   sCnt2 = .Value
            .Col = 5:   sCnt3 = .Value
            .Col = 6:   sCnt = .Value
            Call SuscTrandBody(sSUS, sMic, sCnt1, sCnt2, sCnt3, sCnt)
        Next
    End With
    
    Printer.EndDoc
End Sub

