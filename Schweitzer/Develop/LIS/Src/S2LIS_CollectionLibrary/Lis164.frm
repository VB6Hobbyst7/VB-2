VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm164BatchCol 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9195
   ClientLeft      =   -315
   ClientTop       =   420
   ClientWidth     =   14490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14490
   WindowState     =   2  '최대화
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00DBE6E6&
      Caption         =   "전체제외(&A)"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1380
      TabIndex        =   22
      Top             =   1650
      Width           =   1560
   End
   Begin VB.CheckBox chkPrintFg 
      BackColor       =   &H00DBE6E6&
      Caption         =   "출력안함"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8865
      TabIndex        =   21
      Top             =   450
      Width           =   1305
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Left            =   8475
      TabIndex        =   13
      Top             =   45
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "출력 옵션"
      LeftGab         =   100
   End
   Begin VB.Frame fraPrtOption 
      BackColor       =   &H00DBE6E6&
      Height          =   1320
      Left            =   8475
      TabIndex        =   14
      Top             =   270
      Width           =   5925
      Begin MedControls1.LisLabel lblPage 
         Height          =   255
         Left            =   4335
         TabIndex        =   20
         Top             =   975
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "부"
         Appearance      =   0
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   4020
         TabIndex        =   18
         Top             =   900
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCopy 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   345
         Left            =   3255
         TabIndex        =   17
         Top             =   915
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드 Only"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Top             =   675
         Width           =   3180
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드Lable And 채혈 리스트"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   420
         Visible         =   0   'False
         Width           =   3180
      End
      Begin MedControls1.LisLabel lblColList 
         Height          =   255
         Left            =   855
         TabIndex        =   19
         Top             =   945
         Visible         =   0   'False
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   450
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "채혈리스트 출력장수"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00F4F0F2&
      Caption         =   "실행(&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      Height          =   7050
      Left            =   75
      ScaleHeight     =   6990
      ScaleWidth      =   8355
      TabIndex        =   3
      Top             =   1950
      Width           =   8415
      Begin FPSpread.vaSpread tblPtList 
         Height          =   6990
         Left            =   0
         TabIndex        =   4
         Tag             =   "15109"
         Top             =   0
         Width           =   8340
         _Version        =   196608
         _ExtentX        =   14711
         _ExtentY        =   12330
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   23
         MaxRows         =   26
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis164.frx":0000
         VisibleCols     =   3
         VisibleRows     =   25
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00DBE6E6&
      Height          =   6270
      Left            =   8490
      ScaleHeight     =   6210
      ScaleWidth      =   5880
      TabIndex        =   5
      Top             =   2250
      Width           =   5940
      Begin MedControls1.LisLabel lblColNm 
         Height          =   330
         Left            =   345
         TabIndex        =   8
         Top             =   555
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPtCount 
         Height          =   330
         Left            =   345
         TabIndex        =   9
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin FPSpread.vaSpread tblCount 
         Height          =   5970
         Left            =   2415
         TabIndex        =   6
         Tag             =   "15109"
         Top             =   0
         Width           =   3465
         _Version        =   196608
         _ExtentX        =   6112
         _ExtentY        =   10530
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   14737632
         MaxCols         =   3
         MaxRows         =   18
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis164.frx":07D6
         VisibleCols     =   3
         VisibleRows     =   15
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   345
         TabIndex        =   31
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "♣ 채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   345
         TabIndex        =   32
         Top             =   1065
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "♣ 환자수"
         Appearance      =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   2400
         X2              =   2400
         Y1              =   0
         Y2              =   4770
      End
      Begin VB.Label Label4 
         BackColor       =   &H00DBE6E6&
         Caption         =   "명"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1620
         TabIndex        =   7
         Tag             =   "20104"
         Top             =   1515
         Width           =   270
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   75
      TabIndex        =   10
      Top             =   45
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "진료과 선택"
      LeftGab         =   100
   End
   Begin MSComctlLib.ProgressBar pbrPtCnt 
      Height          =   150
      Left            =   8595
      TabIndex        =   11
      Top             =   2010
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   300
      Left            =   8490
      TabIndex        =   12
      Top             =   1605
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "진행 상황"
      LeftGab         =   100
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   435
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1320
      Left            =   75
      TabIndex        =   23
      Top             =   255
      Width           =   8340
      Begin VB.TextBox txtCorpCd 
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
         Height          =   375
         Left            =   5145
         TabIndex        =   36
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdCorpList 
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
         Height          =   375
         Left            =   6000
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   35
         Top             =   240
         Width           =   315
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   360
         Left            =   2265
         TabIndex        =   25
         Top             =   255
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
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
         Height          =   375
         Left            =   1920
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txtDeptCd 
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
         Height          =   375
         Left            =   1065
         TabIndex        =   26
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdGetOrders 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6885
         Style           =   1  '그래픽
         TabIndex        =   24
         Tag             =   "15101"
         Top             =   690
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   375
         Left            =   1065
         TabIndex        =   28
         Top             =   720
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd  H:mm:ss"
         Format          =   109772800
         CurrentDate     =   36342.5951388889
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   29
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "부서코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   30
         Top             =   720
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "처방일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCorpNm 
         Height          =   360
         Left            =   6345
         TabIndex        =   34
         Top             =   255
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   4050
         TabIndex        =   37
         Top             =   240
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "거래처코드"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   4
      Left            =   75
      TabIndex        =   33
      Top             =   1575
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "환자리스트"
      Appearance      =   0
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00D8DEDA&
      FillStyle       =   0  '단색
      Height          =   330
      Index           =   1
      Left            =   8490
      Shape           =   4  '둥근 사각형
      Top             =   1920
      Width           =   5910
   End
End
Attribute VB_Name = "frm164BatchCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSQL                  As clsLISSqlCollection
Private objCollect              As clsLISCollectioin
'Private WithEvents objMyList    As clsPopUpList
Private WithEvents objMyList    As clsS2DLP
Attribute objMyList.VB_VarHelpID = -1

Private CleanFg                 As Boolean
Private blnInitFg               As Boolean
Private intPtCount              As Integer
Private intErrCount             As Integer

Private Const lngMaxRows = 25
Private Const lngRowHeight = 12
Public Event LastFormUnload()

Private Sub cmdClear_Click()
    Call ClearRtn(1)
    txtDeptCd.SetFocus
End Sub

'** 예수병원 추가루틴 By M.G.Choi 2005.04.08
'   거래처별 조회 가능하게 하기 위해
Private Sub cmdCorpList_Click()
    Dim strSQL  As String
    
    strSQL = " select custcode, custname " & _
             "   from oras1.sg1custt "
    
    '% 거래처정보 리스트를 팝업한다.
'    Set objMyList = New clsPopUpList
'    With objMyList
'        .Connection = DBConn
'        .FormCaption = "거래처 조회"
'
'        .ColumnHeaderText = "거래처코드;거래처명"
'        .Tag = "CorpID"
'         Call .LoadPopUp(strSQL)
'        If .SelectedString <> "" Then
'            txtCorpCd.Text = medGetP(.SelectedString, 1, ";")
'            lblCorpNm.Caption = medGetP(.SelectedString, 2, ";")
'        End If
'    End With
'    Set objMyList = Nothing
' 2009.04.17 양성현 팝업메뉴 변경
'% 병동코드 리스트를 팝업한다.
    Set objMyList = New clsS2DLP
    With objMyList
        .Caption = "거래처 조회"
        .HeadName = "거래처코드;거래처명"
        .Tag = "CorpID"
         Call .ListPop(strSQL, 2700, cmdWardList.Left)
        If .SelectedString <> "" Then
            txtCorpCd.Text = medGetP(.SelectedString, 1, ";")
            lblCorpNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    Set objMyList = Nothing



End Sub

Private Sub dtpToTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub


Private Sub Form_Activate()

    If blnInitFg Then Exit Sub
    
    txtCopy.Text = 1
    dtpToTime.Value = Format(GetSystemDate, "YYYY-MM-DD HH:MM:SS")
    CleanFg = True
    intErrCount = 0
    txtDeptCd.Text = ""
    txtDeptCd.SetFocus
    chkPrintFg.Value = 0
    optOption(1).Value = True
    
    blnInitFg = True
    
End Sub

Private Sub Form_Deactivate()
    Set objMyList = Nothing
End Sub

Private Sub Form_Load()

    Me.Show
    blnInitFg = False
    Set objSQL = New clsLISSqlCollection
    Set objCollect = New clsLISCollectioin
    
End Sub
Private Sub chkAll_Click()
    With tblPtList
        .Col = 1: .Col2 = 1
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .Value = chkAll.Value
        .BlockMode = False
    End With
End Sub

'& 출력 Option 선택
Private Sub chkPrintFg_Click()
    If chkPrintFg.Value = 1 Then
        optOption(0).Value = False
        optOption(1).Value = False
        fraPrtOption.Enabled = False
    Else
        optOption(1).Value = True
        fraPrtOption.Enabled = True
    End If
End Sub

'% 종료
Private Sub cmdExit_Click()
    Unload Me
    Set objMyList = Nothing
    Set objSQL = Nothing
    Set objCollect = Nothing
    If IsLastForm Then RaiseEvent LastFormUnload
    
End Sub

'% 일괄채혈 수행
Private Sub cmdGenerate_Click()
    Dim Resp        As VbMsgBoxResult
'    Dim SaveWardId  As String
    Dim SavePtId    As String
    Dim sWorkarea   As String
    Dim sAccdt      As String
    
    Dim sBuildCd    As String
    Dim sBuildNm    As String
    Dim sWorkDt     As String
    Dim sWorkTm     As String
    
    Dim iAccseq     As Long
    Dim SelCount    As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer

    Set objCollect = New clsLISCollectioin

    sWorkDt = Format(GetSystemDate, CS_DateDbFormat)
    sWorkTm = Format(GetSystemDate, CS_TimeDbFormat)

    Call objCollect.SetWardCol(sWorkDt, sWorkTm, txtDeptCd.Text)

    tblCount.Row = 0
    intErrCount = 0
    SelCount = 0
    SavePtId = ""

    'Locking...
    txtDeptCd.Enabled = False
    txtDeptCd.BackColor = &H8000000F
    cmdWardList.Enabled = False
    dtpToTime.Enabled = False
    cmdGetOrders.Enabled = False

    Call MouseRunning  '13
    
'    Dim objBuild As clsBasisData
    Dim strBuild As String
    
    With tblPtList
        For i = 1 To intPtCount
            .Row = i

            If pbrPtCnt.Value >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 1
            pbrPtCnt.Value = pbrPtCnt.Value + 1
            DoEvents

            '* 제외버튼 Check
            .Col = 1
            If .Value = 1 Then GoTo Skip

            SelCount = SelCount + 1

            '* 채혈수행
            .Col = 15
            If Trim(.Value) <> "" Then Call DoCollection(i)
            DoEvents

            .Col = 17
            If Trim(.Value) <> "" Then Call DoCollectionForBBS(i)

            DoEvents
            '* Delivery Location 별 Count
            .Col = 2
            For j = 1 To objCollect.ColCount
                Call objCollect.GetLabNumbers(j, sWorkarea, sAccdt, iAccseq, sBuildCd)
                
'                Set objBuild = Nothing
'                Set objBuild = New clsBasisData
                sBuildNm = GetBuildNm(sBuildCd)
'                Set objBuild = Nothing
                
'                Call ObjLISComCode.Building.KeyChange(sBuildCd)
'                sBuildNm = ObjLISComCode.Building.Fields("buildnm")

                For k = 1 To tblCount.DataRowCnt
                    tblCount.Row = k
                    tblCount.Col = 3
                    If tblCount.Value = sBuildCd Then
                        '* 검체수 Count
                        tblCount.Col = 2
                        tblCount.Text = CStr(Val(tblCount.Text) + 1)
                        GoTo NextCol
                    End If
                Next

                If tblCount.DataRowCnt >= tblCount.MaxRows Then tblCount.MaxRows = tblCount.MaxRows + 1
                tblCount.Row = tblCount.DataRowCnt + 1
                tblCount.Col = 1: tblCount.Value = sBuildNm
                tblCount.Col = 2: tblCount.Text = "1"
                tblCount.Col = 3: tblCount.Value = sBuildCd
NextCol:
            Next

            '* 환자수 Count
            .Row = i
            .Col = 3
            If SavePtId <> Trim(.Value) Then
                lblPtCount.Caption = Val(lblPtCount.Caption) + 1
                SavePtId = .Value
            End If

            '* 채혈 Class Initialize
            objCollect.InitRtn
            DoEvents
Skip:
        Next

        '채혈자
        lblColNm.Caption = ObjSysInfo.EmpId

    End With

    If SelCount = 0 Then
        MouseDefault  '0
        Call cmdClear_Click
        MsgBox "처리된 데이타가 없습니다..", vbInformation, "Message"
        Exit Sub
    End If
    
    pbrPtCnt.Value = pbrPtCnt.Max
    DoEvents

    Call MouseDefault  '0

    If intErrCount > 0 Then
        MsgBox CStr(intErrCount) & "건의 오류가 발생했습니다.."
    Else
        If optOption(0).Value Then
            Resp = MsgBox("모두 정상적으로 채취처리 되었습니다.." & vbCrLf & _
                                    "채취리스트를 지금 출력하시겠습니까 ? ", vbYesNo, "채취리스트 출력")
            If Resp = vbYes Then
                For i = 1 To tblCount.DataRowCnt
                    tblCount.Row = i
                    tblCount.Col = 3:  sBuildCd = tblCount.Value
                    tblCount.Col = 1:  sBuildNm = tblCount.Value
                    For j = 1 To Val(txtCopy.Text)
                        Call PrintColList(txtDeptCd.Text, lblWardNm.Caption, sWorkDt, sWorkTm, sBuildCd, sBuildNm)
                    Next
                Next
            End If
        Else
            Call MsgBox("모두 정상적으로 채취처리 되었습니다..", vbInformation, "메세지")
        End If

        Call ClearRtn(0)
        txtDeptCd.SetFocus
   End If

End Sub

'& 채혈 클래스 objCollect 를 이용하여 해당 환자들의 처방을 채혈수행한다.
Private Sub DoCollection(ByVal Row As Long)
    Dim tmpRs       As Recordset
    Dim tmpDate     As String
    Dim tmpTime     As String

    Dim SqlStmt     As String
    
    Dim tmpData()   As String
    Dim i           As Integer

    Dim Success     As Boolean

    ReDim tmpData(0 To 16)

    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)
    
    With tblPtList
        .Row = Row
        .Col = 3:  tmpData(1) = .Value                          '환자ID
    End With
        
    ' 처방내역 검색
    SqlStmt = objSQL.SqlReadOutOrder(tmpData(1), tmpDate, tmpTime, txtDeptCd.Text)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then
        Set tmpRs = Nothing
        Exit Sub
    End If
    

    With tblPtList
        .Row = Row
        tmpData(0) = Mid(Format(Now, "YYYY"), 4)
        .Col = 3:  tmpData(1) = .Value                          '환자ID
       
        .Col = 4:  tmpData(2) = .Value                                      '환자명
        .Col = 14: tmpData(3) = .Value                                      '환자성별
        .Col = 7:
            If IsDate(Format(.Value, CS_DateMask)) Then
               tmpData(4) = DateDiff("y", Format(.Value, CS_DateMask), Now) '환자일령
            Else
               tmpData(4) = 50000       '생년월일이 정확하지 않을경우 Max값 - 2000.6.16 김미경
            End If
        .Col = 8:   tmpData(5) = .Value                                     '입원일
        tmpData(6) = Format(Now, CS_DateDbFormat)                           '입력일
        tmpData(7) = Format(Now, CS_TimeDbFormat)                           '입력시간
        tmpData(8) = ObjSysInfo.EmpId                                       '입력자
        tmpData(9) = ""                                                     '원접수번호
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat)            '채혈일
        objCollect.ColTm = Format(GetSystemDate, "HHMMSS")              '채혈일
        tmpData(11) = ObjSysInfo.EmpId                                      '채혈자
        .Col = 9:   tmpData(12) = .Value                                    '병동ID
        .Col = 12:  tmpData(13) = .Value                                    '병실ID
        .Col = 12:  tmpData(14) = .Value                                    '호실ID
        tmpData(15) = ""                                                    '침상ID
        tmpData(16) = ObjSysInfo.BuildingCd                                 '** 채혈이 수행되는 건물코드
        
        Call objCollect.SetColData(tmpData)
    End With

    ReDim tmpData(0 To 20)

    With tmpRs
        For i = 1 To .RecordCount
            tmpData(0) = ObjSysInfo.BuildingCd
            
            tmpData(1) = Trim("" & .Fields("WorkArea").Value)   'WorkArea
            tmpData(2) = Trim("" & .Fields("SpcCd").Value)      'SpcCd
            tmpData(3) = Trim("" & .Fields("StoreCd").Value)    'StoreCd
            tmpData(4) = Trim("" & .Fields("StatFg").Value)
            tmpData(5) = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateLongMask) & " " & _
                         Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)        '희망채취일시
            
            tmpData(6) = Trim("" & .Fields("TestDiv").Value)    'TestDiv
            tmpData(7) = Trim("" & .Fields("MultiFg").Value)    'MultiFg
            tmpData(8) = Trim("" & .Fields("SpcGrp").Value)     'SpcGrp
            tmpData(9) = Trim("" & .Fields("OrdDt").Value)      'OrdDt
            tmpData(10) = Trim("" & .Fields("OrdNo").Value)     'OrdNo
            tmpData(11) = Trim("" & .Fields("OrdSeq").Value)    'OrdSeq
            tmpData(12) = Trim("" & .Fields("OrdCd").Value)     'OrdCd
            tmpData(13) = Trim("" & .Fields("DeptCd").Value)    'DeptCd
            tmpData(14) = Trim("" & .Fields("OrdDoct").Value)   'OrdCd
            tmpData(15) = Trim("" & .Fields("MajDoct").Value)   'OrdCd
            tmpData(16) = Trim("" & .Fields("AbbrNm5").Value)   '처방 약어명
            tmpData(17) = Trim("" & .Fields("LabelCnt").Value)  '라벨출력장수
            
'            Call ObjLISComCode.LisItem.KeyChange(Trim("" & .Fields("TestCd").Value))
            tmpData(18) = GetLabDiv(Trim("" & .Fields("TestCd").Value)) 'ObjLISComCode.LisItem.Fields("labdiv")    'LabDiv

'            Call ObjLISComCode.LisSpc.KeyChange(tmpData(2))         '검체코드
'            tmpData(19) = ObjLISComCode.LisSpc.Fields("spcabbr")    '검체약어명
'            tmpData(20) = ObjLISComCode.LisSpc.Fields("labrange")   '미생물접수번호범위
            
            Call GetSpcInfo(tmpData(2), tmpData(19), tmpData(20))
'-----내가
            Call objCollect.SetAddLabCollect(tmpData)
            
            .MoveNext
            
        Next
    End With

    ' 채혈 수행
    Success = objCollect.DoCollection
    If Not Success Then
        tblPtList.Row = Row: tblPtList.Row2 = Row
        tblPtList.Col = -1
        tblPtList.BlockMode = True
        tblPtList.ForeColor = &HFF&       '빨간색
        tblPtList.BlockMode = False
        intErrCount = intErrCount + 1
    End If
    
    Set tmpRs = Nothing

End Sub

Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b"
    strSQL = strSQL & " where " & DBW("b.cdindex=", LC3_WorkArea)
    strSQL = strSQL & " and a.workarea=b.cdval1"
    strSQL = strSQL & " and " & DBW("a.testcd=", vTestCd)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
    GetLabDiv = Rs.Fields("field2").Value & ""
    End If
    Set Rs = Nothing
End Function

Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbr As String, _
                            ByRef vLabRng As String)
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select  a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm,  " & _
            " a.field1 multifg, a.field2 spcgrp, b.field2 labrange " & _
            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
            " where " & DBW("a.cdindex =", LC3_Specimen) & _
            " and " & DBW("a.cdval1=", vSpcCd) & _
            " and    " & DBJ("b.cdindex ='C217'") & _
            " and    " & DBJ("b.cdval1  =* a.field2")

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
    
    vSpcAbbr = Rs.Fields("spcabbr").Value & ""
    vLabRng = Rs.Fields("labrange").Value & ""
    End If
    
    Set Rs = Nothing
End Sub

'% 병동별로 현재 입원중인 환자들의 처방을 검색한다.
Private Sub cmdGetOrders_Click()
    Dim objStatus   As jProgressBar.clsProgress
    Dim Rs          As Recordset
    Dim SqlStmt     As String
    Dim tmpPtId     As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim tmpOrdDiv   As String
    Dim tmpSpcCd    As String
    Dim i           As Integer

    If Trim(txtDeptCd.Text) = "" Then
        MsgBox "부서코드를 입력하세요.", vbInformation, "진료과선택"
        txtDeptCd.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCorpCd.Text) = "" Then
        MsgBox "거래처코드를 입력하세요.", vbInformation, "거래처선택"
        txtCorpCd.SetFocus
        Exit Sub
    End If
    
    '2001-11-07 : 오래된 병동일괄채혈 내역 삭제 --------------------------------------------------
   
    
    Set objStatus = New jProgressBar.clsProgress
    
    With objStatus
        .Container = Me
        .Left = LisLabel1.Left
        .Top = LisLabel1.Top
        .Width = LisLabel1.Width
        .Height = 280
        .Message = "오래된 병동일괄 채취내역을 삭제하고 있습니다..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = LisLabel1.Width
'        .XPos = LisLabel1.Left
'        .YPos = LisLabel1.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "오래된 병동일괄채취 내역을 삭제하고 있습니다..."
'        .Max = 100
'        .Value = 50
    End With

    Set objCollect = New clsLISCollectioin
    If Not objCollect.Archive_WardColData(txtDeptCd.Text) Then
        MsgBox "병동일괄채취 내역 Archive중 오류가 발생했습니다." & vbCrLf & _
                "전산실 혹은 임상병리과로 연락바랍니다. (☎" & ObjSysInfo.HelpLine & ")", vbCritical, "오류발생"
    End If
    '---------------------------------------------------------------------------------------------
    
    Call TableClear(1)
    
    Call MouseRunning
    
    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)
    
    '** 수정 예수병원 출장건지 관련 거래처코드 조회 추가 By M.G.Choi
    '-- 원본 ======================================================
'    SqlStmt = objSQL.SqlOutOrder(tmpDate, tmpTime, txtDeptCd.Text)
    '==============================================================
    
    '-- 수정 ======================================================
    SqlStmt = objSQL.SqlOutOrder_New(tmpDate, tmpTime, txtDeptCd.Text, txtCorpCd.Text)
    '==============================================================
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        MsgBox "채취대상이 없습니다..", vbInformation, "외래채취"
        cmdGenerate.Enabled = False
        MouseDefault
        Set objStatus = Nothing
        Set objCollect = Nothing
        Exit Sub
    End If

    With tblPtList
        .ReDraw = False
        .MaxRows = 0
'        objStatus.Value = 0
        objStatus.Max = Rs.RecordCount
        
        If Rs.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
        Else
            .MaxRows = Rs.RecordCount
        End If
        .Row = 1
        intPtCount = 0

        For i = 1 To Rs.RecordCount
            objStatus.Value = i
            objStatus.Message = "처방을 검색하고 있습니다...(" & i & " 건)"
            
            If tmpPtId <> Trim(Rs.Fields("PtId").Value & "") Then
'
                intPtCount = intPtCount + 1
                .Row = intPtCount
                .Col = 2: .Text = "" & Rs.Fields("DeptCd").Value    '병동ID
                .Col = 3: .Text = "" & Rs.Fields("PtId").Value     '환자ID
                .Col = 4: .Text = "" & Rs.Fields("PtNm").Value   '성명
                .Col = 7: .Text = "" & Rs.Fields("DOB").Value    '생년월일
                .Col = 14:
                
                .Text = Trim("" & Rs.Fields("Sex").Value)
                If IsNumeric(.Text) Then
                    .Text = Choose((Val(.Text) Mod 2) + 1, "F", "M")
                End If
                tmpPtId = "" & Rs.Fields("PtId").Value
            End If
            
            If P_PayDtUsed And (Trim(Rs.Fields("paydt").Value & "") = "") And Rs.Fields("orddiv").Value & "" = LIS_ORDDIV Then
                '-- 원본 =======================
                ' - 모르겠음 첨 로딩 시 색깔변경 안됨
'                .Col = 2: .Col2 = .MaxCols
'                .BlockMode = True
'                .ForeColor = DCM_LightRed
'                .BlockMode = False
                '===============================
                
                .Col = 2: .Col2 = .MaxCols
                .ForeColor = DCM_LightRed
                .Col = 3: .Col2 = .MaxCols
                .ForeColor = DCM_LightRed
                .Col = 4: .Col2 = .MaxCols
                .ForeColor = DCM_LightRed
                .Col = 5: .Col2 = .MaxCols
                .ForeColor = DCM_LightRed
                .Col = 6: .Col2 = .MaxCols
                .ForeColor = DCM_LightRed
            Else
                .Col = 2: .Col2 = .MaxCols
                .BlockMode = True
                .ForeColor = vbBlack
                .BlockMode = False
            End If
            
            .Col = 9: .Text = "" & Rs.Fields("DeptCd").Value  '진료과
            .Col = 10: .Text = "" & Rs.Fields("OrdDoct").Value '처방의
            .Col = 11: .Text = "" & Rs.Fields("MajDoct").Value '주치의

            tmpStatFg = "" & Rs.Fields("StatFg").Value      '응급여부
            tmpOrdDiv = "" & Rs.Fields("orddiv").Value             '처방구분
            tmpSpcCd = "" & Rs.Fields("SpcCd").Value               '검체

            If tmpStatFg = "1" Then
                .Col = 5
                If InStr(1, .Value, Rs.Fields("SpcNm").Value) = 0 Then
                    .Value = .Value & Rs.Fields("SpcNm").Value & ", "     '응급검체
                End If
            Else
                .Col = 6
                If InStr(1, .Value, Rs.Fields("SpcNm").Value) = 0 Then
                    .Value = .Value & Rs.Fields("SpcNm").Value & ", "
                End If
            End If
'
            Select Case tmpOrdDiv
            Case LIS_ORDDIV:   '임상
                .Col = 15: .ForeColor = vbRed: .Text = "√"     '처방구분√※
            Case BBS_ORDDIV:   '혈액
                .Col = 17: .ForeColor = vbRed: .Text = "√"     '처방구분√※

            End Select
            .Col = 19: .Text = Format(GetSystemDate, "YY-MM-DD")
            .Col = 20: .Text = Format(GetSystemDate, "HH:MM")

            Rs.MoveNext
        Next


        pbrPtCnt.Min = 0
        pbrPtCnt.Max = .DataRowCnt + 10
        pbrPtCnt.Value = 0

        .Row = 1: .Row2 = .MaxRows
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
        .ReDraw = True
    End With
        
    cmdGenerate.Enabled = True
    CleanFg = False
    Set objStatus = Nothing
    Set objCollect = Nothing
    Set Rs = Nothing

    Call MouseDefault

End Sub

' 기준시간이 변경되면 Clear
Private Sub dtpToTime_Change()
    If Not CleanFg Then Call TableClear(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set objSQL = Nothing
    Set objCollect = Nothing
    Set objMyList = Nothing
End Sub

Private Sub optOption_Click(Index As Integer)
    
    Select Case Index
        Case 0, 2: txtCopy.Text = 1
                   txtCopy.Enabled = True
        Case 1:    txtCopy.Text = 0
                   txtCopy.Enabled = False
    End Select

End Sub

Private Sub cmdWardList_Click()
'% 병동코드 리스트를 팝업한다.
'    Set objMyList = New clsPopUpList
'    With objMyList
'        .Connection = DBConn
'        .FormCaption = "부서코드 조회"
'
'        .ColumnHeaderText = "부서코드;부서명"
'        .Tag = "WardID"
'         Call .LoadPopUp(objSQL.SqlGetBatchDept) ', 2700, cmdWardList.Left)
'        If .SelectedString <> "" Then
'            txtDeptCd.Text = medGetP(.SelectedString, 1, ";")
'            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
'        End If
'    End With
'    Set objMyList = Nothing

' 2009.04.17 양성현 팝업메뉴 변경
'% 병동코드 리스트를 팝업한다.
    Set objMyList = New clsS2DLP
    With objMyList
        .Caption = "부서코드 조회"
        .HeadName = "부서코드,부서명"
        .Tag = "WardID"
'         Call .ListPop(objSQL.SqlGetBatchDept, 2700, cmdWardList.Left)
         Call .ListPop(objSQL.SqlGetBatchDept, 2700, cmdWardList.Left)
        If .SelectedString <> "" Then
            txtDeptCd.Text = medGetP(.SelectedString, 1, ";")
            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    Set objMyList = Nothing




End Sub


Private Sub ClearRtn(ByVal intOpt As Integer)
    'Unlocking...
    txtDeptCd.Enabled = True
    txtDeptCd.BackColor = vbWhite
    cmdWardList.Enabled = True
    dtpToTime.Enabled = True
    cmdGetOrders.Enabled = True
    cmdGenerate.Enabled = False

    txtDeptCd.Text = ""
    lblWardNm.Caption = ""
    dtpToTime.Value = Format(Now, "YYYY/MM/DD HH:MM:SS")
    pbrPtCnt.Value = 0
    chkPrintFg = 0
    optOption(1).Value = True
    intErrCount = 0
    Call TableClear(intOpt)

End Sub


'% Table들을 Clear한다
Private Sub TableClear(ByVal intOpt As Integer)
    tblPtList.MaxRows = 0
    tblPtList.MaxRows = 50
    If intOpt = 1 Then
        lblColNm.Caption = ""
        lblPtCount.Caption = ""
        tblCount.MaxRows = 0
        tblCount.MaxRows = 50
        CleanFg = True
    End If
End Sub

Private Sub PrintColList(ByVal pDeptCd As String, ByVal pWardNm As String, ByVal pWorkDt As String, _
                        ByVal pWorkTm As String, ByVal pBuildCd As String, ByVal pBuildNm As String)

    Dim MyReport As clsWardColList
    Dim strTitleNm As String
    
    strTitleNm = "외래 채취 리스트"
    Set MyReport = New clsWardColList
    With MyReport
        .WardId = pDeptCd
        .WardNm = pWardNm
        .WorkDt = pWorkDt
        .WorkTm = pWorkTm
        .BuildCd = pBuildCd
        .BuildNm = pBuildNm
        .TestDiv = "0"  'chkTestdiv.Value
        .TitleNm = strTitleNm
        .SetCrpt CReport
        Call .Print_ColList
    End With

    Set MyReport = Nothing

End Sub

Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim Rs          As Recordset
    Dim tmpToolTip  As String
    Dim strDeptCd   As String
    Dim strPtid     As String
    Dim strOrdDate  As String
    Dim strWardId   As String
    Dim strBBSORDCd As String
    Dim strLISORDCd As String

    If Row = 0 Then Exit Sub

    tmpToolTip = vbCrLf

    With tblPtList
        .Row = Row
        
        .Col = 2: If Trim(.Value) = "" Then Exit Sub
        
        .Col = 4: tmpToolTip = tmpToolTip & "  " & .Value & vbCrLf & vbCrLf    '환자명
        .Col = 5: tmpToolTip = tmpToolTip & "  응급검체 : " & .Value & vbCrLf  '응급검체
        .Col = 6: tmpToolTip = tmpToolTip & "  일반검체 : " & .Value & vbCrLf  '일반검체
        
        '-- ToolTip 추가사항 : 검사항목 Display
        .Col = 3: strPtid = Trim(.Value)
        strOrdDate = Format(dtpToTime.Value, CS_DateDbFormat)
        strDeptCd = Trim(txtDeptCd.Text)
        
        Set Rs = New Recordset
        Rs.Open objSQL.DeptCd_ORDCD(strPtid, strOrdDate, strDeptCd), DBConn
        
        If Rs.BOF = False Then
            Do Until Rs.EOF = True
                Select Case Rs.Fields("orddiv").Value & ""
                    Case "B"
                        strBBSORDCd = strBBSORDCd & Rs.Fields("abbrnm5").Value & "," '혈액은행 검사항목
                    Case "L"
                        strLISORDCd = strLISORDCd & Rs.Fields("abbrnm5").Value & "," '임상병리 검사항목
                End Select
                Rs.MoveNext
            Loop
        End If
        
        If strBBSORDCd <> "" Then
                tmpToolTip = tmpToolTip & "  혈액은행 검사항목 : " & strBBSORDCd & vbCrLf  '혈액은행 검사항목
        ElseIf strLISORDCd <> "" Then
                tmpToolTip = tmpToolTip & "  임상병리 검사항목 : " & strLISORDCd & vbCrLf  '임상병리 검사항목
        End If
        
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 5000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
    Set Rs = Nothing
End Sub

Private Sub txtCorpCd_Change()
    If Not CleanFg Then Call TableClear(1)
End Sub

Private Sub txtCorpCd_GotFocus()
    With txtCorpCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCorpCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objMyList Is Nothing Then Call cmdCorpList_Click
    End If
End Sub

Private Sub txtCorpCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        If txtCorpCd.Text = "" Then
            txtCorpCd.SetFocus
            Exit Sub
        Else
            
            Dim strCorp As String
            
            strCorp = GetCorpNm(txtCorpCd.Text)
            
            If strCorp = "" Then
                MsgBox "거래처코드를 확인하세요.", vbExclamation
                txtCorpCd.Text = ""
                Call cmdCorpList_Click
                Exit Sub
            Else
                lblCorpNm.Caption = strCorp
                SendKeys "{TAB}"
            End If
        End If
    End If
End Sub

Private Function GetCorpNm(ByVal pCorpCd As String) As String
    Dim strSQL  As String
    Dim Rs      As New ADODB.Recordset
    
    strSQL = " select custname from oras1.sg1custt " & _
             "  where custcode = " & DBS(pCorpCd)
             
    Rs.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        GetCorpNm = Rs.Fields("custname").Value & ""
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Function

Private Sub txtDeptCd_Change()
    If Not CleanFg Then Call TableClear(1)
End Sub

Private Sub txtDeptCd_GotFocus()
    With txtDeptCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDeptCd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then
        If objMyList Is Nothing Then Call cmdWardList_Click
    End If

End Sub

Private Sub txtDeptCd_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        If txtDeptCd.Text = "" Then
            txtDeptCd.SetFocus
            Exit Sub
        Else
            
'            Dim objDept As clsBasisData
            Dim strDept As String
            
'            Set objDept = New clsBasisData
            strDept = GetDeptNm(txtDeptCd.Text)
'            Set objDept = Nothing
            
            If strDept = "" Then
                MsgBox "부서코드를 확인하세요.", vbExclamation
                txtDeptCd.Text = ""
                Call cmdWardList_Click
                Exit Sub
            Else
                lblWardNm.Caption = strDept
                SendKeys "{TAB}"
            End If
            
'            If Not ObjLISComCode.DeptCd.Exists(txtDeptCd.Text) Then
'                MsgBox "부서 코드를 확인하세요.."
'                txtDeptCd.Text = ""
'                Call cmdWardList_Click
'                Exit Sub
'            Else
'                ObjLISComCode.DeptCd.KeyChange txtDeptCd.Text
'                lblWardNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'                SendKeys "{TAB}"
'            End If
        End If
    End If
End Sub

Private Sub DoCollectionForBBS(ByVal Row As Long)
'
    Dim strPtid     As String       '환자id
    Dim strPtnm     As String       '환자명
    Dim strColID    As String      '채혈자
    Dim strColDt    As String      '채혈일
    Dim strColTm    As String      '채혈일시
    Dim strHosilid  As String
    Dim strStatFg   As String
    Dim lngErCnt    As Long
    Dim lngGcnt     As Long
    Dim lngBldRow   As Long
    Dim j           As Long
    Dim sWorkDt     As String
    Dim sWorkTm     As String
    Dim strBBS_BldCd As String
    
    strBBS_BldCd = "10"
If P_IncludeBBSSystem Then

    Dim objDIC As New clsDictionary
    Dim objBBSCollect As New clsBBSCollection
    sWorkDt = Format(GetSystemDate, CS_DateDbFormat)
    sWorkTm = Format(GetSystemDate, CS_TimeDbFormat)
    
    Call objBBSCollect.SetWardCol(txtDeptCd.Text, sWorkDt, sWorkTm)

    With tblPtList
        .Row = Row
        .Col = 3: strPtid = .Value
        .Col = 4: strPtnm = .Value
        .Col = 5
        If .Value = "※" Then   '응급
            lngErCnt = lngErCnt + 1
        Else
            lngGcnt = lngGcnt + 1
        End If
        .Col = 23:
            strStatFg = IIf(.Value = "1", "1", "")
        .Col = 12:  strHosilid = Trim(.Value)
        .Col = 19:  strColDt = Format(.Text, "YYYYMMDD")
        .Col = 20:  strColTm = Format(.Text, "HHMMss")
        strColID = gEmpId
        objDIC.Clear
        objDIC.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"

        objDIC.AddNew strPtid, Join(Array(strPtnm, strColDt, strColTm, strColID, _
                                    enBussDiv.BussDiv_OutPatient, ObjSysInfo.BuildingCd, strHosilid, strStatFg), COL_DIV)
        
            
        If objDIC.RecordCount > 0 Then
            objBBSCollect.WardId = txtDeptCd.Text
            If objBBSCollect.Set_Collect(objDIC, ObjSysInfo.BuildingCd, , True) Then      '일괄채혈내역생성
'                Call ObjLISComCode.Building.KeyChange(ObjSysInfo.BuildingCd)

                lngBldRow = 0
                For j = 1 To tblCount.DataRowCnt
                    tblCount.Row = j: tblCount.Col = 3
                    If tblCount.Value = ObjSysInfo.BuildingCd Then
                        lngBldRow = j
                        Exit For
                    End If
                Next

                If lngBldRow = 0 Then lngBldRow = tblCount.DataRowCnt + 1
                tblCount.Row = lngBldRow
                
'                Dim objBld As clsBasisData
                Dim strBld As String
                
'                Set objBld = New clsBasisData
                strBld = GetBuildNm(ObjSysInfo.BuildingCd)
'                Set objBld = Nothing
                
                tblCount.Col = 1: tblCount.Text = strBld 'ObjLISComCode.Building.Fields("buildnm")
                tblCount.Col = 2: tblCount.Text = Val(tblCount.Text) + 1
                tblCount.Col = 3: tblCount.Text = ObjSysInfo.BuildingCd

                Dim objBar As New clsDictionary

                Set objBar = objBBSCollect.BldDic
                If objBar.RecordCount > 0 Then
                    BarCode_Print objBar
                End If
            End If
        End If
    End With

    Set objBBSCollect = Nothing
    Set objDIC = Nothing
    Set objBar = Nothing
    
End If

End Sub

Private Sub BarCode_Print(objDIC As clsDictionary)
    Dim strPtid     As String
    Dim strPtnm     As String
    Dim strColDt    As String
    Dim strColTm    As String
    Dim strSpcNo    As String
    Dim strAccSeq   As String         'SpcYy-SpcNo 형태의 검체번호
    Dim HosilId     As String
    Dim strStatFg   As String
    Dim strBarW_H   As String
    Dim objBar      As clsBarcode
    
    Set objBar = New clsBarcode
    
'    Set objBAR.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields

    objDIC.MoveFirst

    Do Until objDIC.EOF
        strPtid = medGetP(objDIC.GetString, 1, COL_DIV)
        strPtnm = medGetP(objDIC.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDIC.GetString, 3, COL_DIV)
        strColDt = medGetP(objDIC.GetString, 4, COL_DIV)
        strColDt = Format(Mid(strColDt, 5, 4), "##/##")
        strColTm = Mid(medGetP(objDIC.GetString, 5, COL_DIV), 1, 4)
        strColTm = Format(strColTm, "##:##")
        HosilId = medGetP(objDIC.GetString, 6, COL_DIV)
        strStatFg = medGetP(objDIC.GetString, 7, COL_DIV)
        
        If HosilId <> "" Then
            strBarW_H = txtDeptCd & "/" & HosilId
        Else
            strBarW_H = txtDeptCd
        End If
        
        '검체번호 출력 : 2001.2.8 추가
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '바코드 출력
        'ObjBBSComCode.BarInfo
        objBar.Label_PrintOut _
                        BBSName, "XM", "", strAccSeq, strSpcNo, strPtid, _
                        strPtnm, "", "", strStatFg, strBarW_H, _
                        strColDt, strColTm, "", Val(txtCopy)

        objDIC.MoveNext
    Loop
    
    Set objBar = Nothing
End Sub
