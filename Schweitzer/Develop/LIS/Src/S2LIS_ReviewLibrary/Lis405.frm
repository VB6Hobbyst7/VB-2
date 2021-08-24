VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Begin VB.Form frm405LegacyData 
   Caption         =   "과거결과 조회"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   14775
   WindowState     =   2  '최대화
   Begin VB.PictureBox picPtList 
      Align           =   3  '왼쪽 맞춤
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  '자동
      Height          =   8130
      Left            =   0
      ScaleHeight     =   8070
      ScaleWidth      =   3015
      TabIndex        =   16
      Top             =   900
      Width           =   3075
      Begin VB.Frame fraSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search"
         Height          =   630
         Left            =   45
         TabIndex        =   18
         Tag             =   "136"
         Top             =   645
         Width           =   3060
         Begin VB.OptionButton optSort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1680
            TabIndex        =   21
            Tag             =   "15304"
            Top             =   285
            Width           =   495
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2205
            TabIndex        =   20
            Tag             =   "15305"
            Top             =   270
            Width           =   810
         End
         Begin VB.TextBox txtSearchKey 
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Width           =   1470
         End
      End
      Begin VB.ListBox lstPtList 
         Appearance      =   0  '평면
         BackColor       =   &H00EEFFEC&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   6510
         Left            =   30
         TabIndex        =   17
         Top             =   1395
         Width           =   3060
      End
      Begin VB.Label lblPtList 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Patient List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         TabIndex        =   22
         Tag             =   "106"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FCEFE9&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   14775
      TabIndex        =   3
      Top             =   0
      Width           =   14775
      Begin VB.Frame Frame1 
         BackColor       =   &H00FCEFE9&
         Height          =   885
         Left            =   45
         TabIndex        =   38
         Top             =   -45
         Width           =   6345
         Begin VB.TextBox txtPtId 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   990
            MaxLength       =   10
            TabIndex        =   41
            Top             =   210
            Width           =   1440
         End
         Begin VB.CheckBox chkPtList 
            BackColor       =   &H00FCEFE9&
            Caption         =   "환자검색 리스트(&S)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Tag             =   "40101"
            Top             =   585
            Width           =   2445
         End
         Begin VB.CommandButton cmdRefresh 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Re&fresh"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5175
            Style           =   1  '그래픽
            TabIndex        =   39
            Tag             =   "128"
            Top             =   435
            Width           =   1065
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   285
            Left            =   3315
            TabIndex        =   42
            Top             =   195
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyy-MM-dd"
            Format          =   87687171
            CurrentDate     =   36328
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   285
            Left            =   3315
            TabIndex        =   43
            Top             =   525
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyy-MM-dd"
            Format          =   87687171
            CurrentDate     =   36328
         End
         Begin VB.Label lblPtId 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "환자 ID : "
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   135
            TabIndex        =   46
            Tag             =   "105"
            Top             =   285
            Width           =   945
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "From"
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
            Left            =   2775
            TabIndex        =   45
            Tag             =   "40105"
            Top             =   255
            Width           =   510
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "To"
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
            Left            =   3015
            TabIndex        =   44
            Tag             =   "40110"
            Top             =   585
            Width           =   270
         End
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "&Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13080
         TabIndex        =   5
         Tag             =   "40102"
         Top             =   90
         Width           =   1410
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "종 료 (&X)"
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
         Left            =   13080
         TabIndex        =   4
         Tag             =   "128"
         Top             =   465
         Width           =   1410
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   300
         Left            =   7575
         TabIndex        =   6
         Top             =   120
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   529
         BackColor       =   16703181
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
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblVerifierNm 
         Height          =   300
         Left            =   10905
         TabIndex        =   7
         Top             =   135
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   529
         BackColor       =   16703181
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel lblVerifyDt 
         Height          =   300
         Left            =   10905
         TabIndex        =   8
         Top             =   465
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   529
         BackColor       =   16703181
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
         Caption         =   ""
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7710
         TabIndex        =   47
         Top             =   525
         Width           =   585
      End
      Begin VB.Label lblSexAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성별/나이 : "
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
         Left            =   6465
         TabIndex        =   14
         Tag             =   "108"
         Top             =   510
         Width           =   1185
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성      명 : "
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
         Left            =   6510
         TabIndex        =   13
         Tag             =   "103"
         Top             =   180
         Width           =   1110
      End
      Begin VB.Label lblRptTm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보고일시 : "
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
         Left            =   9735
         TabIndex        =   12
         Tag             =   "40108"
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   8655
         TabIndex        =   11
         Top             =   510
         Width           =   345
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9255
         TabIndex        =   10
         Top             =   495
         Width           =   60
      End
      Begin VB.Label lblVerifier 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보 고 자 : "
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
         Left            =   9750
         TabIndex        =   9
         Tag             =   "40111"
         Top             =   195
         Width           =   1020
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FEDECD&
         Caption         =   "              /"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7575
         TabIndex        =   15
         Top             =   450
         Width           =   2025
      End
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1320
      Left            =   4200
      TabIndex        =   0
      Top             =   3585
      Visible         =   0   'False
      Width           =   6585
      Begin MSComctlLib.ProgressBar barStatus 
         Height          =   195
         Left            =   1410
         TabIndex        =   1
         Top             =   390
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   45
         Top             =   45
         Width           =   6465
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   495
         Picture         =   "Lis405.frx":0000
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Label1"
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   390
         TabIndex        =   2
         Top             =   825
         Width           =   6045
      End
   End
   Begin VB.PictureBox picOrder 
      Align           =   3  '왼쪽 맞춤
      BackColor       =   &H00E0E0E0&
      Height          =   8130
      Left            =   3075
      ScaleHeight     =   8070
      ScaleWidth      =   14715
      TabIndex        =   23
      Top             =   900
      Width           =   14775
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   7665
         Left            =   90
         TabIndex        =   24
         Top             =   390
         Width           =   14520
         _Version        =   196608
         _ExtentX        =   25612
         _ExtentY        =   13520
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
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
         GridShowVert    =   0   'False
         MaxCols         =   18
         MaxRows         =   30
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   16252927
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "Lis405.frx":0442
         TextTip         =   4
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   315
         Left            =   1260
         TabIndex        =   25
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "진료과"
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   315
         Left            =   2325
         TabIndex        =   26
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "검체"
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   3735
         TabIndex        =   27
         Top             =   120
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "검사명"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Left            =   6075
         TabIndex        =   28
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "결과"
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   315
         Left            =   7470
         TabIndex        =   29
         Top             =   120
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "단위"
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   315
         Left            =   8490
         TabIndex        =   30
         Top             =   135
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "HL"
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   315
         Left            =   9105
         TabIndex        =   31
         Top             =   120
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "기준치"
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   315
         Left            =   10860
         TabIndex        =   32
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "소  견"
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "처방일"
      End
   End
   Begin DRcontrol1.DrFrame fraTextResult 
      Height          =   8430
      Left            =   2610
      TabIndex        =   33
      Top             =   450
      Visible         =   0   'False
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   14870
      BorderStyle     =   0   'False
      Appearance      =   0
      Title           =   ""
      DelLine         =   0
      BackColor       =   16707582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FEDECD&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9255
         Style           =   1  '그래픽
         TabIndex        =   48
         Top             =   135
         Width           =   300
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  '평면
         BackColor       =   &H00F5FFF4&
         ForeColor       =   &H80000008&
         Height          =   7590
         Left            =   165
         ScaleHeight     =   7560
         ScaleWidth      =   9375
         TabIndex        =   34
         Top             =   630
         Width           =   9405
         Begin RichTextLib.RichTextBox txtRstCmt1 
            Height          =   7320
            Left            =   90
            TabIndex        =   35
            Top             =   75
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   12912
            _Version        =   393217
            BackColor       =   16121844
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"Lis405.frx":0F72
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label lblRstCmt1 
         BackStyle       =   0  '투명
         Caption         =   "Text Result "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   36
         Tag             =   "40204"
         Top             =   195
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frm405LegacyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
''% 폼단위 전역변수 선언
'
'Private MyPatient As New clsPatient   '환자 클래스
'Private MySql As New clsSqlStatements   'Sql문 클래스
'Private ClearFg As Boolean
'Private OrderFg As Boolean
'Private ResultFg As Boolean
'Private MsgFg As Boolean
'Private OldRow As Long
'Private OldBackColor As Long
'Private TopLeftShow1 As Boolean
'Private TopLeftShow2 As Boolean
'Private TopLeftShow3 As Boolean
'
'
''% 환자리스트 Display 여부
'Private Sub chkPtList_Click()
'    If chkPtList.Value = 1 And picPtList.Visible = False Then
'        picPtList.Visible = True
'        picPtList.Width = 3200
'        'txtSearchKey.Text = ""
'        'lstPtList.Clear
'        txtSearchKey.SetFocus
'    ElseIf chkPtList.Value = 0 And picPtList.Visible = True Then
'        picPtList.Visible = False
'    End If
'End Sub
'
'Private Sub cmdClose_Click()
'    fraTextResult.Visible = False
'End Sub
'
''%종료
'Private Sub cmdExit_Click()
'   Unload Me
'   Set frm404AllResult = Nothing
'End Sub
'
'Private Sub cmdRefresh_Click()
'   '% 처방조회
'   OldRow = 0
'   Call dtpToDate_KeyDown(vbKeyReturn, 0)
'End Sub
'
''% 레포트 출력
'Private Sub cmdReport_Click()
'
'   With tblOrdSheet
'      .ReDraw = False
'
'      .DisplayColHeaders = True
'      .Row = 0
'      .RowHeight(0) = 20
'      .PrintAbortMsg = "출력중... 취소하려면 Cancel 버튼을 누르세요"
'      .PrintJobName = "Result Print"
'      .PrintHeader = "/l ♧  환자별 검사결과/n/n" & _
'                     "/l   환 자 : " & txtPtId.Text & Space(3) & lblPtNm.Caption & Space(3) & lblSex.Caption & " / " & lblAge.Caption & " " & lblAgeDiv.Caption & "/n" & _
'                     "/l   기 간 : " & Format(dtpFromDate.Value, CS_DateFormat) & "  ~  " & Format(dtpToDate.Value, CS_DateFormat) & "/n/n"
'      .PrintFooter = "/cPage /p"
'      .PrintBorder = True
'      .PrintColor = False
'      .PrintGrid = False    'True
'      .PrintMarginTop = 100
'      .PrintMarginBottom = 100
'      .PrintMarginLeft = 2000
'      .PrintMarginRight = 100
'      .Row = 0: .Row2 = .MaxRows
'      .Col = 0: .Col2 = 8
'      .BlockMode = True
'      .PrintType = PrintTypeCellRange
'      .PrintRowHeaders = True
'      .PrintColHeaders = True
'      .PrintShadows = False
'      .PrintUseDataMax = True
'      ' Perform the printing action
'      .Action = ActionPrint
'
'      .BlockMode = False
'      .DisplayColHeaders = False
'      .ReDraw = True
'   End With
'
'End Sub
'
''% 조회기간 입력 (From Date)
'Private Sub dtpFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyReturn Then dtpToDate.SetFocus
'End Sub
'
''% 조회기간 입력 (To Date)
'Private Sub dtpToDate_KeyDown(KeyCode As Integer, Shift As Integer)
'
'   If dtpToDate.Value < dtpFromDate.Value Then
'      MsgBox "기간 입력 오류입니다. 날짜를 조정하십시요.."
'      dtpFromDate.SetFocus
'      Exit Sub
'   End If
'
'   '% 처방조회
'   Dim i As Integer
'   Dim ResultExist As Boolean
'
'   cmdRefresh.Enabled = False
'   dtpFromDate.Enabled = False
'   dtpToDate.Enabled = False
'
'   Call FieldClear
'   Call TableClear
'
'   'Status Bar Popup
'   DoEvents
'   lblStatus.Caption = lblPtNm.Caption & " 님의 검사 결과내역을 검색중입니다..."
'   barStatus.Min = 0
'   barStatus.Max = DateDiff("d", dtpFromDate.Value, dtpToDate.Value) + 1
'   barStatus.Value = 0
'   fraStatus.Visible = True
'   fraStatus.ZOrder 0
'   DoEvents
'
'   With tblOrdSheet
'      '.ReDraw = False
'      .MaxRows = 0
'
'      ResultExist = False
'      ResultExist = ResultExist Or DisplayOrders("3")
'
'      '.ReDraw = True
'      .Col = 1: .Row = 1: .Action = ActionActiveCell
'   End With
'   fraStatus.Visible = False
'   cmdRefresh.Enabled = True
'   dtpFromDate.Enabled = True
'   dtpToDate.Enabled = True
'
'   If Not ResultExist Then
'      MsgBox "이 환자는 입력하신 기간동안에 보고된 결과가 없습니다."
'      dtpFromDate.SetFocus
'      Exit Sub
'   End If
'
'   ClearFg = False
'   OrderFg = True
'   ResultFg = False
'   tblOrdSheet.SetFocus
'
'End Sub
'
''% 환자ID, 처방일(채혈일)을 기준으로 처방내역을 검색한다.
'Private Function DisplayOrders(ByVal pTestDiv As String) As Boolean
'
'   Dim i As Integer, j As Integer
'   Dim SqlStmt As String
'   Dim ColCnt As Integer
'   Dim tmpTestNm As String
'   Dim tmpRs As New Recordset
'   Dim SvKeyDt As String, SvDeptNm As String, SvSpcNm As String
'   Dim pWorkArea As String, pAccDt As String, pAccSeq As String
'
'   'barStatus.Value = (pTestDiv + 1) * 30
'   'lblStatus.Caption = lblPtNm.Caption & " 님의 " & Choose(pTestDiv + 1, "일반", "특수", "미생물") & "검사 결과내역을 검색중입니다..."
'
'   Screen.MousePointer = vbArrowHourglass  '13
'
'   SqlStmt = " select substring(a.orddt ,3,2)+'-'+substring(a.orddt,5,2)+'-'+substring(a.orddt,7,2) as OrdDate, a.ordtm, " & _
'                 " a.deptcd, b.dept_nm as DeptNm, a.spccd, a.spcnm, a.testcd, a.testnm, a.testseq, a.rstcd, a.rstunit, a.hidiv, a.refvalfrom, a.refvalto, " & _
'                 " substring(a.vfydt,3,2)+'-'+substring(a.vfydt,5,2)+'-'+substring(a.vfydt,7,2)+' '+substring(a.vfytm,1,2)+':'+substring(a.vfytm,3,2) as VfyDtTm, " & _
'                 " a.orddt, a.vfydt, a.vfyid, c.empnm as VfyNm, a.mfyfg, a.mesg " & _
'                 " from LIS99_DB..h7lab999 a, " & TB_HIS003 & " b, " & TB_LAB015 & " c " & _
'                 " where a.ptid = " & txtPtId.Text & " " & _
'                 " and     a.orddt >= '" & Format(dtpFromDate.Value, CS_DateDbFormat) & "' " & _
'                 " and     a.orddt <= '" & Format(dtpToDate.Value, CS_DateDbFormat) & "' " & _
'                 " and     b.hosp_gb =  '" & HosptGb & "' " & _
'                 " and     b.dept_cd =* a.deptcd " & _
'                 " and     c.empid =* a.vfyid " & _
'                 " order by orddt desc, deptcd, ordtm, spccd, testcd, testseq"
'   'Query
'   ColCnt = tmpRs.OpenCursor(DBConn, SqlStmt)
'
'   SvKeyDt = "": SvSpcNm = ""
'
'   DoEvents
'
'   DisplayOrders = False
'   With tblOrdSheet
'
'      .ReDraw = False
'
'      While (tmpRs.FetchCursor(ColCnt))
'
'         barStatus.Value = DateDiff("d", tmpRs.GetValue("OrdDate"), dtpToDate.Value)
'
'         .MaxRows = .MaxRows + 1
'         .Row = .MaxRows
'            If SvKeyDt <> Trim("" & tmpRs.GetValue("OrdDate")) Then
'               .Col = 1:  .Value = Trim("" & tmpRs.GetValue("OrdDate")):  .FontBold = True: .ForeColor = vbBlack    '-- 처방일
'               .Col = 2:  .Value = Trim("" & tmpRs.GetValue("DeptNm")):  .FontBold = False: .ForeColor = vbBlack     '--진료과
'               .Col = 3:  .Value = Trim("" & tmpRs.GetValue("SpcNm")):  .FontBold = True: .ForeColor = &H7477EF   '&H8000&       ''&HC000&       '-- 검체명
'               SvKeyDt = Trim("" & tmpRs.GetValue("OrdDate"))
'               SvDeptNm = Trim("" & tmpRs.GetValue("DeptNm"))
'               SvSpcNm = Trim("" & tmpRs.GetValue("SpcNm"))
'            Else
'               .Col = 1:  .Value = Trim("" & tmpRs.GetValue("OrdDate")):  .FontBold = True: .Value = "" '.ForeColor = .BackColor     '-- 처방일
'               If SvDeptNm <> Trim("" & tmpRs.GetValue("DeptNm")) Then
'                  .Col = 2:  .Value = Trim("" & tmpRs.GetValue("DeptNm")):  .FontBold = False: .ForeColor = vbBlack    '--진료과
'                  .Col = 3:  .Value = Trim("" & tmpRs.GetValue("SpcNm")):  .FontBold = True: .ForeColor = &H7477EF    '-- 검체명
'                  SvDeptNm = Trim("" & tmpRs.GetValue("DeptNm"))
'                  SvSpcNm = Trim("" & tmpRs.GetValue("SpcNm"))
'               Else
'                  .Col = 2:  .Value = Trim("" & tmpRs.GetValue("DeptNm")):  .FontBold = False: .Value = ""  '.ForeColor = .BackColor        '-- 진료과
'                  If SvSpcNm <> Trim("" & tmpRs.GetValue("SpcNm")) Then
'                     .Col = 3:  .Value = Trim("" & tmpRs.GetValue("SpcNm")):  .FontBold = True: .ForeColor = &H7477EF    '-- 검체명
'                     SvSpcNm = Trim("" & tmpRs.GetValue("SpcNm"))
'                  Else
'                     .Col = 3:  .Value = Trim("" & tmpRs.GetValue("SpcNm")):  .FontBold = True: .Value = "" '.ForeColor = .BackColor     '-- 검체명
'                  End If
'               End If
'            End If
'         .Col = 4:  tmpTestNm = Mid(Trim("" & tmpRs.GetValue("TestNm")), 1, 35)
'                       If tmpRs.GetValue("TestSeq") > "01" Then tmpTestNm = "    " & tmpTestNm
'                      .Value = tmpTestNm & " " & String(35 - Len(tmpTestNm), "."):  .ForeColor = &HB9602F      'vbBlue   '&HE48372     '약간 파란색       '-- 검사명
'         .Col = 5:  .ForeColor = &H404080 '갈색       '-- 결과명(코드일 경우..)
'            '.FontName = "굴림"
'            '.FontBold = True
'            If Trim(tmpRs.GetValue("VfyDt")) = "" Then
'               .Value = "미확": .ForeColor = &HC0C0C0:     .FontBold = False:   '.FontName = "돋움"     '회색
'            Else
'               .TypeHAlign = TypeHAlignCenter
'               .Value = Trim(tmpRs.GetValue("RstCd")) ':  .FontBold = True
'            End If
'         .Col = 6:  .Value = Trim("" & tmpRs.GetValue("RstUnit"))         '-- 결과단위
'         .Col = 7       '-- High / Low
'            .Value = ""
'            If Trim(tmpRs.GetValue("VfyDt")) <> "" Then
'                    If Trim(tmpRs.GetValue("HiDiv")) = "H" Then .Value = "▲": .ForeColor = &H7477EF    '약간 붉은색
'                    If Trim(tmpRs.GetValue("HiDiv")) = "L" Then .Value = "▼": .ForeColor = &HDF6A3E       '   &HE48372     '약간 파란색
'            End If
'         .Col = 8:
'            If Val(tmpRs.GetValue("RefValFrom")) <> 0 Or Val(tmpRs.GetValue("RefValTo")) <> 0 Then
'                .Value = Val(tmpRs.GetValue("RefValFrom")) & " - " & Val(tmpRs.GetValue("RefValTo"))    '참고치
'            End If
'         .Col = 9: .Value = Trim("" & tmpRs.GetValue("Mesg"))            '-- 소견
'         'If tmpRs.GetValue("Mesg") <> "" Then Debug.Print tmpRs.GetValue("Mesg")
'
'         .Col = 10: .Value = Trim("" & tmpRs.GetValue("OrdDt"))            '-- 처방일
'         .Col = 11: .Value = Trim("" & tmpRs.GetValue("VfyDtTm"))         '-- 보고일시
'         .Col = 12: .Value = Trim("" & tmpRs.GetValue("VfyNm"))            '-- 보고자
'         .Col = 13: .Value = Trim("" & tmpRs.GetValue("TestCd"))          '-- 검사코드
'         .Col = 14: .Value = Trim("" & tmpRs.GetValue("SpcCd"))          '-- 검체코드
'         .Col = 15: .Value = Trim("" & tmpRs.GetValue("DeptCd"))             '-- 진료과
'         .Col = 16: .Value = Trim("" & tmpRs.GetValue("OrdDate"))             '-- 진료과
'         .Col = 17: .Value = Trim("" & tmpRs.GetValue("DeptNm"))             '-- 진료과
'         .Col = 18: .Value = Trim("" & tmpRs.GetValue("SpcNm"))             '-- 진료과
'
'        DisplayOrders = True
'
'        DoEvents
'
'      Wend
'
'    .Row = -1: .Col = 5: .Col2 = 6
'    .BlockMode = True
'    .AllowCellOverflow = True
'    .BlockMode = False
'
'      .RowHeight(-1) = 11.5
'      .ReDraw = True
'      barStatus.Value = barStatus.Max
'      fraStatus.Visible = False
'
'      tmpRs.CloseCursor
'
'      barStatus.Value = barStatus.Max
'      If .MaxRows < 32 Then .MaxRows = 32
'
'   End With
'
'   Screen.MousePointer = vbDefault
'   Set tmpRs = Nothing
'
'End Function
'
'
'Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption
'    MsgFg = False
'
'End Sub
'
''% 처방 선택(Click)하면 해당 결과 디스플레이...
'Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
'
'   Dim pWorkArea As String
'   Dim pAccDt As String
'   Dim pAccSeq As String
'   Dim pTestDiv As String
'
'   If Row = 0 Then Exit Sub
'   If OldRow = Row Then Exit Sub
'
'   With tblOrdSheet
'
'      .Row = Row
'      .Col = 3:  If .Value = "" Then Exit Sub
'
'      If OldRow > 0 Then
'         .Row = OldRow
'         .Col = -1
'         .BackColor = OldBackColor
'      End If
'
'      .Row = Row
'      .Col = -1
'      OldRow = Row
'      OldBackColor = .BackColor
'      .BackColor = &HD9ECFF ' &HFCEFE9   ' &HF5FFF4       '연두색
'
'      .Col = 11:  lblVerifyDt.Caption = .Value      '보고일시
'      .Col = 12:  lblVerifierNm.Caption = .Value  '보고자
'
'   End With
'
'End Sub
'
'
''% 폼 로드
'Private Sub Form_Load()
'   Me.Top = (Screen.Height - Me.Height) / 2
'   Me.Left = (Screen.Width - Me.Width) / 2
'   Me.Show
'   chkPtList.Value = 0
'   Call chkPtList_Click
'   OrderFg = False
'   ResultFg = False
'   ClearFg = True
'   OldRow = 0
'   TopLeftShow = False
'   optSort(0).Value = True
'
'   dtpFromDate.Value = DateAdd("d", -90, Now)
'   dtpToDate.Value = Now
'
'   'Set MyPatient.MyOraSE = OraSe
'   Set MyPatient.MyDB = DBConn
'   txtPtId.SetFocus
'End Sub
'
''% 환자리스트에서 환자 선택
'Private Sub lstPtList_Click()
'   Dim tmpStr As String
'
'   tmpStr = lstPtList.List(lstPtList.ListIndex)
'   txtPtId.SetFocus
'   txtPtId = medShift(tmpStr, " ")
'   DoEvents
'   Call txtPtId_KeyPress(vbKeyReturn)
'
'End Sub
'
''% 정렬 기준 선택
'Private Sub optSort_Click(Index As Integer)
'   If txtSearchKey.Text <> "" Then
'      Call txtSearchKey_KeyPress(vbKeyReturn)
'   End If
'End Sub
'
'
'Private Sub tblOrdSheet_DblClick(ByVal Col As Long, ByVal Row As Long)
'    With tblOrdSheet
'        .Row = Row
'        .Col = 9
'        If Trim(.Value) = "" Then Exit Sub
'        txtRstCmt1.Text = .Value
'        fraTextResult.Visible = True
'        fraTextResult.ZOrder 0
'    End With
'End Sub
'
''% 처방테이블 Set Focus
'Private Sub tblOrdSheet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''   If OrderFg Then tblOrdSheet.SetFocus
'End Sub
'
''처방내역 테이블에 ToolTip 보여주기...
'
'Private Sub tblOrdSheet_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
'
'   Dim strTmp As String
'
'   With tblOrdSheet
'      If Not TopLeftShow1 Then
'         .Row = OldTop
'         .Col = 1:  .Value = "" '.ForeColor = .BackColor
'      End If
'      If Not TopLeftShow2 Then
'         .Row = OldTop
'         .Col = 2:  .Value = "" '.ForeColor = .BackColor
'      End If
'      If Not TopLeftShow3 Then
'         .Row = OldTop
'         .Col = 3:  .Value = "" '.ForeColor = .BackColor
'      End If
'
'      .Row = NewTop
'      .Col = 1:
'      If .Value <> "" Then
'         TopLeftShow1 = True
'      Else
'         TopLeftShow1 = False
'         .Col = 16:  strTmp = .Value
'         .Col = 1:   .Value = strTmp
'      End If
'      .Col = 2:
'      If .Value <> "" Then
'         TopLeftShow2 = True
'      Else
'         TopLeftShow2 = False
'         .Col = 17:  strTmp = .Value
'         .Col = 2:   .Value = strTmp
'      End If
'      .Col = 3:
'      If .Value <> "" Then
'         TopLeftShow3 = True
'      Else
'         TopLeftShow3 = False
'         .Col = 18:  strTmp = .Value
'         .Col = 3:   .Value = strTmp
'         .Col = 3:  .ForeColor = &H7477EF   '약간 빨간색
'      End If
''      If .ForeColor <> .BackColor Then
''         TopLeftShow1 = True
''      Else
''         TopLeftShow1 = False
''         .Col = 1:  .ForeColor = vbBlack
''      End If
''      .Col = 2:
''      If .ForeColor <> .BackColor Then
''         TopLeftShow2 = True
''      Else
''         TopLeftShow2 = False
''         .Col = 2:  .ForeColor = vbBlack
''      End If
''      .Col = 3:
''      If .ForeColor <> .BackColor Then
''         TopLeftShow3 = True
''      Else
''         TopLeftShow3 = False
''         .Col = 3:  .ForeColor = &H7477EF   '약간 빨간색
''      End If
'   End With
'
'End Sub
'
''% 환자ID가 변경되면 화면Clear
'Private Sub txtPtId_Change()
'   If Not ClearFg Then
'      Call ClearRtn
'   End If
'End Sub
'
''% 환자 ID
'Private Sub txtPtId_GotFocus()
'   With txtPtId
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
'End Sub
'
''% 환자정보 검색
'Private Sub txtPtId_KeyPress(KeyAscii As Integer)
'   If KeyAscii = vbKeyReturn Then
'        SendKeys "{TAB}"
'   End If
'End Sub
'
'
'
'Private Sub txtPtId_LostFocus()
'
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'    If ActiveControl.Name = cmdExit.Name Then Exit Sub
'    If ActiveControl.Name = chkPtList.Name Then Exit Sub
'    If MsgFg Then Exit Sub
'
'    If txtPtId.Text = "" Then
'        txtPtId.SetFocus
'        Exit Sub
'    End If
'
'      With MyPatient
'         If .PtntQuery(txtPtId.Text) Then
'            lblPtNm.Caption = .PtNm
'            lblSex.Caption = .SexNm
'            lblAge.Caption = .Age
'            lblAgeDiv.Caption = .AgeDiv
'            'lblDeptNm.Caption = .DeptNm
'            'lblBedinDt.Caption = Format(.BedInDt, CS_DateMask)
'            'lblBedoutDt.Caption = Format(.BedOutDt, CS_DateMask)
'            ClearFg = False
'         Else
'            MsgFg = True
'            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
'            txtPtId.SetFocus
'            MsgFg = False
'            Call txtPtId_GotFocus
'            Exit Sub
'         End If
'      End With
'      If ActiveControl.Name <> cmdRefresh.Name Then dtpFromDate.SetFocus
'
'
'End Sub
'
''% 텍스트결과 박스1 더블클릭 - Invisible
'Private Sub txtRstCmt1_DblClick()
'   fraTextResult.Visible = False
'End Sub
'
''% Popup Frame 더블클릭 - Invisible
'Private Sub fraTextResult_DblClick()
'   fraTextResult.Visible = False
'End Sub
'
'
''% 환자 검색 (ID 또는 성명으로...)
'Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii <> vbKeyReturn Or txtSearchKey.Text = "" Then Exit Sub
'   If optSort(0).Value Then
'      Call MyPatient.PatientSearch(lstPtList, txtSearchKey.Text, 1)  '환자ID
'   Else
'      Call MyPatient.PatientSearch(lstPtList, txtSearchKey.Text, 2)  '환자성명
'   End If
'
'End Sub
'
''% Clear 루틴
'Private Sub ClearRtn()
'   lblPtNm.Caption = ""
'   lblSex.Caption = ""
'   lblAge.Caption = ""
'   lblAgeDiv.Caption = ""
'   Call FieldClear
'   Call TableClear
'   ClearFg = True
'   OrderFg = False
'   MsgFg = False
'   OldRow = 0
'End Sub
'
'Private Sub FieldClear()
'
'   'lblDoctNm.Caption = ""
'   'lblCollectorNm.Caption = ""
'   'lblReceiverNm.Caption = ""
'   lblVerifierNm.Caption = ""
'   'lblOrdDt.Caption = ""
'   'lblCollectDt.Caption = ""
'   'lblReceiveDt.Caption = ""
'   lblVerifyDt.Caption = ""
'   'txtRstCmt.Text = ""
'   'txtSamCmt1.Text = ""
'   txtRstCmt1.Text = ""
'   'lblWorkArea.Caption = ""
'   'lblSpecimenNm.Caption = ""
'
'End Sub
'
'Private Sub TableClear()
'   tblOrdSheet.MaxRows = 0
'   tblOrdSheet.MaxRows = 100
'   OldRow = 0
'   TopLeftShow = False
'End Sub
'
'Private Sub HighlightText(ByVal pTextBox As Object, ByVal pText As String, ByVal InitFg As Boolean, Optional COLOR As Long = &H80&)
'   With pTextBox
'      If InitFg Then
'         .SelStart = 0
'         .SelLength = Len(.Text)
'         .SelColor = &H0&
'         '.SelBold = False
'      End If
'
'      Dim Point2 As Long
'      Point2 = .Find(pText, 0, , rtfWholeWord)
'      If Point2 <> -1 Then
'         .SelStart = Point2
'         .SelLength = Len(pText)
'         .SelColor = COLOR         '&HFF8080       '&H8080FF           '&HDF6A3E
'         '.SelBold = True
'      End If
'      .SelLength = 0
'   End With
'End Sub
'
'
'Public Sub Call_ToDate_LostFocus()
'
'   Call dtpToDate_KeyDown(vbKeyReturn, 0)
'
'End Sub
'
'
'Public Sub Call_PtId_KeyPress()
'
'   Call txtPtId_KeyPress(vbKeyReturn)
'
'End Sub
'
'
'
'

