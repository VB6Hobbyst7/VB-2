VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS301_Barcode 
   BackColor       =   &H00DBE6E6&
   Caption         =   "혈액입고"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
   Begin VB.CheckBox chkBar 
      BackColor       =   &H00800000&
      Caption         =   "바코드입력"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   2700
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   135
      Value           =   1  '확인
      Width           =   1275
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   4050
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   75
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "혈액정보"
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   75
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "입고정보(입력정보)"
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Height          =   8175
      Left            =   4050
      TabIndex        =   11
      Top             =   300
      Width           =   10410
      Begin FPSpread.vaSpread tblEntList 
         Height          =   7185
         Left            =   180
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   810
         Width           =   10050
         _Version        =   196608
         _ExtentX        =   17727
         _ExtentY        =   12674
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridShowVert    =   0   'False
         MaxCols         =   9
         MaxRows         =   100
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS301_Barcode.frx":0000
         TextTip         =   4
         ScrollBarTrack  =   2
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   270
         Index           =   0
         Left            =   915
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   476
         BackColor       =   0
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   255
         Index           =   1
         Left            =   3225
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   450
         BackColor       =   7632879
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   270
         Index           =   2
         Left            =   2055
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   270
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   476
         BackColor       =   14641726
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "400cc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2610
         TabIndex        =   20
         Tag             =   "103"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "250cc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3825
         TabIndex        =   19
         Tag             =   "103"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "320cc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1440
         TabIndex        =   18
         Tag             =   "103"
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "총 입고 건수 "
         Height          =   180
         Left            =   7335
         TabIndex        =   17
         Top             =   345
         Width           =   1080
      End
      Begin VB.Label lblEntCnt 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "30"
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
         Left            =   8460
         TabIndex        =   16
         Top             =   345
         Width           =   210
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00DBF2FD&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   450
         Left            =   6180
         Shape           =   4  '둥근 사각형
         Top             =   225
         Width           =   4035
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10485
      Style           =   1  '그래픽
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "15101"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11805
      Style           =   1  '그래픽
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "124"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13125
      Style           =   1  '그래픽
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "128"
      Top             =   8565
      Width           =   1320
   End
   Begin VB.CommandButton cmdClearAll 
      BackColor       =   &H00FEDECD&
      Caption         =   " Clear Table"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9165
      Style           =   1  '그래픽
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8565
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   315
      Left            =   75
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1875
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "입고정보(입력정보확인)"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1500
      Left            =   75
      TabIndex        =   25
      Top             =   300
      Width           =   3945
      Begin VB.TextBox txtBldNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1455
         TabIndex        =   0
         Text            =   "10-04-123456"
         Top             =   165
         Width           =   2340
      End
      Begin VB.TextBox txtCompo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1455
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   615
         Width           =   2340
      End
      Begin VB.TextBox txtABO 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1455
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1050
         Width           =   2340
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   180
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "혈액형"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   180
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   615
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "제제"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   180
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "혈액번호"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   6405
      Left            =   75
      TabIndex        =   21
      Top             =   2085
      Width           =   3945
      Begin VB.Frame fraVol 
         BackColor       =   &H00DBE6E6&
         Height          =   1110
         Left            =   1560
         TabIndex        =   40
         Top             =   1395
         Width           =   2295
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "기타"
            Height          =   270
            Index           =   3
            Left            =   825
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   480
            Width           =   825
         End
         Begin VB.TextBox txtVolume 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            Height          =   285
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   765
            Width           =   870
         End
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "320cc"
            Height          =   270
            Index           =   0
            Left            =   30
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   195
            Width           =   795
         End
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "400cc"
            Height          =   270
            Index           =   1
            Left            =   825
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   195
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "250cc"
            Height          =   270
            Index           =   2
            Left            =   30
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "cc"
            Height          =   180
            Left            =   990
            TabIndex        =   46
            Top             =   840
            Width           =   210
         End
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H80000005&
         Caption         =   "화면지움(&R)"
         CausesValidation=   0   'False
         Height          =   1635
         Left            =   2085
         Style           =   1  '그래픽
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4530
         Width           =   1635
      End
      Begin VB.CommandButton cmdEntry 
         BackColor       =   &H80000005&
         Caption         =   "입력(&E)"
         Height          =   1635
         Left            =   210
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   4530
         Width           =   1635
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0FFFF&
         Height          =   690
         Left            =   1575
         ScaleHeight     =   630
         ScaleWidth      =   2235
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   165
         Width           =   2295
         Begin VB.Label lblRh 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   1260
            TabIndex        =   49
            Top             =   105
            Width           =   270
         End
         Begin VB.Label lblABO 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "AB"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   690
            TabIndex        =   23
            Top             =   105
            Width           =   570
         End
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Left            =   180
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   315
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "혈액형"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   360
         Left            =   180
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   900
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "제제"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   180
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "채혈일자"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpColDt 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   3
         Top             =   2670
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   69337091
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   180
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3495
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "폐기예정일자"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpExpDt 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3510
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   69337091
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   180
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3945
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "보관일수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   180
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1485
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "Volume"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   180
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3075
         Width           =   1245
         _ExtentX        =   2196
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
         Caption         =   "유효일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblAvailableDt 
         Height          =   360
         Left            =   1575
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3075
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   14411494
         ForeColor       =   8421631
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
         Caption         =   "2005-05-30"
         Appearance      =   0
      End
      Begin VB.Label lblAvailable 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1590
         TabIndex        =   47
         Top             =   3960
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "일"
         Height          =   255
         Index           =   0
         Left            =   2460
         TabIndex        =   37
         Top             =   4080
         Width           =   180
      End
      Begin VB.Label lblCompo 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   1575
         TabIndex        =   32
         Top             =   915
         Width           =   2295
      End
   End
   Begin VB.Label lblBldNo 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   990
      TabIndex        =   50
      Top             =   8685
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCompoCd 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3465
      TabIndex        =   48
      Top             =   8685
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmBBS301_Barcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum tcColumn
    tcNo = 1
    tcCompo
    tcABO
    tcRh
    tcVol
    
    tcBldNo '6
    tcCompoCd
    tcColDt
    tcExpDt
End Enum

Private Const RowHeight& = 12
Private Const MENU_DEL& = 1

Private Sub cmdClear_Click()
    txtBldNo.Text = ""
    Call ClearInput
    Call cmdClearAll_Click
End Sub

Private Sub cmdClearAll_Click()
'혈액입고 대기 리스트 지움
    Call medClearTable(tblEntList)
    lblEntCnt.Caption = ""
End Sub

Private Sub cmdEntry_Click()
    Dim strColor As String
    
    If ValidateEntry = False Then Exit Sub
    
    Select Case txtVolume.Text
        Case "320": strColor = &H0&
        Case "400": strColor = &HDF6A3E
        Case "250": strColor = &H7477EF
        Case Else: strColor = &H8000&
    End Select
    
    With tblEntList
        .RowHeight(-1) = RowHeight
        If .DataRowCnt > .MaxRows Then
            .MaxRows = .MaxRows + 1
        End If
        
        .Row = .DataRowCnt + 1
        
        .Col = tcColumn.tcNo: .value = .DataRowCnt + 1
        .Col = tcColumn.tcCompo: .value = lblCompo.Caption
        .Col = tcColumn.tcABO: .value = lblABO.Caption
        .Col = tcColumn.tcRh: .value = lblRh.Caption
        .Col = tcColumn.tcVol: .value = txtVolume.Text
        .Col = tcColumn.tcBldNo: .value = lblBldNo.Caption
        .Col = tcColumn.tcCompoCd: .value = lblCompoCd.Caption
        .Col = tcColumn.tcColDt: .value = Format(dtpColDt.value, "yyyy-mm-dd")
        .Col = tcColumn.tcExpDt: .value = Format(dtpExpDt.value, "yyyy-mm-dd")
                        
        .Col = 1: .COL2 = .MaxCols
        .Row = .Row: .Row2 = .Row
        .BlockMode = True
        .ForeColor = strColor
        .BlockMode = False
        
        lblEntCnt.Caption = .DataRowCnt
        
        .Row = .Row
        .Col = 1
        .Action = ActionActiveCell
    End With
    
    Call cmdReset_Click
    
    txtBldNo.SetFocus
End Sub

Private Function ValidateEntry() As Boolean
    ValidateEntry = False
    
    On Error Resume Next
    
    If txtBldNo.Text = "" Then
        MsgBox "혈액번호를 입력하십시오.", vbExclamation
        txtBldNo.SetFocus
        Exit Function
    End If
    
    If txtCompo.Text = "" Then
        MsgBox "제제/용량을 입력하십시오.", vbExclamation
        txtCompo.SetFocus
        Exit Function
    End If
    
    If txtABO.Text = "" Then
        MsgBox "ABO/Rh를 입력하십시오.", vbExclamation
        txtABO.SetFocus
        Exit Function
    End If
    
    ValidateEntry = True
End Function

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS301_Barcode = Nothing
End Sub

Private Sub cmdReset_Click()
    txtBldNo.Text = ""
    Call ClearInput
    
    On Error Resume Next
    txtBldNo.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim objSQL As clsBBSSQLStatement
    Dim strSQL As String
    Dim i As Long
    
    Dim strBldNo As String
    Dim strBldSrc As String
    Dim strBldYy As String
    Dim lngBldNo As Long
    Dim strCompoCd As String
    Dim lngVolumn As String
    Dim strABO As String
    Dim strRh As String
    Dim strColDt As String
    Dim lngAvailable As Long
    Dim strExpDt As String
    
    If tblEntList.DataRowCnt = 0 Then
        MsgBox "입고할 정보를 입력하십시오.", vbExclamation
        Exit Sub
    End If
        
    Set objSQL = New clsBBSSQLStatement
    
    On Error GoTo ErrTrap
    
    DBConn.BeginTrans
    With tblEntList
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = tcColumn.tcCompoCd: strCompoCd = .value
            .Col = tcColumn.tcABO: strABO = .value
            .Col = tcColumn.tcRh: strRh = .value
            .Col = tcColumn.tcVol: lngVolumn = Val(.value)
            .Col = tcColumn.tcBldNo: strBldNo = .value
                                   strBldSrc = medGetP(strBldNo, 1, "-")
                                   strBldYy = medGetP(strBldNo, 2, "-")
                                   lngBldNo = Val(medGetP(strBldNo, 3, "-"))
            .Col = tcColumn.tcColDt: strColDt = Format(.value, PRESENTDATE_FORMAT)
            .Col = tcColumn.tcExpDt: strExpDt = Format(.value, PRESENTDATE_FORMAT)
            lngAvailable = DateDiff("d", Format(strColDt, "####-##-##"), Format(strExpDt, "####-##-##"))
            
            strSQL = objSQL.SetBldStorage(False, strBldSrc, strBldYy, lngBldNo, strCompoCd, _
                                            lngVolumn, strABO, strRh, "", "0", "0", strColDt, _
                                            "", "", lngAvailable, strExpDt, "", Format(GetSystemDate, PRESENTDATE_FORMAT), _
                                            Format(GetSystemDate, PRESENTTIME_FORMAT), ObjMyUser.EmpId, "10", "0")
        
            DBConn.Execute strSQL
        Next i
    End With
    
    DBConn.CommitTrans
    
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    
    Call cmdClear_Click
    
    Exit Sub
ErrTrap:
    DBConn.RollbackTrans
    MsgBox "처리도중 오류가 발생하였습니다.", vbExclamation
End Sub

Private Sub dtpColdt_Change()
'    dtpExpDt.value = DateAdd("d", Val(lblAvailable.Caption) - 1, Format(dtpColDt.value, "####-##-##"))
'2005/05/30 modify by legends
'폐기일자를 유효일자 다음날로 설정하기 위해서
'    dtpExpDt.value = DateAdd("d", Val(lblAvailable.Caption) - 1, dtpColDt.value)
    lblAvailableDt.Caption = Format(DateAdd("d", Val(lblAvailable.Caption) - 1, dtpColDt.value), "yyyy-MM-dd")
    dtpExpDt.value = DateAdd("d", Val(lblAvailable.Caption), dtpColDt.value)
End Sub

Private Sub dtpColDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    txtBldNo.Text = ""
    Call ClearInput
    dtpColDt = GetSystemDate
    dtpExpDt = GetSystemDate
    Call cmdClearAll_Click
End Sub

Private Sub ClearInput()
'혈액입고 입력 및 확인 창 지움
    lblBldNo.Caption = ""
    txtCompo.Text = ""
    txtABO.Text = ""
    lblABO.Caption = ""
    lblRh.Caption = ""
    lblCompo.Caption = ""
    optVol(3).value = True
'    dtpColDt = GetSystemDate
'    dtpExpDt = GetSystemDate
    lblAvailable.Caption = ""
End Sub

Private Sub lblCompoCd_Change()
    If lblCompoCd.Caption = "" Then Exit Sub
    
    lblAvailable.Caption = medGetP(Get_CompNm(lblCompoCd.Caption), 2, COL_DIV)
'    dtpExpDt.value = dtpColDt.value + Val(lblAvailable.Caption)
'2005/05/30 modify by legends
'폐기일자를 유효일자 다음날로 설정하기 위해서
'    dtpExpDt.value = DateAdd("d", Val(lblAvailable.Caption) - 1, dtpColDt.value)
    lblAvailableDt.Caption = Format(DateAdd("d", Val(lblAvailable.Caption) - 1, dtpColDt.value), "yyyy-MM-dd")
    dtpExpDt.value = DateAdd("d", Val(lblAvailable.Caption), dtpColDt.value)
End Sub

Private Sub optVol_Click(Index As Integer)
    Select Case Index
        Case 0: txtVolume.Text = "320"
        Case 1: txtVolume.Text = "400"
        Case 2: txtVolume.Text = "250"
        Case 3: txtVolume.Text = ""
    End Select
End Sub

Private Sub tblEntList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim objPop As clsPopupMenu
    
    If tblEntList.DataRowCnt = 0 Then Exit Sub
    If Row < 1 Then Exit Sub
    tblEntList.Col = Col
    tblEntList.Row = Row
    If tblEntList.value = "" Then Exit Sub
    
    Set objPop = New clsPopupMenu
    
    With objPop
        .AddMenu MENU_DEL, "삭제"
        
        .PopupMenus Me.hwnd
        
        If .MenuID = MENU_DEL Then
            If MsgBox("입고 대기리스트에서 삭제하시겠습니까?", vbYesNo + vbDefaultButton2) = vbYes Then
                tblEntList.Row = Row
                tblEntList.Action = ActionDeleteRow
                tblEntList.RowHeight(-1) = RowHeight
                
                If tblEntList.DataRowCnt = 0 Then
                    lblEntCnt.Caption = ""
                Else
                    lblEntCnt.Caption = tblEntList.DataRowCnt
                End If
            End If
        End If
    End With
    
    Set objPop = Nothing
End Sub

Private Sub txtABO_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtABO_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtABO.Text = "" Then Exit Sub
        
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtABO_Validate(Cancel As Boolean)
'    If txtBldNo.Text = "" Then
'        MsgBox "혈액번호를 입력하십시오.", vbExclamation
'        Cancel = False
'        SendKeys "+({TAB}{TAB})"
'        GoTo ErrTrap
'    End If
'
'    If txtCompo.Text = "" Then
'        MsgBox "제제코드를 입력하십시오.", vbExclamation
'        Cancel = False
'        SendKeys "+{TAB}"
'        GoTo ErrTrap
'    End If
    
'    If txtABO.Text = "" Then
'        MsgBox "혈액형코드를 입력하십시오.", vbExclamation
'        Cancel = True
'        GoTo ErrTrap
'    End If
    
    If txtABO.Text = "" Then Exit Sub
    
    If CheckExistABO = False Then
        MsgBox "혈액형코드가 존재하지 않습니다.", vbExclamation
        Cancel = True
        GoTo ErrTrap
    End If
    
ErrTrap:
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Function CheckExistABO() As Boolean
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_COM003 & _
             " where " & DBW("cdindex=", BC2_RC_ABO) & _
             " and " & DBW("cdval1=", txtABO.Text)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF Then
        CheckExistABO = False
    Else
        lblABO.Caption = Mid(Rs.Fields("field1").value & "", 1, Len(Rs.Fields("field1").value & "") - 1)
        lblRh.Caption = Mid(Rs.Fields("field1").value & "", Len(Rs.Fields("field1").value & ""))
        CheckExistABO = True
    End If
    
    Set Rs = Nothing
End Function

Private Sub txtBldNo_Change()
    If txtBldNo.Text = "" Then Exit Sub

    If txtCompo.Text <> "" Then
        Call ClearInput
    End If
    
    If chkBar.value = 1 Then Exit Sub
    Dim lngLen As Long
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtBldNo.Text = "" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If Len(txtBldNo.Text) <> 3 Or Len(txtBldNo.Text) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If

    If chkBar.value = 1 Then Exit Sub
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub txtBldNo_Validate(Cancel As Boolean)
    Dim strBldNo As String
    
'    If txtBldNo.Text = "" Then
'        MsgBox "혈액번호를 입력하십시오.", vbExclamation
'        Cancel = True
'        GoTo ErrTrap
'    End If
    
    If txtBldNo.Text = "" Then Exit Sub
    
    If chkBar.value = 1 Then
        If Len(txtBldNo.Text) < 7 Then
            MsgBox "혈액번호를 확인하십시오.", vbExclamation
            Cancel = True
            GoTo ErrTrap
        End If
        
        strBldNo = Mid(txtBldNo.Text, 1, 2) & "-" & _
                   Mid(txtBldNo.Text, 3, 2) & "-" & _
                   Format(Mid(txtBldNo.Text, 5, 6), "00000#")
    Else
        strBldNo = Mid(txtBldNo.Text, 1, 6) & Format(Mid(txtBldNo.Text, 7, 6), "00000#")
    End If
            
    lblBldNo.Caption = strBldNo
            
    Exit Sub
ErrTrap:
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtCompo_Change()
    If lblCompo.Caption <> "" Then
        lblCompo.Caption = ""
        lblCompoCd.Caption = ""
        optVol(3).value = True
        txtVolume.Text = ""
    End If
End Sub

Private Sub txtCompo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtCompo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCompo_Validate(Cancel As Boolean)
    Dim BldSrc As String
    Dim BldYY As String
    Dim BldNo As String
    Dim CompoCd As String
    Dim strBldNo As String
        
'    If txtBldNo.Text = "" Then
'        MsgBox "혈액번호를 입력하십시오.", vbExclamation
'        Cancel = False
'        SendKeys "+{TAB}"
'        GoTo ErrTrap
'    End If
    
'    If txtCompo.Text = "" Then
'        MsgBox "제제코드를 입력하십시오.", vbExclamation
'        Cancel = True
'        GoTo ErrTrap
'    End If

    If txtCompo.Text = "" Then Exit Sub
        
    '마스터에 등록된 제제를 입력했는지 여부 체크
    If CheckExistCompo = False Then
        MsgBox "제제코드가 존재하지 않습니다.", vbExclamation
        SendKeys "{Home}+{End}"
        Cancel = True
        
        Exit Sub
    End If
    
    '혈액용량 구하기
    If CheckVolByCompo = False Then
        MsgBox "혈액용량 코드가 존재하지 않습니다.", vbExclamation
        Cancel = True
        GoTo ErrTrap
    End If
    
    '스프레드에 이미 입력되어 있는지 여부 확인
    '같은 혈액번호와 같은 제제인 경우에 중복
    If CheckDup(lblBldNo.Caption & vbTab & lblCompoCd.Caption) Then
        MsgBox "입고 대기중인 혈액입니다.", vbExclamation
        Cancel = True
        GoTo ErrTrap
    End If
    
    '이미 입고된 혈액체크 BBS401 체크
    Dim objSQL As New clsGetSqlStatement
    
    strBldNo = lblBldNo.Caption
    
    BldSrc = medGetP(strBldNo, 1, "-")
    BldYY = medGetP(strBldNo, 2, "-")
    BldNo = Format(medGetP(strBldNo, 3, "-"), "######") '강제로 0을 채우면 안됨 bldno는 numeric 형이기땜시
    CompoCd = lblCompoCd.Caption
    
    If objSQL.BloodExistChk(BldSrc, BldYY, BldNo, CompoCd) Then
        MsgBox "이미 입고된 혈액입니다.", vbExclamation
        Cancel = True
        GoTo ErrTrap
    End If
    
    Set objSQL = Nothing
    Exit Sub
    
ErrTrap:
    Set objSQL = Nothing
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Function CheckExistCompo() As Boolean
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_COM003 & _
             " where " & DBW("cdindex=", BC2_RC_COMPO) & _
             " and " & DBW("cdval1=", txtCompo.Text)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF Then
        CheckExistCompo = False
    Else
        lblCompoCd.Caption = Rs.Fields("field1").value & ""
        lblCompo.Caption = GetCompoNm(Rs.Fields("field1").value & "")
        CheckExistCompo = True
    End If
    
    Set Rs = Nothing
End Function

Private Function GetCompoNm(ByVal vCompoCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_BBS006 & _
             " where " & DBW("compocd=", vCompoCd) & _
             " and (expdt='' or expdt is null) "
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        GetCompoNm = ""
    Else
        GetCompoNm = Rs.Fields("componm").value & ""
    End If
    
    Set Rs = Nothing
End Function

Private Function CheckVolByCompo() As Boolean
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_COM003 & _
            " where " & DBW("cdindex=", BC2_RC_VOL) & _
            " and " & DBW("cdval1=", Mid(txtCompo.Text, Len(txtCompo.Text), 1))
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF Then
        CheckVolByCompo = False
    Else
        Select Case Rs.Fields("field1").value & ""
            Case "250"
                optVol(2).value = True
            Case "320"
                optVol(0).value = True
            Case "400"
                optVol(1).value = True
            Case Else
                optVol(3).value = True
                txtVolume.Text = ""
        End Select
        
        CheckVolByCompo = True
    End If
    
    Set Rs = Nothing
End Function

Private Function CheckDup(ByVal vCheckVal As String)
    Dim strClip As String
    
    With tblEntList
        .Row = 1: .Row2 = .DataRowCnt
        .Col = tcColumn.tcBldNo: .COL2 = tcColumn.tcCompoCd
        .BlockMode = True
        strClip = .ClipValue
        .BlockMode = False
        
        If InStr(strClip, vCheckVal) Then
            CheckDup = True
        Else
            CheckDup = False
        End If
    End With
End Function

Private Sub txtVolume_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
