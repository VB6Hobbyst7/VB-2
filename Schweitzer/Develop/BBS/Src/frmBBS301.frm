VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS301 
   BackColor       =   &H00DBE6E6&
   Caption         =   "혈액입고"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmBBS301.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   15240
   WindowState     =   2  '최대화
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
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
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
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
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
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
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
      Caption         =   "입고정보"
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   4065
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   45
      Width           =   10440
      _ExtentX        =   18415
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   8190
      Left            =   90
      TabIndex        =   6
      Top             =   285
      Width           =   3930
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   60
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3750
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
         Caption         =   "Volume"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3360
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
         Caption         =   "Component"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   60
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2970
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
         Caption         =   "Center"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdLocalCd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2580
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   31
         Top             =   5370
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   1110
         Left            =   1290
         TabIndex        =   24
         Top             =   3675
         Width           =   2355
         Begin VB.OptionButton optVo 
            BackColor       =   &H00DBE6E6&
            Caption         =   "250cc"
            Height          =   270
            Index           =   2
            Left            =   30
            TabIndex        =   29
            Top             =   480
            Width           =   825
         End
         Begin VB.OptionButton optVo 
            BackColor       =   &H00DBE6E6&
            Caption         =   "400cc"
            Height          =   270
            Index           =   1
            Left            =   825
            TabIndex        =   28
            Top             =   195
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optVo 
            BackColor       =   &H00DBE6E6&
            Caption         =   "320cc"
            Height          =   270
            Index           =   0
            Left            =   30
            TabIndex        =   27
            Top             =   195
            Width           =   795
         End
         Begin VB.TextBox txtVolumn 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            Height          =   285
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   765
            Width           =   870
         End
         Begin VB.OptionButton optVo 
            BackColor       =   &H00DBE6E6&
            Caption         =   "기타"
            Height          =   270
            Index           =   3
            Left            =   825
            TabIndex        =   25
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "cc"
            Height          =   180
            Left            =   990
            TabIndex        =   30
            Top             =   840
            Width           =   210
         End
      End
      Begin VB.TextBox txtLocalCd 
         Appearance      =   0  '평면
         Height          =   360
         Left            =   1275
         TabIndex        =   22
         Top             =   5370
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.ComboBox cboCenter 
         Height          =   300
         Left            =   1290
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   2970
         Width           =   2415
      End
      Begin VB.ComboBox cboCompo 
         Height          =   300
         ItemData        =   "frmBBS301.frx":144A
         Left            =   1290
         List            =   "frmBBS301.frx":144C
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CheckBox chkLocal 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Local 병원입고여부"
         Height          =   315
         Left            =   90
         TabIndex        =   19
         Top             =   5010
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DBE6E6&
         Height          =   2055
         Left            =   180
         TabIndex        =   7
         Top             =   195
         Width           =   3600
         Begin VB.PictureBox Picture3 
            Height          =   1315
            Left            =   120
            ScaleHeight     =   1260
            ScaleWidth      =   1215
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   615
            Width           =   1275
            Begin VB.OptionButton optABO 
               BackColor       =   &H00DBE6E6&
               Caption         =   "A"
               Height          =   315
               Index           =   0
               Left            =   0
               Style           =   1  '그래픽
               TabIndex        =   17
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton optABO 
               BackColor       =   &H00DBE6E6&
               Caption         =   "B"
               Height          =   315
               Index           =   1
               Left            =   0
               Style           =   1  '그래픽
               TabIndex        =   16
               Top             =   320
               Width           =   1215
            End
            Begin VB.OptionButton optABO 
               BackColor       =   &H00DBE6E6&
               Caption         =   "O"
               Height          =   315
               Index           =   2
               Left            =   0
               Style           =   1  '그래픽
               TabIndex        =   15
               Top             =   630
               Width           =   1215
            End
            Begin VB.OptionButton optABO 
               BackColor       =   &H00DBE6E6&
               Caption         =   "AB"
               Height          =   315
               Index           =   3
               Left            =   0
               Style           =   1  '그래픽
               TabIndex        =   14
               Top             =   950
               Width           =   1215
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   435
            Left            =   1400
            ScaleHeight     =   375
            ScaleWidth      =   2070
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   180
            Width           =   2130
            Begin VB.OptionButton optRh 
               BackColor       =   &H00DBE6E6&
               Caption         =   "-"
               Height          =   375
               Index           =   1
               Left            =   1050
               Style           =   1  '그래픽
               TabIndex        =   12
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton optRh 
               BackColor       =   &H00DBE6E6&
               Caption         =   "+"
               Height          =   375
               Index           =   0
               Left            =   0
               Style           =   1  '그래픽
               TabIndex        =   11
               Top             =   0
               Width           =   1050
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00C0FFFF&
            Height          =   1315
            Left            =   1400
            ScaleHeight     =   1260
            ScaleWidth      =   2070
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   615
            Width           =   2130
            Begin VB.Label lblABO 
               Alignment       =   2  '가운데 맞춤
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "AB+"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   27.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   555
               Left            =   345
               TabIndex        =   9
               Top             =   330
               Width           =   1425
            End
         End
         Begin MedControls1.LisLabel LisLabel1 
            Height          =   405
            Left            =   120
            TabIndex        =   18
            Top             =   180
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   714
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
      End
      Begin MedControls1.LisLabel lblLocalNm 
         Height          =   360
         Left            =   1275
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   5760
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel Label3 
         Height          =   360
         Left            =   45
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5760
         Visible         =   0   'False
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
         Caption         =   "병  원  명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel Label13 
         Height          =   360
         Left            =   45
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   5370
         Visible         =   0   'False
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
         Caption         =   "병원 코드"
         Appearance      =   0
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   3650
         Y1              =   4935
         Y2              =   4935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   30
         X2              =   3650
         Y1              =   4950
         Y2              =   4950
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Height          =   8175
      Left            =   4050
      TabIndex        =   37
      Top             =   285
      Width           =   10410
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   2055
         Left            =   120
         TabIndex        =   38
         Top             =   165
         Width           =   4905
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   0
            Left            =   90
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1350
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
            Caption         =   "폐기일자"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   1
            Left            =   90
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   960
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
            Caption         =   "채혈일자"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   2
            Left            =   90
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   570
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
         Begin VB.TextBox txtBldNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            Height          =   360
            Left            =   1290
            MaxLength       =   12
            TabIndex        =   41
            Top             =   570
            Width           =   3285
         End
         Begin VB.TextBox txtAvailable 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            Height          =   360
            Left            =   3825
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1335
            Width           =   750
         End
         Begin VB.CheckBox chkBar 
            BackColor       =   &H00DBE6E6&
            Caption         =   "바코드입력"
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   240
            Width           =   1755
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
            Height          =   375
            Left            =   1290
            TabIndex        =   42
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   59899907
            CurrentDate     =   36799
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
            Height          =   375
            Left            =   1305
            TabIndex        =   43
            Top             =   1350
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   59899907
            CurrentDate     =   36799
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   4
            Left            =   2625
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1335
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
            Caption         =   "보관일수"
            Appearance      =   0
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "일"
            Height          =   255
            Index           =   0
            Left            =   4620
            TabIndex        =   44
            Top             =   1485
            Width           =   180
         End
      End
      Begin FPSpread.vaSpread tblEntList 
         Height          =   5730
         Left            =   135
         TabIndex        =   45
         Top             =   2385
         Width           =   10065
         _Version        =   196608
         _ExtentX        =   17754
         _ExtentY        =   10107
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         GridShowVert    =   0   'False
         MaxCols         =   8
         MaxRows         =   20
         OperationMode   =   1
         ScrollBars      =   1
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS301.frx":144E
         TextTip         =   4
         ScrollBarTrack  =   2
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   270
         Index           =   0
         Left            =   5205
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1170
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
         Left            =   7515
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1200
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
         Left            =   6345
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1170
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
         Left            =   7605
         TabIndex        =   53
         Top             =   1830
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "총 입고 건수 "
         Height          =   180
         Left            =   6360
         TabIndex        =   52
         Top             =   1830
         Width           =   1080
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
         Left            =   5730
         TabIndex        =   51
         Tag             =   "103"
         Top             =   1200
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
         Left            =   8115
         TabIndex        =   50
         Tag             =   "103"
         Top             =   1215
         Width           =   525
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
         Left            =   6900
         TabIndex        =   49
         Tag             =   "103"
         Top             =   1200
         Width           =   525
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00DBF2FD&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   450
         Left            =   5205
         Shape           =   4  '둥근 사각형
         Top             =   1710
         Width           =   4035
      End
   End
End
Attribute VB_Name = "frmBBS301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'혈액입고
'Coding By Legends
Private WithEvents objListPop   As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1

Private objMySQL                As clsBBSSQLStatement

Private CurrentCnt As Long
Private CurrentRow As Long
Private CurrentCol As Long

Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1
'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu

Private Sub cboCompo_Click()
    Dim code As String
    Dim keepday As String
    
    code = medGetP(cboCompo.Text, 1, " ")
    keepday = medGetP(Get_CompNm(code), 2, COL_DIV)
    
    txtAvailable = keepday
    dtpExpDt = dtpColDt + Val(txtAvailable)
    
End Sub

Private Sub chkLocal_Click()
'
    If chkLocal.value = 1 Then
        txtLocalCd.Enabled = True
        cmdLocalCd.Enabled = True
        lblLocalNm.Enabled = True
    Else
        txtLocalCd.Enabled = False
        cmdLocalCd.Enabled = False
        lblLocalNm.Enabled = False
    End If
End Sub

Private Sub cmdClear_Click()
    Call FormInitialize
    Call FormClear
    txtBldNo.SetFocus
    
End Sub

Private Sub cmdClearAll_Click()
'
'    Call FormInitialize
    With tblEntList
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .MaxCols = 8
    End With
    lblEntCnt = ""
    
    CurrentCnt = 0: CurrentRow = 0: CurrentCol = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS301 = Nothing
End Sub

Private Sub cmdLocalCd_Click()
'
    Dim SSQL    As String
    
    SSQL = "select cdval1,field1 from " & T_COM003 & " where " & DBW("cdindex=", BC2_LOCAL)
    Set objListPop = New clsPopUpList
    With objListPop
        txtLocalCd.Text = "": lblLocalNm.Caption = ""
'        .BackColor = Me.BackColor
        .FormCaption = "Local 병원조회": .ColumnHeaderText = "코드;코드명"
'        .Width = .Width + 300: .ColSize(0) = 1000
        Call .LoadPopUp(SSQL) ', 2350, 7650)
        If .SelectedString <> "" Then
            txtLocalCd = medGetP(.SelectedString, 1, ";")
            lblLocalNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    
    Set objListPop = Nothing
    
End Sub
Private Function Save_Check() As Boolean
    '정보 누락 체크
    If optABO(0).value = False And optABO(1).value = False And _
        optABO(2).value = False And optABO(3).value = False Or _
        (optRh(0).value = False And optRh(1).value = False) Then
        
        MsgBox "혈액형을 선택하세요", vbInformation, "정보확인"
        Exit Function
    End If
    
    If cboCenter.Text = "" Then
        MsgBox "Center을 선택하세요.", vbInformation, "정보확인"
        cboCenter.SetFocus
        Exit Function
    End If
    
    If cboCompo.Text = "" Then
        MsgBox "Component을 선택하세요.", vbInformation, "정보확인"
        cboCompo.SetFocus
        Exit Function
    End If
    
    If txtVolumn.Text = "" Then
        MsgBox "Volumn을 입력하세요.", vbInformation, "정보확인"
        txtVolumn.SetFocus
        Exit Function
    End If
    If Val(lblEntCnt.Caption) <= 0 Then Exit Function
    Save_Check = True
    
End Function
Private Sub cmdSave_Click()
    
    Dim lngFirstCol  As Long
    Dim lngRowcnt    As Long
    Dim strABO       As String
    Dim strRh        As String
    Dim lngPos       As Long
    Dim strBldNumber As String  'Ex:20-98-042006
    Dim strBldSrc    As String  '혈액번호(앞자리 두개) 20
    Dim strBldYY     As String  '혈액번호(년도) 98
    Dim lngBldNo     As Long    '혈액번호(일련번호) 042006
    Dim strColDt     As String
    Dim StrExpDt     As String
    Dim strCompocd   As String
    Dim lngAvailable As Long
    Dim strCenterCd  As String
    Dim lngLoopCnt   As Long
    
    Dim SSQL         As String
    
    Dim i As Long
    Dim j As Long
    '--------
    '입력체크
    '--------
    If Save_Check = False Then Exit Sub
    
    lngPos = Len(lblABO.Caption)
    strABO = Mid(lblABO.Caption, 1, lngPos - 1)
    strRh = Mid(lblABO.Caption, lngPos)
    strCompocd = Mid(cboCompo.Text, 1, 2)
    strCenterCd = Mid(cboCenter.Text, 1, 2)
    lngFirstCol = Val(lblEntCnt.Caption) \ 20
    lngLoopCnt = lngFirstCol + 1
    
    Set objMySQL = New clsBBSSQLStatement
    
On Error GoTo Blood_Enter_Error
    DBConn.BeginTrans
    
    With tblEntList
        
        For i = 0 To lngFirstCol
            If lngLoopCnt = i + 1 Then
                For j = 1 To Val(lblEntCnt.Caption) Mod 20
                     .Row = j
                     .Col = i * 4 + 2: strBldNumber = .Text
                     .Col = i * 4 + 3: strColDt = Format(.Text, PRESENTDATE_FORMAT)
                     .Col = i * 4 + 4: StrExpDt = Format(.Text, PRESENTDATE_FORMAT)
                                 
                     strBldSrc = Mid(strBldNumber, 1, 2)
                     strBldYY = Mid(strBldNumber, 4, 2)
                     lngBldNo = Val(Mid(strBldNumber, 7))
                     lngAvailable = DateDiff("d", Format(strColDt, "####-##-##"), Format(StrExpDt, "####-##-##"))
                     
                     SSQL = objMySQL.SetBldStorage(False, strBldSrc, strBldYY, lngBldNo, strCompocd, _
                                                         Val(txtVolumn.Text), strABO, strRh, "", "0", "0", strColDt, _
                                                         "", "", lngAvailable, StrExpDt, "", Format(GetSystemDate, PRESENTDATE_FORMAT), _
                                                         Format(GetSystemDate, PRESENTTIME_FORMAT), Trim(ObjMyUser.EmpId), strCenterCd, "0")
                     DBConn.Execute SSQL
                Next j
            Else
                For j = 1 To 20
                     .Row = j
                     .Col = i * 4 + 2: strBldNumber = .Text
                     .Col = i * 4 + 3: strColDt = Format(.Text, PRESENTDATE_FORMAT)
                     .Col = i * 4 + 4: StrExpDt = Format(.Text, PRESENTDATE_FORMAT)
                                 
                     strBldSrc = Mid(strBldNumber, 1, 2)
                     strBldYY = Mid(strBldNumber, 4, 2)
                     lngBldNo = Val(Mid(strBldNumber, 7))
                     lngAvailable = DateDiff("d", Format(strColDt, "####-##-##"), Format(StrExpDt, "####-##-##"))
                     
                     SSQL = objMySQL.SetBldStorage(False, strBldSrc, strBldYY, lngBldNo, strCompocd, _
                                                         Val(txtVolumn.Text), strABO, strRh, "", "0", "0", strColDt, _
                                                         "", "", lngAvailable, StrExpDt, "", Format(GetSystemDate, PRESENTDATE_FORMAT), _
                                                         Format(GetSystemDate, PRESENTTIME_FORMAT), Trim(ObjMyUser.EmpId), strCenterCd, "0")
                     DBConn.Execute SSQL
                Next j
            End If
        Next i
    End With
        
    DBConn.CommitTrans
    Call FormInitialize
    txtBldNo.SetFocus
    CurrentCol = 0: CurrentRow = 0: CurrentCnt = 0
    Set objMySQL = Nothing
    Exit Sub
    
    
Blood_Enter_Error:
    DBConn.RollbackTrans
    MsgBox "정상적으로 처리되지 않았습니다.", vbInformation, "정보확인"
    Set objMySQL = Nothing
End Sub

Private Sub dtpColdt_Change()
    dtpExpDt = dtpColDt + Val(txtAvailable)
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Call FormInitialize
    Call FormClear
    CurrentCnt = 0: CurrentRow = 0: CurrentCol = 0
End Sub
Private Sub FormClear()
    Dim ii As Integer
    
    For ii = 0 To 3
        optABO(ii).value = False
    Next

    For ii = 0 To 1
        optRh(ii).value = False
    Next
    
    lblABO.Caption = ""
    optVo(0).value = True

End Sub


Private Sub FormInitialize()
'폼 초기화

    Dim objcom003 As clsCom003
    Dim objSQL    As New clsGetSqlStatement
    Dim i As Long
    
    
    
'    objSql.setDbConn DBConn
    Call objSQL.Compolist(cboCompo)
    Set objSQL = Nothing
    
    
    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter)
    Set objcom003 = Nothing

'    For i = 0 To 3
'        optABO(i).value = False
'    Next
'
'    For i = 0 To 1
'        optRh(i).value = False
'    Next
'
'    lblABO.Caption = ""

    cboCenter.ListIndex = medComboFind(cboCenter, ObjSysInfo.BuildingCd & Space(1) & ObjSysInfo.BuildingNm)
    cboCompo.ListIndex = -1
'    txtVolumn.Text = ""
'    optVo(0).value = True
    txtVolumn.Locked = False
    txtLocalCd.Enabled = False
    cmdLocalCd.Enabled = False
    lblLocalNm.Enabled = False
    
    txtBldNo.Text = ""
    dtpColDt = GetSystemDate
    dtpExpDt = GetSystemDate
    txtAvailable.Text = ""
    
    lblEntCnt.Caption = ""
    
    With tblEntList
        .Col = -1
        .Row = -1
        .Action = ActionClearText
    End With
    
    
End Sub
'
'
'Private Sub mnuDelete_Click()
''라인 삭제
'
'    Dim aryList()         As String
'    Dim lngFirstCol       As Long     '값을 지워야 하는 첫번째 Col
'    Dim lngCut            As Long    'ActivCol을 4로 나눈 몫
'    Dim lngStartPos       As Long    '선택된 셀의 위치
'    Dim strNextRowCol1Val As String
'    Dim strNextRowCol2Val As String
'    Dim strNextRowCol3Val As String
'    Dim strNextRowCol4Val As String
'    Dim i                 As Long
'    Dim j                 As Long
'    Dim k                 As Long
'    Dim lngBlockCnt       As Long    '블럭 수(0 부터 시작 ex: 1 이면 블럭이 두개)
'    Dim lngCurrentBlock   As Long    '현재 선택된 블럭
'    Dim strTemp           As String
'
'    With tblEntList
'        If .Row = 0 Then Exit Sub
'        .Row = .ActiveRow
'
'        If .value = "" Then Exit Sub
'        If .ActiveCol Mod 4 = 0 Then Exit Sub
'
'        lngStartPos = .ActiveRow + 1    '줄을 끌어 올릴 첫번째 라인
'        lngCut = .ActiveCol \ 4
'        lngFirstCol = lngCut * 4 + 2      '삭제되는 블럭의 두번째 컬럼값
'
'        lngBlockCnt = (CurrentCnt - 1) \ 20
'        lngCurrentBlock = (lngCut * 4) / 4
'
'        '선택된 부분의 내용지우기
'        .Col = lngFirstCol: .COL2 = lngFirstCol + 2
'        .Row2 = .ActiveRow
'        .BlockMode = True
'        .Action = ActionClearText
'        .BlockMode = False
'        .Row = ((CurrentCnt - 1) Mod 20) + 1: .Col = lngBlockCnt * 4 + 1: .value = ""
'
''        For i = lngStartPos To .MaxRows
'
'            Call CellMove(lngCurrentBlock, lngStartPos, .MaxRows, lngCurrentBlock, lngStartPos - 1, .MaxRows - 1)
'
''            .Row = i
''            .Col = lngFirstCol + 0: strNextRowCol2Val = .Value
''            .Col = lngFirstCol + 1: strNextRowCol3Val = .Value
''            .Col = lngFirstCol + 2: strNextRowCol4Val = .Value
''
''            .Row = i - 1
''            .Col = lngFirstCol + 0: .Value = strNextRowCol2Val
''            .Col = lngFirstCol + 1: .Value = strNextRowCol3Val
''            .Col = lngFirstCol + 2: .Value = strNextRowCol4Val
''        Next i
'
'        For i = lngCurrentBlock + 1 To lngBlockCnt
'
'            Call CellMove(i, 1, 1, i - 1, .MaxRows, .MaxRows)
'
''            .Row = 1
''            .Col = i * 4 + 2: strNextRowCol2Val = .Value
''            .Col = i * 4 + 3: strNextRowCol3Val = .Value
''            .Col = i * 4 + 4: strNextRowCol4Val = .Value
''
''            .Row = .MaxRows
''            .Col = (i - 1) * 4 + 2: .Value = strNextRowCol2Val
''            .Col = (i - 1) * 4 + 3: .Value = strNextRowCol3Val
''            .Col = (i - 1) * 4 + 4: .Value = strNextRowCol4Val
'
'            Call CellMove(i, 2, .MaxRows, i, 1, .MaxRows - 1)
'
''            .Col = i * 4 + 2: .Col2 = i * 4 + 4
''            .Row = 2: .Row2 = .MaxRows
''            .BlockMode = True
''            strTemp = .ClipValue
''            .BlockMode = False
''
''            .Row = 1: .Row2 = .MaxRows - 1
''            .BlockMode = True
''            .ClipValue = strTemp
''            .BlockMode = False
'        Next i
'
'
'
'        CurrentCnt = CurrentCnt - 1
'        CurrentRow = CurrentRow - 1
'        If CurrentRow < 1 Then
'            CurrentRow = .MaxRows
'            CurrentCol = CurrentCol - 4
'        End If
'
'        lblEntCnt.Caption = Val(lblEntCnt.Caption) - 1
'    End With
'
'End Sub

Private Sub CellMove(ByVal iColF As Long, ByVal iRowF1 As Long, ByVal iRowF2 As Long, _
                     ByVal iColT As Long, ByVal iRowT1 As Long, ByVal iRowT2 As Long)
    
    Dim strTemp As String
    With tblEntList
        .Col = iColF * 4 + 2: .COL2 = iColF * 4 + 4
        .Row = iRowF1: .Row2 = iRowF2
        .BlockMode = True
        strTemp = .ClipValue
        .Action = ActionClearText
        .BlockMode = False
        
        .Col = iColT * 4 + 2: .COL2 = iColT * 4 + 4
        .Row = iRowT1: .Row2 = iRowT2
        .BlockMode = True
        .ClipValue = strTemp
        .BlockMode = False
    End With

End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            Dim aryList()         As String
            Dim lngFirstCol       As Long     '값을 지워야 하는 첫번째 Col
            Dim lngCut            As Long    'ActivCol을 4로 나눈 몫
            Dim lngStartPos       As Long    '선택된 셀의 위치
            Dim strNextRowCol1Val As String
            Dim strNextRowCol2Val As String
            Dim strNextRowCol3Val As String
            Dim strNextRowCol4Val As String
            Dim i                 As Long
            Dim j                 As Long
            Dim k                 As Long
            Dim lngBlockCnt       As Long    '블럭 수(0 부터 시작 ex: 1 이면 블럭이 두개)
            Dim lngCurrentBlock   As Long    '현재 선택된 블럭
            Dim strTemp           As String
            
            With tblEntList
                If .Row = 0 Then Exit Sub
                .Row = .ActiveRow
                 
                If .value = "" Then Exit Sub
                If .ActiveCol Mod 4 = 0 Then Exit Sub
                
                lngStartPos = .ActiveRow + 1    '줄을 끌어 올릴 첫번째 라인
                lngCut = .ActiveCol \ 4
                lngFirstCol = lngCut * 4 + 2      '삭제되는 블럭의 두번째 컬럼값
                
                lngBlockCnt = (CurrentCnt - 1) \ 20
                lngCurrentBlock = (lngCut * 4) / 4
                
                '선택된 부분의 내용지우기
                .Col = lngFirstCol: .COL2 = lngFirstCol + 2
                .Row2 = .ActiveRow
                .BlockMode = True
                .Action = ActionClearText
                .BlockMode = False
                .Row = ((CurrentCnt - 1) Mod 20) + 1: .Col = lngBlockCnt * 4 + 1: .value = ""
                
        '        For i = lngStartPos To .MaxRows
                    
                    Call CellMove(lngCurrentBlock, lngStartPos, .MaxRows, lngCurrentBlock, lngStartPos - 1, .MaxRows - 1)
                    
        '            .Row = i
        '            .Col = lngFirstCol + 0: strNextRowCol2Val = .Value
        '            .Col = lngFirstCol + 1: strNextRowCol3Val = .Value
        '            .Col = lngFirstCol + 2: strNextRowCol4Val = .Value
        '
        '            .Row = i - 1
        '            .Col = lngFirstCol + 0: .Value = strNextRowCol2Val
        '            .Col = lngFirstCol + 1: .Value = strNextRowCol3Val
        '            .Col = lngFirstCol + 2: .Value = strNextRowCol4Val
        '        Next i
        
                For i = lngCurrentBlock + 1 To lngBlockCnt
                    
                    Call CellMove(i, 1, 1, i - 1, .MaxRows, .MaxRows)
                    
        '            .Row = 1
        '            .Col = i * 4 + 2: strNextRowCol2Val = .Value
        '            .Col = i * 4 + 3: strNextRowCol3Val = .Value
        '            .Col = i * 4 + 4: strNextRowCol4Val = .Value
        '
        '            .Row = .MaxRows
        '            .Col = (i - 1) * 4 + 2: .Value = strNextRowCol2Val
        '            .Col = (i - 1) * 4 + 3: .Value = strNextRowCol3Val
        '            .Col = (i - 1) * 4 + 4: .Value = strNextRowCol4Val
                    
                    Call CellMove(i, 2, .MaxRows, i, 1, .MaxRows - 1)
                    
        '            .Col = i * 4 + 2: .Col2 = i * 4 + 4
        '            .Row = 2: .Row2 = .MaxRows
        '            .BlockMode = True
        '            strTemp = .ClipValue
        '            .BlockMode = False
        '
        '            .Row = 1: .Row2 = .MaxRows - 1
        '            .BlockMode = True
        '            .ClipValue = strTemp
        '            .BlockMode = False
                Next i
                
                
                
                CurrentCnt = CurrentCnt - 1
                CurrentRow = CurrentRow - 1
                If CurrentRow < 1 Then
                    CurrentRow = .MaxRows
                    CurrentCol = CurrentCol - 4
                End If
        
                lblEntCnt.Caption = Val(lblEntCnt.Caption) - 1
            End With
    End Select
End Sub

Private Sub optABO_Click(Index As Integer)
    lblABO = optABO(Index).Caption
    
    If optRh(0).value = True Then
        lblABO = lblABO & "+"
    ElseIf optRh(1).value = True Then
        lblABO = lblABO & "-"
    End If
End Sub

Private Sub optRh_Click(Index As Integer)
    Dim i As Long
    
    For i = 0 To 3
        If optABO(i).value = True Then
            lblABO = optABO(i).Caption
            Exit For
        End If
    Next i
    lblABO = lblABO & optRh(Index).Caption
End Sub

Private Sub optVo_Click(Index As Integer)

    Select Case Index
        Case 0: txtVolumn = "320"
        Case 1: txtVolumn = "400"
        Case 2: txtVolumn = "250"
        Case 3: txtVolumn.Text = ""
    End Select
End Sub

Private Sub tblEntList_RightClick(ByVal ClickType As Integer, _
                                  ByVal Col As Long, ByVal Row As Long, _
                                  ByVal MouseX As Long, ByVal MouseY As Long)
    With tblEntList
        .Col = Col
        .Row = Row
        .Action = ActionActiveCell
    End With
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE"
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Delete"
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
End Sub



Private Sub txtBldNo_Change()
    If txtBldNo = "" Then Exit Sub
    If cboCompo.Text = "" Then
        MsgBox "혈액제제를 선택하신후 등록하십시요", vbInformation + vbOKOnly, "혈액입고"
        txtBldNo = ""
        Exit Sub
    End If
    If txtVolumn = "" Then
        txtBldNo = ""
        MsgBox "혈액용량을 선택하신후 등록하십시요", vbInformation + vbOKOnly, "혈액입고"
        Exit Sub
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

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    Dim i As Long: Dim j As Long
    Dim Row As Long
    Dim Col As Long
    Dim Cnt As Long
    
    Dim BldNo As String
    Dim BldYY  As String
    Dim BldSrc As String
    Dim CompoCd As String
    
    Dim strBldno  As String
    
'-
    If Len(txtBldNo) <> 3 Or Len(txtBldNo) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn Then
        If Trim(txtBldNo) = "" Then Exit Sub
        If chkBar.value = 1 Then
            If Len(txtBldNo) < 7 Then
                MsgBox "혈액번호를 확인하세요.", vbInformation + vbOKOnly, "혈액번호오류"
                txtBldNo.SelStart = 0
                txtBldNo.SelLength = Len(txtBldNo)
                txtBldNo.SetFocus
                Exit Sub
            End If
            strBldno = Mid(txtBldNo, 1, 2) & "-" & _
                       Mid(txtBldNo, 3, 2) & "-" & _
                       Format(Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2), "00000#")
        Else
            strBldno = Mid(txtBldNo, 1, 6) & Format(Mid(txtBldNo, 7), "00000#")
        End If
        
        
        
        '이미 입고된 혈액체크
        Dim objSQL As New clsGetSqlStatement
        
        BldSrc = medGetP(strBldno, 1, "-")
        BldYY = medGetP(strBldno, 2, "-")
        BldNo = Format(medGetP(strBldno, 3, "-"), "######")
        CompoCd = Mid(cboCompo.Text, 1, 2)
        
        If objSQL.BloodExistChk(BldSrc, BldYY, BldNo, CompoCd) = True Then
            MsgBox "이미 입고된 혈액입니다.", vbInformation + vbOKOnly, "혈액입고"
            txtBldNo = ""
            Exit Sub
        End If
        Set objSQL = Nothing
        
        
        
        '같은 혈액번호를 입력했을 경우 중복값 체크
        If DupCheck(strBldno) Then
            With txtBldNo
                .SelStart = 0
                .SelLength = Len(.Text)
                Exit Sub
            End With
        End If
        
        
        
        Call GetNextRowCol(Row, Col, Cnt)
        
        Dim strColor As String
        
        Select Case txtVolumn.Text
            Case "320": strColor = &H0&
            Case "400": strColor = &HDF6A3E
            Case "250": strColor = &H7477EF
            Case Else: strColor = &H8000&
        End Select
        
        With tblEntList
            .Row = Row
            .Col = Col:     .value = Cnt: .ForeColor = strColor
        
            .Col = Col + 1: .value = strBldno: .ForeColor = strColor

            .Col = Col + 2: .Text = Format(dtpColDt, "YYYY-MM-DD"): .ForeColor = strColor
            .Col = Col + 3: .Text = Format(dtpExpDt, "YYYY-MM-DD"): .ForeColor = strColor
            lblEntCnt = Cnt
            
            If .LeftCol < (.MaxCols - 7) Then
                .LeftCol = .MaxCols - 7
            End If
            
            .Row = Row
            .Col = Col
            .Action = ActionActiveCell
        End With

        txtBldNo = ""

        KeyAscii = 0
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

Private Function DupCheck(ByVal pBldNo As String) As Boolean
'중복값을 체크한다.

    Dim strClip As String
    
    With tblEntList
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strClip = .ClipValue
        .BlockMode = False
        
        If InStr(strClip, pBldNo) Then
            DupCheck = True
        Else
            DupCheck = False
        End If
    End With
    
End Function

Private Sub GetNextRowCol(Row As Long, Col As Long, Cnt As Long)
    If CurrentCol = 0 Then CurrentCol = 1
    
    With tblEntList
        CurrentCnt = CurrentCnt + 1
        CurrentRow = CurrentRow + 1
        If CurrentRow > .MaxRows Then
            CurrentRow = 1
            CurrentCol = CurrentCol + 4
            
            If CurrentCol > .MaxCols Then
                AddTblColumn
                .LeftCol = CurrentCol - 4
            End If
        End If
    End With
    
    Cnt = CurrentCnt
    Row = CurrentRow
    Col = CurrentCol
End Sub

Private Sub AddTblColumn()
    
    Dim MaxCols         As Long
    Dim Header          As String
    Dim CellType        As Long
    Dim HAlign          As Long
    Dim VAlign          As Long
    Dim tpDateCentury   As Boolean
    Dim tpDateMax       As Variant
    Dim tpDateMin       As Variant
    Dim tpDateSeparator As Variant
    Dim tpDateFormat    As Variant
    
    Dim i As Long
    
    With tblEntList
        MaxCols = .MaxCols
        
        .MaxCols = .MaxCols + 4
        
        For i = 1 To 4
            .ColWidth(MaxCols + i) = .ColWidth(i)
            .Row = 0
            .Col = i: Header = .value
            
            .Col = MaxCols + i: .value = Header
            
            .Row = -1
            .Col = i: CellType = .CellType: HAlign = .TypeHAlign: VAlign = .TypeVAlign
                      tpDateCentury = .TypeDateCentury
                      tpDateMax = .TypeDateMax
                      tpDateMin = .TypeDateMin
                      tpDateSeparator = .TypeDateSeparator
                      tpDateFormat = .TypeDateFormat
            .Col = MaxCols + i: .CellType = CellType: .TypeHAlign = HAlign: .TypeVAlign = VAlign
                      .TypeDateCentury = tpDateCentury
                      .TypeDateMax = tpDateMax
                      .TypeDateMin = tpDateMin
                      .TypeDateSeparator = tpDateSeparator
                      .TypeDateFormat = tpDateFormat
            
        Next i
        
        .Col = MaxCols + 1: .COL2 = MaxCols + 4
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 16  'outline
        .Action = ActionSetCellBorder
        .BlockMode = False
    End With
End Sub

Private Sub txtLocalCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtLocalCd.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLocalCd_LostFocus()
    Dim Rs         As Recordset
    Dim objLocalCd As clsBBSSQLStatement
    
    If Trim(txtLocalCd.Text) = "" Then Exit Sub
    
    Set objLocalCd = New clsBBSSQLStatement
    Set Rs = New Recordset
    Rs.Open objLocalCd.GetLocalHp(Trim(txtLocalCd.Text)), DBConn
    
    If Rs.EOF Then
        MsgBox "존재하지 않는 병원코드입니다.", vbInformation, "정보확인"
        
        With txtLocalCd
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
        Set Rs = Nothing
        Set objLocalCd = Nothing
        Exit Sub
    End If
    
    lblLocalNm.Caption = Rs.Fields("field1").value & ""
    
    Set Rs = Nothing
    Set objLocalCd = Nothing
End Sub
