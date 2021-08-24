VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm408AccResult 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   15120
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel lblTitle1 
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   609
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
      Caption         =   "접수내역 검색"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblTitle3 
      Height          =   345
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   5985
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   609
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
      Caption         =   "검체접수현황"
      Appearance      =   0
   End
   Begin Crystal.CrystalReport crtReport 
      Left            =   6615
      Top             =   4245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame frmPop 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7230
      Left            =   9690
      TabIndex        =   41
      Top             =   1845
      Visible         =   0   'False
      Width           =   5025
      Begin VB.PictureBox picFootNote 
         Appearance      =   0  '평면
         BackColor       =   &H00EFFEFE&
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   0
         ScaleHeight     =   1155
         ScaleWidth      =   4995
         TabIndex        =   48
         Top             =   5625
         Width           =   5025
         Begin RichTextLib.RichTextBox txtSamCmt 
            Height          =   1155
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   2037
            _Version        =   393217
            BackColor       =   15728382
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frm408.frx":0000
            MouseIcon       =   "frm408.frx":00A5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫      기"
         Height          =   405
         Left            =   15
         TabIndex        =   47
         Top             =   6810
         Width           =   5010
      End
      Begin VB.Frame fraLisResult 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   885
         Left            =   15
         TabIndex        =   42
         Top             =   105
         Width           =   4995
         Begin VB.CheckBox chkSamCmt 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sample Comment"
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3240
            TabIndex        =   43
            Tag             =   "40205"
            Top             =   435
            Value           =   1  '확인
            Width           =   1605
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   255
            Index           =   16
            Left            =   90
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   495
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   450
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
            Caption         =   "검체 "
            Appearance      =   0
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "감수성/소견"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   45
            Top             =   135
            Width           =   1170
         End
         Begin VB.Label lblSpecimenNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Serum"
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
            Left            =   810
            TabIndex        =   44
            Top             =   540
            Width           =   645
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00DF6A3E&
            FillStyle       =   0  '단색
            Height          =   360
            Left            =   60
            Shape           =   4  '둥근 사각형
            Top             =   45
            Width           =   1470
         End
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   4650
         Left            =   0
         TabIndex        =   46
         Top             =   990
         Width           =   5025
         _Version        =   196608
         _ExtentX        =   8864
         _ExtentY        =   8202
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   3
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   11
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frm408.frx":0207
         UnitType        =   0
         UserResize      =   0
         VisibleCols     =   8
         VisibleRows     =   22
         TextTip         =   4
      End
   End
   Begin VB.PictureBox picProcess 
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      Height          =   345
      Left            =   4980
      ScaleHeight     =   345
      ScaleWidth      =   9735
      TabIndex        =   31
      Top             =   1830
      Width           =   9735
      Begin MedControls1.LisLabel lblKindDate 
         Height          =   180
         Left            =   180
         TabIndex        =   32
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
         Caption         =   "보고일"
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   180
         Left            =   7830
         TabIndex        =   33
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   180
         Left            =   1185
         TabIndex        =   34
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
         Caption         =   "검 체"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   180
         Index           =   0
         Left            =   2805
         TabIndex        =   35
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   180
         Left            =   4500
         TabIndex        =   36
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
         Caption         =   "결 과"
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   180
         Left            =   5520
         TabIndex        =   37
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
         Caption         =   "단 위"
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   180
         Left            =   6360
         TabIndex        =   38
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   180
         Left            =   6960
         TabIndex        =   39
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
         Caption         =   "DP"
      End
      Begin MedControls1.LisLabel LisLabel10 
         Height          =   180
         Left            =   8880
         TabIndex        =   40
         Top             =   120
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   318
         BackColor       =   -2147483643
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
         Caption         =   "More"
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   660
      Left            =   4965
      TabIndex        =   26
      Top             =   1185
      Width           =   9750
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00E0E0E0&
         Caption         =   "다음(&N) >>"
         Height          =   465
         Left            =   6225
         Style           =   1  '그래픽
         TabIndex        =   57
         Tag             =   "128"
         Top             =   135
         Width           =   1080
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<< 이전(&P)"
         Height          =   465
         Left            =   5115
         Style           =   1  '그래픽
         TabIndex        =   56
         Tag             =   "128"
         Top             =   135
         Width           =   1080
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "종료(&X)"
         Height          =   465
         Left            =   8160
         Style           =   1  '그래픽
         TabIndex        =   52
         Tag             =   "128"
         Top             =   135
         Width           =   1320
      End
      Begin VB.CheckBox chkToolTip 
         BackColor       =   &H00DBE6E6&
         Caption         =   "상세정보 조회"
         Height          =   180
         Left            =   3495
         TabIndex        =   29
         Top             =   285
         Value           =   1  '확인
         Width           =   1500
      End
      Begin VB.CheckBox chkRefVal 
         BackColor       =   &H00DBE6E6&
         Caption         =   "참고치 조회"
         Height          =   180
         Left            =   2040
         TabIndex        =   28
         Top             =   285
         Width           =   1425
      End
      Begin VB.PictureBox Picture1 
         Height          =   405
         Left            =   435
         ScaleHeight     =   345
         ScaleWidth      =   6000
         TabIndex        =   27
         Top             =   1395
         Width           =   6060
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   225
         Left            =   300
         TabIndex        =   30
         Top             =   270
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   397
         BackColor       =   14641726
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
         Alignment       =   1
         Caption         =   "일반결과"
         Appearance      =   0
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00DF6A3E&
         FillStyle       =   0  '단색
         Height          =   420
         Index           =   0
         Left            =   180
         Shape           =   4  '둥근 사각형
         Top             =   150
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1245
      Left            =   4965
      TabIndex        =   14
      Top             =   -45
      Width           =   9750
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   13
         Left            =   7140
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   165
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
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
         Caption         =   "보 고 자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   14
         Left            =   7140
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   510
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
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
         Caption         =   "보고일시"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   255
         Index           =   15
         Left            =   300
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   825
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
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
         Caption         =   "환자ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   4
         Left            =   2460
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   165
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
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
         Caption         =   "성  명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   2460
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   510
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
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
         Caption         =   "성별/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   300
         Index           =   6
         Left            =   2460
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   855
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
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
         Caption         =   "상 병 명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   4845
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   510
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
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
         Caption         =   "병     실"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   4845
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   165
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
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
         Caption         =   "진 료 과"
         Appearance      =   0
      End
      Begin VB.TextBox txtPtid 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1065
         TabIndex        =   20
         Top             =   510
         Width           =   1215
      End
      Begin MedControls1.LisLabel lblVerifierNm 
         Height          =   315
         Left            =   8130
         TabIndex        =   16
         Top             =   165
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   315
         Left            =   5820
         TabIndex        =   17
         Top             =   510
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   5820
         TabIndex        =   18
         Top             =   165
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblVerifyDt 
         Height          =   315
         Left            =   8130
         TabIndex        =   19
         Top             =   510
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   300
         Left            =   3450
         TabIndex        =   21
         Top             =   165
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDisease 
         Height          =   285
         Left            =   3450
         TabIndex        =   22
         Top             =   870
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   503
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   3450
         TabIndex        =   23
         Top             =   510
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
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
         Alignment       =   1
         Caption         =   "/"
         Appearance      =   0
      End
      Begin VB.Label Label4 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "환자 ID :"
         Height          =   180
         Left            =   255
         TabIndex        =   15
         Top             =   570
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   825
      Left            =   75
      TabIndex        =   10
      Top             =   6255
      Width           =   4890
      Begin VB.CommandButton cmdRefleshCount 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   3660
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   210
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtpDate2 
         Height          =   300
         Left            =   870
         TabIndex        =   12
         Top             =   300
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   529
         _Version        =   393216
         Format          =   7274497
         CurrentDate     =   37480
      End
      Begin VB.Label Label3 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수일 :"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   345
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvwCount 
      Height          =   2010
      Left            =   75
      TabIndex        =   8
      Top             =   7065
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   3545
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사구분"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "접수건수"
         Object.Width           =   3440
      EndProperty
   End
   Begin MedControls1.LisLabel lblTitle2 
      Height          =   345
      Left            =   75
      TabIndex        =   6
      Top             =   2760
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   609
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
      Caption         =   "접수리스트"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   6855
      Left            =   4965
      TabIndex        =   24
      Top             =   2220
      Width           =   9735
      _Version        =   196608
      _ExtentX        =   17171
      _ExtentY        =   12091
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
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
      GrayAreaBackColor=   14411494
      GridColor       =   14013909
      GridShowVert    =   0   'False
      MaxCols         =   45
      OperationMode   =   4
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   16252927
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frm408.frx":1DF8
      TextTip         =   4
   End
   Begin MSComctlLib.ListView lvwAccList 
      Height          =   2880
      Left            =   75
      TabIndex        =   25
      Top             =   3105
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   5080
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "번호"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "접수"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "상태"
         Object.Width           =   1288
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "환자ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "환자명"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "검체명"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "접수시간"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "접수번호"
         Object.Width           =   0
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtRstCmt1 
      Height          =   1455
      Left            =   9345
      TabIndex        =   50
      Top             =   60
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frm408.frx":6503
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   2460
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   4890
      Begin MedControls1.LisLabel lblDate 
         Height          =   285
         Left            =   60
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   615
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
         Caption         =   "접수 일 자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblQKind 
         Height          =   285
         Left            =   60
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
         Caption         =   "Work Area"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestCd 
         Height          =   285
         Left            =   60
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1485
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
         Caption         =   "검 사 명 "
         Appearance      =   0
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Caption         =   "Frame5"
         Height          =   1140
         Left            =   3705
         TabIndex        =   63
         Top             =   615
         Width           =   1155
         Begin VB.OptionButton optStat 
            BackColor       =   &H00DBE6E6&
            Caption         =   "응급"
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
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   66
            Tag             =   "15305"
            Top             =   120
            Width           =   945
         End
         Begin VB.OptionButton optStat 
            BackColor       =   &H00DBE6E6&
            Caption         =   "비응급"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   65
            Tag             =   "15305"
            Top             =   480
            Width           =   945
         End
         Begin VB.OptionButton optStat 
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
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   64
            Tag             =   "15305"
            Top             =   855
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.ComboBox cboSpcCd 
         Height          =   300
         ItemData        =   "frm408.frx":65A0
         Left            =   1140
         List            =   "frm408.frx":65A2
         Style           =   2  '드롭다운 목록
         TabIndex        =   62
         Top             =   1920
         Width           =   2520
      End
      Begin VB.ComboBox cboTestCd 
         Height          =   300
         ItemData        =   "frm408.frx":65A4
         Left            =   1140
         List            =   "frm408.frx":65A6
         Style           =   2  '드롭다운 목록
         TabIndex        =   60
         Top             =   1485
         Width           =   2520
      End
      Begin MSComCtl2.DTPicker dtpTime 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "AMPM h시 n분 s초"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   2475
         TabIndex        =   58
         Top             =   615
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Format          =   7274498
         UpDown          =   -1  'True
         CurrentDate     =   37502
      End
      Begin VB.OptionButton optWorkSheet 
         BackColor       =   &H00DBE6E6&
         Caption         =   "WorkSheet"
         Height          =   180
         Left            =   2475
         TabIndex        =   53
         Top             =   255
         Width           =   1275
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조  회(&Q)"
         Height          =   495
         Left            =   3720
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   1755
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtpDate1 
         Height          =   300
         Left            =   1155
         TabIndex        =   4
         Top             =   615
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Format          =   7274497
         CurrentDate     =   37480
      End
      Begin VB.OptionButton optVfydt 
         BackColor       =   &H00DBE6E6&
         Caption         =   "보고일"
         Height          =   180
         Left            =   1297
         TabIndex        =   3
         Top             =   255
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optAccdt 
         BackColor       =   &H00DBE6E6&
         Caption         =   "접수일"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   255
         Width           =   1020
      End
      Begin VB.CommandButton cmdWSList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3405
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   55
         Top             =   1035
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.ComboBox cboWA 
         Height          =   300
         ItemData        =   "frm408.frx":65A8
         Left            =   1155
         List            =   "frm408.frx":65AA
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   1050
         Width           =   2520
      End
      Begin VB.TextBox txtWorkCd 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         TabIndex        =   59
         Top             =   1050
         Visible         =   0   'False
         Width           =   2235
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   45
         Left            =   30
         TabIndex        =   61
         Top             =   525
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   79
         BackColor       =   9342088
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
         Caption         =   ""
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
         Caption         =   "검체 선택 "
         Appearance      =   0
      End
   End
   Begin RichTextLib.RichTextBox rtfResult 
      Height          =   8070
      Left            =   75
      TabIndex        =   51
      Top             =   45
      Visible         =   0   'False
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   14235
      _Version        =   393217
      BackColor       =   16777207
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9000
      TextRTF         =   $"frm408.frx":65AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstWSCode 
      Appearance      =   0  '평면
      BackColor       =   &H00F4FDFF&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   1215
      TabIndex        =   54
      Top             =   1680
      Visible         =   0   'False
      Width           =   4980
   End
End
Attribute VB_Name = "frm408AccResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSQL As S2LIS_SqlLib.clsLISHospital05
Private MySql As New clsLISSqlReview     'Sql문 클래스
Private objPtInfo As clsPatientInfo
Private objLab032 As clsComcode032
Private objLab301 As clsWSBuild

Public Event LastFormUnload()
Public Event ThisFormUnload()

Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_PRT& = 1

Private StopFg As Boolean
Private OldRow As Long
Private OldBackColor As Long
Private aryMesg() As String

Private Sub cboWA_Click()
    Dim RS As New Recordset
    Dim strWA As String
    
    strWA = Trim(medGetP(cboWA.Text, 2, Space(60)))
    Set RS = objSQL.LoadTestCdNm(strWA)

    'cboTestCd 콤보박스에 검사명, 검사코드 출력
    cboTestCd.Clear
    If Not RS.EOF Then
        cboTestCd.AddItem "모든검사항목" & Space(60) & "All"
        With RS
            Do Until .EOF
                cboTestCd.AddItem .Fields("abbrnm10").Value & "" & Space(60) & .Fields("testcd").Value & ""
                .MoveNext
            Loop
        End With
        cboTestCd.ListIndex = 0
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub chkRefVal_Click()
    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New Recordset
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    Dim I As Long, J As Long
    Dim SqlStmt As String
    
    With tblOrdSheet
        For I = 1 To .MaxRows
            '참고치 검색
            .Row = I
            .Col = 8: If .Value <> CS_QuestionMark Then GoTo RefSkip
            
            .Col = 25:  tmpSex = Trim(.Value)
            .Col = 26:  tmpAgeDay = Trim(.Value)
            .Col = 27:  tmpTestCd = Trim(.Value)
            .Col = 28:  tmpSpcCd = Trim(.Value)
            .Col = 29:  tmpVfyDt = Trim(.Value)
                        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
         
            SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open SqlStmt, DBConn
            If tmpRs1.EOF Then
                '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
                '--> 거의 Both로 등록됨.
                SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = Nothing
                Set tmpRs1 = New Recordset
                tmpRs1.Open SqlStmt, DBConn
            End If
            If tmpRs1.EOF Then
                tmpRefCd = Space(5)
            Else
                tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
                If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
                   tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
            End If
            Set tmpRs1 = Nothing
            For J = I To .MaxRows
                .Row = J
                .Col = 27   '참고치
                If Trim(.Value) = tmpTestCd Then _
                    .Col = 8: .Value = tmpRefCd: .ForeColor = DCM_Green
            Next
         
            DoEvents
        
RefSkip:
        Next
    End With
    Set tmpRs1 = Nothing
End Sub

Private Sub chkSamCmt_Click()
    If chkSamCmt.Value = 1 Then
        picFootNote.Visible = True
        tblResult.Height = tblResult.Height - picFootNote.Height
    ElseIf chkSamCmt.Value = 0 Then
        picFootNote.Visible = False
        tblResult.Height = tblResult.Height + picFootNote.Height
    End If
End Sub

Private Sub chkToolTip_Click()
    If chkToolTip.Value = 1 Then
        tblOrdSheet.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent ThisFormUnload
    
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdNext_Click()
    Dim intIndex As Integer
    
    If lvwAccList.ListItems.Count = 0 Then Exit Sub
    If lvwAccList.SelectedItem Is Nothing Then Exit Sub
    
    intIndex = lvwAccList.SelectedItem.Index + 1
    If intIndex > lvwAccList.ListItems.Count Then Exit Sub
    
    lvwAccList.ListItems(intIndex).Selected = True
    Call lvwAccList_ItemClick(lvwAccList.SelectedItem)
End Sub

Private Sub cmdPre_Click()
    Dim intIndex As Integer
    
    If lvwAccList.ListItems.Count = 0 Then Exit Sub
    If lvwAccList.SelectedItem Is Nothing Then Exit Sub
    
    intIndex = lvwAccList.SelectedItem.Index - 1
    If intIndex = 0 Then Exit Sub
    
    lvwAccList.ListItems(intIndex).Selected = True
    Call lvwAccList_ItemClick(lvwAccList.SelectedItem)
End Sub

Private Sub cmdQuery_Click()
    Dim RS          As New Recordset
    Dim ObjDic      As clsDictionary
    Dim objPro      As jProgressBar.clsProgress
    Dim itmFound    As ListItem
    
    Dim strWorkArea As String               'WorkArea
    Dim strDate     As String               '접수/보고일자
    Dim strStsCd    As String               '검체코드
    Dim strWSCd     As String               'WorkSheet Code
    Dim strTime     As String               '접수시간
    Dim strTestcd   As String               '검사코드
    Dim strSpcCd    As String               '검체코드
    
    
    Dim objQ    As New clsDictionary
    Dim objS    As New clsDictionary
    Dim aryD()  As String
    Dim kk      As Integer
    Dim ll      As Integer
            
    Dim I As Integer
    Dim J As Integer
    
    Call ClearLabel
    Call medClearTable(tblOrdSheet)
    lvwAccList.ListItems.Clear

    '접수일/보고일, WorkArea, WorkSheet Code, 검사코드
    strDate = Format(dtpDate1.Value, "yyyyMMdd")
    strWSCd = Trim(medGetP(lstWSCode.Text, 1, vbTab))
    strTime = Format(dtpTime.Value, "HHmmss")
    strTestcd = Trim(medGetP(cboTestCd.Text, 2, Space(60)))
    strSpcCd = Trim(medGetP(cboSpcCd.Text, 2, Space(60)))
    
    If strSpcCd = "All" Then
        objSQL.SpcCd = ""
    Else
        objSQL.SpcCd = strSpcCd
    End If
    
    If optStat(0).Value = True Then
        objSQL.Stat = "1"
    ElseIf optStat(1).Value = True Then
        objSQL.Stat = "0"
    Else
        objSQL.Stat = ""
    End If
    
    
    '----------------------------------------------------------'
    '   WorkSheet 별조회
    '----------------------------------------------------------'
    If optWorkSheet.Value = True Then
        If strWSCd = "" Then
            MsgBox "WorkSheet 를 선택하신 후 조회하세요", vbInformation + vbOKOnly, "Info"
            Set RS = Nothing
            Exit Sub
        End If
        Set RS = objSQL.GetAccList(strDate, strTime, strWSCd)
        Set ObjDic = New clsDictionary
        
        '조회결과를 Dictionary 에 담기
        I = 1
        J = 1
        With ObjDic
            .FieldInialize "spcyy,spcno, keyseq", _
                             "seq,workarea,accdt,accseq,stscd,ptid,rcvtm,name,field5,workcd"
            .Sort = False
            .DeleteAll
            
            If Not RS.EOF Then
                Do Until RS.EOF
                    If .Exists(RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & "" & COL_DIV & J) Then
                        I = I - 1
                        J = J + 1
                        .AddNew Join(Array(RS.Fields("spcyy").Value & "", RS.Fields("spcno").Value & "", J), COL_DIV), _
                                     Join(Array(I, RS.Fields("workarea").Value & "", RS.Fields("accdt").Value & "", _
                                            RS.Fields("accseq").Value & "", RS.Fields("stscd").Value & "", _
                                            RS.Fields("ptid").Value & "", RS.Fields("rcvtm").Value & "", RS.Fields("name").Value & "", _
                                            RS.Fields("field5").Value & "", RS.Fields("workcd").Value & ""), COL_DIV)
                        
                    Else
                        J = 1
                        .AddNew Join(Array(RS.Fields("spcyy").Value & "", RS.Fields("spcno").Value & "", J), COL_DIV), _
                                     Join(Array(I, RS.Fields("workarea").Value & "", RS.Fields("accdt").Value & "", _
                                            RS.Fields("accseq").Value & "", RS.Fields("stscd").Value & "", _
                                            RS.Fields("ptid").Value & "", RS.Fields("rcvtm").Value & "", RS.Fields("name").Value & "", _
                                            RS.Fields("field5").Value & "", RS.Fields("workcd").Value & ""), COL_DIV)
                    End If
                    
                    I = I + 1
                    RS.MoveNext
                Loop
            End If
        End With
        
        '결과를 ListView에 출력
        I = 1
        lvwAccList.ListItems.Clear
        ObjDic.MoveFirst

        'ProgressBar 처리
        Set objPro = Nothing
        Set objPro = New jProgressBar.clsProgress
        With objPro
            .Container = Me
            .Left = lblTitle2.Left
            .Top = lblTitle2.Top
            .Width = lblTitle2.Width
            .Height = lblTitle2.Height
            .Max = ObjDic.RecordCount
        End With
        ObjDic.MoveFirst
        
        With lvwAccList
            Do Until ObjDic.EOF
                If ObjDic.Fields("workcd") = strWSCd Then
                    Set itmFound = .FindItem(Format(ObjDic.Fields("seq"), "0000"))
                    If itmFound Is Nothing Then
                        Set itmFound = .ListItems.Add(, , Format(ObjDic.Fields("seq"), "0000"))
                        
                        Select Case Trim(ObjDic.Fields("stscd"))
                            Case enStsCd.StsCd_LIS_Order
                                strStsCd = "처방"
                            Case enStsCd.StsCd_LIS_Collection
                                strStsCd = "채혈"
                            Case enStsCd.StsCd_LIS_Accession
                                strStsCd = "접수"
                            Case enStsCd.StsCd_LIS_InProcess
                                strStsCd = "검사중"
                            Case enStsCd.StsCd_LIS_MidRst
                                strStsCd = "중간보고"
                            Case enStsCd.StsCd_LIS_FinRst
                                strStsCd = "결과"
                            Case enStsCd.StsCd_LIS_Modify
                                strStsCd = "수정"
                            Case enStsCd.StsCd_LIS_Cancel
                                strStsCd = "취소"
                        End Select
                        
                        itmFound.SubItems(1) = ObjDic.Fields("accseq")
                        itmFound.SubItems(2) = strStsCd
                        itmFound.SubItems(3) = ObjDic.Fields("ptid")
                        itmFound.SubItems(4) = ObjDic.Fields("name")
                        itmFound.SubItems(5) = ObjDic.Fields("field5")
                        
                        strTime = Trim(ObjDic.Fields("rcvtm"))
                        itmFound.SubItems(6) = Mid(strTime, 1, 2) & "시 " & Mid(strTime, 3, 2) & "분"
                        itmFound.SubItems(7) = ObjDic.Fields("workarea") & "-" & _
                                               Mid$(ObjDic.Fields("accdt"), 3) & "-" & _
                                               ObjDic.Fields("accseq")
                    End If
                    
                End If
                
                I = I + 1
                objPro.Value = I
                ObjDic.MoveNext
            Loop
        End With

    '----------------------------------------------------------'
    '   접수일로 조회시
    '----------------------------------------------------------'
    ElseIf optAccdt = True Then
        strWorkArea = Trim(medGetP(cboWA.Text, 2, Space(60)))
        
        '모든검사 항목으로 조회시
        If strTestcd = "All" Then
            Set RS = objSQL.GetAccInfo01(strDate, strWorkArea, strTime)
            
            'ProgressBar 설정
            Set objPro = Nothing
            Set objPro = New jProgressBar.clsProgress
            With objPro
                .Container = Me
                .Left = lblTitle2.Left
                .Top = lblTitle2.Top
                .Width = lblTitle2.Width
                .Height = lblTitle2.Height
                .Max = RS.RecordCount
            End With
            
            I = 1
            lvwAccList.ListItems.Clear
            With lvwAccList
                Do Until RS.EOF
                    Set itmFound = .FindItem(Format(I, "0000"))
                    If itmFound Is Nothing Then
                        Set itmFound = .ListItems.Add(, , Format(I, "0000"))

                        Select Case Trim(RS.Fields("stscd").Value & "")
                            Case enStsCd.StsCd_LIS_Order
                                strStsCd = "처방"
                            Case enStsCd.StsCd_LIS_Collection
                                strStsCd = "채혈"
                            Case enStsCd.StsCd_LIS_Accession
                                strStsCd = "접수"
                            Case enStsCd.StsCd_LIS_InProcess
                                strStsCd = "검사중"
                            Case enStsCd.StsCd_LIS_MidRst
                                strStsCd = "중간보고"
                            Case enStsCd.StsCd_LIS_FinRst
                                strStsCd = "결과"
                            Case enStsCd.StsCd_LIS_Modify
                                strStsCd = "수정"
                            Case enStsCd.StsCd_LIS_Cancel
                                strStsCd = "취소"
                        End Select

                        itmFound.SubItems(1) = RS.Fields("accseq").Value & ""
                        itmFound.SubItems(2) = strStsCd
                        itmFound.SubItems(3) = RS.Fields("ptid").Value & ""
                        itmFound.SubItems(4) = RS.Fields("name").Value & ""
                        itmFound.SubItems(5) = RS.Fields("field5").Value & ""
                        strTime = Trim(RS.Fields("rcvtm").Value & "")
                        itmFound.SubItems(6) = Mid(strTime, 1, 2) & "시 " & Mid(strTime, 3, 2) & "분"
                        itmFound.SubItems(7) = RS.Fields("workarea").Value & "" & "-" & _
                                               Mid$(RS.Fields("accdt").Value & "", 3) & "-" & _
                                               RS.Fields("accseq").Value & ""
                    End If
    
                    I = I + 1
                    objPro.Value = I
                    RS.MoveNext
                Loop
            End With
            
        '지정된 검사항목으로 조회시
        Else
            Set RS = objSQL.GetAccInfoByTime(strDate, 1, strWorkArea, strTime)
            
            Me.MousePointer = 11
            lvwAccList.ListItems.Clear
            
            objS.Clear:            objS.FieldInialize "spcno", "seq"
            objQ.Clear:            objQ.FieldInialize "testcd", "data"

            objQ.Sort = False: objS.Sort = False
            
            I = 1
            If Not RS.EOF Then
                Set objPro = Nothing
                Set objPro = New jProgressBar.clsProgress
                With objPro
                    .Container = Me
                    .Left = lblTitle2.Left
                    .Top = lblTitle2.Top
                    .Width = lblTitle2.Width
                    .Height = lblTitle2.Height
                    .Max = RS.RecordCount * 2
                End With
                
                Do Until RS.EOF
                    If objS.Exists(RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & "") Then
                    Else
                        objS.AddNew RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & "", I
                        I = I + 1
                    End If
                    ll = ll + 1
                    objPro.Value = ll
                    objPro.Message = " 검체저장 순번을 수집합니다......(" & ll & "/" & objPro.Max & ")"
                    RS.MoveNext
                Loop
                
                RS.MoveFirst
                Do Until RS.EOF
                    If objS.Exists(RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & "") Then
                        objS.KeyChange RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & ""
                        I = objS.Fields("seq")
                    End If

                    If objQ.Exists(RS.Fields("testcd").Value & "") Then
                        objQ.KeyChange RS.Fields("testcd").Value & ""
                        objQ.Fields("data") = objQ.Fields("data") & "★" & I & "§" & RS.Fields("workarea").Value & "" & "§" & _
                                                                   RS.Fields("accdt").Value & "" & "§" & _
                                                                   RS.Fields("accseq").Value & "" & "§" & _
                                                                   RS.Fields("stscd").Value & "" & "§" & _
                                                                   RS.Fields("ptid").Value & "" & "§" & _
                                                                   RS.Fields("name").Value & "" & "§" & _
                                                                   RS.Fields("field5").Value & "" & "§" & _
                                                                   RS.Fields("rcvtm").Value & ""
                    Else
                        objQ.AddNew RS.Fields("testcd").Value & "", I & "§" & RS.Fields("workarea").Value & "" & "§" & _
                                                                   RS.Fields("accdt").Value & "" & "§" & _
                                                                   RS.Fields("accseq").Value & "" & "§" & _
                                                                   RS.Fields("stscd").Value & "" & "§" & _
                                                                   RS.Fields("ptid").Value & "" & "§" & _
                                                                   RS.Fields("name").Value & "" & "§" & _
                                                                   RS.Fields("field5").Value & "" & "§" & _
                                                                   RS.Fields("rcvtm").Value & ""
                    End If
                    ll = ll + 1
                    objPro.Value = ll
                    objPro.Message = " 검사항목을 수집합니다..........(" & ll & "/" & objPro.Max & ")"
                    RS.MoveNext
                Loop
            End If
            
            objS.Sort = True
            objQ.Sort = True
            
            objQ.MoveFirst
            If objQ.Exists(strTestcd) Then
                objQ.KeyChange strTestcd
                aryD() = Split(objQ.Fields("data"), "★")
                Set objPro = Nothing
                Set objPro = New jProgressBar.clsProgress
                With objPro
                    .Container = Me
                    .Left = lblTitle2.Left
                    .Top = lblTitle2.Top
                    .Width = lblTitle2.Width
                    .Height = lblTitle2.Height
                    .Max = UBound(aryD)
                End With
                objPro.Message = "자료를 Display 합니다....."
                For kk = LBound(aryD) To UBound(aryD)
                    With lvwAccList
                        Set itmFound = .FindItem(Format(medGetP(aryD(kk), 1, "§"), "0000"))
                        If itmFound Is Nothing Then
                            Set itmFound = .ListItems.Add(, , Format(medGetP(aryD(kk), 1, "§"), "0000"))

                            Select Case Trim(medGetP(aryD(kk), 5, "§"))
                                Case enStsCd.StsCd_LIS_Order
                                    strStsCd = "처방"
                                Case enStsCd.StsCd_LIS_Collection
                                    strStsCd = "채혈"
                                Case enStsCd.StsCd_LIS_Accession
                                    strStsCd = "접수"
                                Case enStsCd.StsCd_LIS_InProcess
                                    strStsCd = "검사중"
                                Case enStsCd.StsCd_LIS_MidRst
                                    strStsCd = "중간보고"
                                Case enStsCd.StsCd_LIS_FinRst
                                    strStsCd = "결과"
                                Case enStsCd.StsCd_LIS_Modify
                                    strStsCd = "수정"
                                Case enStsCd.StsCd_LIS_Cancel
                                    strStsCd = "취소"
                            End Select

                            itmFound.SubItems(1) = medGetP(aryD(kk), 4, "§")
                            itmFound.SubItems(2) = strStsCd
                            itmFound.SubItems(3) = medGetP(aryD(kk), 6, "§")
                            itmFound.SubItems(4) = medGetP(aryD(kk), 7, "§")
                            itmFound.SubItems(5) = medGetP(aryD(kk), 8, "§")
                            strTime = Trim(medGetP(aryD(kk), 7, "§"))
                            itmFound.SubItems(6) = Mid(strTime, 1, 2) & "시 " & Mid(strTime, 3, 2) & "분"
                            itmFound.SubItems(7) = medGetP(aryD(kk), 2, "§") & "-" & _
                                                   medGetP(aryD(kk), 3, "§") & "-" & _
                                                   medGetP(aryD(kk), 4, "§")
                        End If

                    End With
                    ll = ll + 1
                    objPro.Value = ll

                Next
            End If
            Me.MousePointer = 0
            Set objQ = Nothing
            GoTo Skip
 
            '결과를 ListView에 출력
            I = 1
            lvwAccList.ListItems.Clear
            ObjDic.MoveFirst
    
            'ProgressBar 설정
            Set objPro = Nothing
            Set objPro = New jProgressBar.clsProgress
            With objPro
                .Container = Me
                .Left = lblTitle2.Left
                .Top = lblTitle2.Top
                .Width = lblTitle2.Width
                .Height = lblTitle2.Height
                .Max = ObjDic.RecordCount
            End With
            ObjDic.MoveFirst
    
            I = 1
            With lvwAccList
                Do Until ObjDic.EOF
                    DoEvents
                    If ObjDic.Fields("testcd") = strTestcd Then
                        Set itmFound = .FindItem(Format(ObjDic.Fields("seq"), "0000"))
                        If itmFound Is Nothing Then
                            Set itmFound = .ListItems.Add(, , Format(ObjDic.Fields("seq"), "0000"))
    
                            Select Case Trim(ObjDic.Fields("stscd"))
                                Case enStsCd.StsCd_LIS_Order
                                    strStsCd = "처방"
                                Case enStsCd.StsCd_LIS_Collection
                                    strStsCd = "채혈"
                                Case enStsCd.StsCd_LIS_Accession
                                    strStsCd = "접수"
                                Case enStsCd.StsCd_LIS_InProcess
                                    strStsCd = "검사중"
                                Case enStsCd.StsCd_LIS_MidRst
                                    strStsCd = "중간보고"
                                Case enStsCd.StsCd_LIS_FinRst
                                    strStsCd = "결과"
                                Case enStsCd.StsCd_LIS_Modify
                                    strStsCd = "수정"
                                Case enStsCd.StsCd_LIS_Cancel
                                    strStsCd = "취소"
                            End Select
    
                            itmFound.SubItems(1) = ObjDic.Fields("accseq")
                            itmFound.SubItems(2) = strStsCd
                            itmFound.SubItems(3) = ObjDic.Fields("ptid")
                            itmFound.SubItems(4) = ObjDic.Fields("name")
                            itmFound.SubItems(5) = ObjDic.Fields("field5")
                            strTime = Trim(ObjDic.Fields("rcvtm"))
                            itmFound.SubItems(6) = Mid(strTime, 1, 2) & "시 " & Mid(strTime, 3, 2) & "분"
                        End If
                    End If
    
                    I = I + 1
                    objPro.Value = I
                    ObjDic.MoveNext
                Loop
            End With
        End If
        
    '----------------------------------------------------------'
    '   보고일로 조회시
    '----------------------------------------------------------'
    ElseIf optVfydt = True Then
        
        strWorkArea = Trim(medGetP(cboWA.Text, 2, Space(60)))
        
        '모든검사 항목으로 조회시
        If strTestcd = "All" Then
            Set RS = objSQL.GetAccInfo02(strDate, strWorkArea, strTime)
            
            'ProgressBar 설정
            Set objPro = Nothing
            Set objPro = New jProgressBar.clsProgress
            With objPro
                .Container = Me
                .Left = lblTitle2.Left
                .Top = lblTitle2.Top
                .Width = lblTitle2.Width
                .Height = lblTitle2.Height
                .Max = RS.RecordCount
            End With
            
            I = 1
            lvwAccList.ListItems.Clear
            With lvwAccList
                Do Until RS.EOF
                    Set itmFound = .FindItem(Format(I, "0000"))
                    If itmFound Is Nothing Then
                        Set itmFound = .ListItems.Add(, , Format(I, "0000"))

                        Select Case Trim(RS.Fields("stscd").Value & "")
                            Case enStsCd.StsCd_LIS_Order
                                strStsCd = "처방"
                            Case enStsCd.StsCd_LIS_Collection
                                strStsCd = "채혈"
                            Case enStsCd.StsCd_LIS_Accession
                                strStsCd = "접수"
                            Case enStsCd.StsCd_LIS_InProcess
                                strStsCd = "검사중"
                            Case enStsCd.StsCd_LIS_MidRst
                                strStsCd = "중간보고"
                            Case enStsCd.StsCd_LIS_FinRst
                                strStsCd = "결과"
                            Case enStsCd.StsCd_LIS_Modify
                                strStsCd = "수정"
                            Case enStsCd.StsCd_LIS_Cancel
                                strStsCd = "취소"
                        End Select

                        itmFound.SubItems(1) = RS.Fields("accseq").Value & ""
                        itmFound.SubItems(2) = strStsCd
                        itmFound.SubItems(3) = RS.Fields("ptid").Value & ""
                        itmFound.SubItems(4) = RS.Fields("name").Value & ""
                        itmFound.SubItems(5) = RS.Fields("field5").Value & ""
                        strTime = Trim(RS.Fields("rcvtm").Value & "")
                        itmFound.SubItems(6) = Mid(strTime, 1, 2) & "시 " & Mid(strTime, 3, 2) & "분"
                        itmFound.SubItems(7) = RS.Fields("workarea").Value & "" & "-" & _
                                               Mid$(RS.Fields("accdt").Value & "", 3) & "-" & _
                                               RS.Fields("accseq").Value & ""
                    End If
    
                    I = I + 1
                    objPro.Value = I
                    RS.MoveNext
                Loop
            End With
            
        '지정된 검사항목으로 조회시
        Else
            Set RS = objSQL.GetVfyInfoByTime(strDate, 1, strWorkArea, strTime)
            
            Me.MousePointer = 11
            lvwAccList.ListItems.Clear
            
            objS.Clear:            objS.FieldInialize "spcno", "seq"
            objQ.Clear:            objQ.FieldInialize "testcd", "data"

            objQ.Sort = False: objS.Sort = False
            
            I = 1
            If Not RS.EOF Then
                Set objPro = Nothing
                Set objPro = New jProgressBar.clsProgress
                With objPro
                    .Container = Me
                    .Left = lblTitle2.Left
                    .Top = lblTitle2.Top
                    .Width = lblTitle2.Width
                    .Height = lblTitle2.Height
                    .Max = RS.RecordCount * 2
                End With
                
                Do Until RS.EOF
                    If objS.Exists(RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & "") Then
                    Else
                        objS.AddNew RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & "", I
                        I = I + 1
                    End If
                    ll = ll + 1
                    objPro.Value = ll
                    objPro.Message = " 검체저장 순번을 수집합니다......(" & ll & "/" & objPro.Max & ")"
                    RS.MoveNext
                Loop
                
                RS.MoveFirst
                Do Until RS.EOF
                    If objS.Exists(RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & "") Then
                        objS.KeyChange RS.Fields("spcyy").Value & "" & COL_DIV & RS.Fields("spcno").Value & ""
                        I = objS.Fields("seq")
                    End If

                    If objQ.Exists(RS.Fields("testcd").Value & "") Then
                        objQ.KeyChange RS.Fields("testcd").Value & ""
                        objQ.Fields("data") = objQ.Fields("data") & "★" & I & "§" & RS.Fields("workarea").Value & "" & "§" & _
                                                                   RS.Fields("accdt").Value & "" & "§" & _
                                                                   RS.Fields("accseq").Value & "" & "§" & _
                                                                   RS.Fields("stscd").Value & "" & "§" & _
                                                                   RS.Fields("ptid").Value & "" & "§" & _
                                                                   RS.Fields("name").Value & "" & "§" & _
                                                                   RS.Fields("field5").Value & "" & "§" & _
                                                                   RS.Fields("rcvtm").Value & ""
                    Else
                        objQ.AddNew RS.Fields("testcd").Value & "", I & "§" & RS.Fields("workarea").Value & "" & "§" & _
                                                                   RS.Fields("accdt").Value & "" & "§" & _
                                                                   RS.Fields("accseq").Value & "" & "§" & _
                                                                   RS.Fields("stscd").Value & "" & "§" & _
                                                                   RS.Fields("ptid").Value & "" & "§" & _
                                                                   RS.Fields("name").Value & "" & "§" & _
                                                                   RS.Fields("field5").Value & "" & "§" & _
                                                                   RS.Fields("rcvtm").Value & ""
                    End If
                    ll = ll + 1
                    objPro.Value = ll
                    objPro.Message = " 검사항목을 수집합니다..........(" & ll & "/" & objPro.Max & ")"
                    RS.MoveNext
                Loop
            End If
            
            objS.Sort = True
            objQ.Sort = True
            
            objQ.MoveFirst
            If objQ.Exists(strTestcd) Then
                objQ.KeyChange strTestcd
                aryD() = Split(objQ.Fields("data"), "★")
                Set objPro = Nothing
                Set objPro = New jProgressBar.clsProgress
                With objPro
                    .Container = Me
                    .Left = lblTitle2.Left
                    .Top = lblTitle2.Top
                    .Width = lblTitle2.Width
                    .Height = lblTitle2.Height
                    .Max = UBound(aryD)
                End With
                objPro.Message = "자료를 Display 합니다....."
                For kk = LBound(aryD) To UBound(aryD)
                    With lvwAccList
                        Set itmFound = .FindItem(Format(medGetP(aryD(kk), 1, "§"), "0000"))
                        If itmFound Is Nothing Then
                            Set itmFound = .ListItems.Add(, , Format(medGetP(aryD(kk), 1, "§"), "0000"))

                            Select Case Trim(medGetP(aryD(kk), 5, "§"))
                                Case enStsCd.StsCd_LIS_Order
                                    strStsCd = "처방"
                                Case enStsCd.StsCd_LIS_Collection
                                    strStsCd = "채혈"
                                Case enStsCd.StsCd_LIS_Accession
                                    strStsCd = "접수"
                                Case enStsCd.StsCd_LIS_InProcess
                                    strStsCd = "검사중"
                                Case enStsCd.StsCd_LIS_MidRst
                                    strStsCd = "중간보고"
                                Case enStsCd.StsCd_LIS_FinRst
                                    strStsCd = "결과"
                                Case enStsCd.StsCd_LIS_Modify
                                    strStsCd = "수정"
                                Case enStsCd.StsCd_LIS_Cancel
                                    strStsCd = "취소"
                            End Select

                            itmFound.SubItems(1) = medGetP(aryD(kk), 4, "§")
                            itmFound.SubItems(2) = strStsCd
                            itmFound.SubItems(3) = medGetP(aryD(kk), 6, "§")
                            itmFound.SubItems(4) = medGetP(aryD(kk), 7, "§")
                            itmFound.SubItems(5) = medGetP(aryD(kk), 8, "§")
                            strTime = Trim(medGetP(aryD(kk), 7, "§"))
                            itmFound.SubItems(6) = Mid(strTime, 1, 2) & "시 " & Mid(strTime, 3, 2) & "분"
                            itmFound.SubItems(7) = medGetP(aryD(kk), 2, "§") & "-" & _
                                                   medGetP(aryD(kk), 3, "§") & "-" & _
                                                   medGetP(aryD(kk), 4, "§")
                        End If

                    End With
                    ll = ll + 1
                    objPro.Value = ll

                Next
            End If
            Me.MousePointer = 0
            Set objQ = Nothing
            GoTo Skip
 
            '결과를 ListView에 출력
            I = 1
            lvwAccList.ListItems.Clear
            ObjDic.MoveFirst
    
            'ProgressBar 설정
            Set objPro = Nothing
            Set objPro = New jProgressBar.clsProgress
            With objPro
                .Container = Me
                .Left = lblTitle2.Left
                .Top = lblTitle2.Top
                .Width = lblTitle2.Width
                .Height = lblTitle2.Height
                .Max = ObjDic.RecordCount
            End With
            ObjDic.MoveFirst
    
            I = 1
            With lvwAccList
                Do Until ObjDic.EOF
                    DoEvents
                    If ObjDic.Fields("testcd") = strTestcd Then
                        Set itmFound = .FindItem(Format(ObjDic.Fields("seq"), "0000"))
                        If itmFound Is Nothing Then
                            Set itmFound = .ListItems.Add(, , Format(ObjDic.Fields("seq"), "0000"))
    
                            Select Case Trim(ObjDic.Fields("stscd"))
                                Case enStsCd.StsCd_LIS_Order
                                    strStsCd = "처방"
                                Case enStsCd.StsCd_LIS_Collection
                                    strStsCd = "채혈"
                                Case enStsCd.StsCd_LIS_Accession
                                    strStsCd = "접수"
                                Case enStsCd.StsCd_LIS_InProcess
                                    strStsCd = "검사중"
                                Case enStsCd.StsCd_LIS_MidRst
                                    strStsCd = "중간보고"
                                Case enStsCd.StsCd_LIS_FinRst
                                    strStsCd = "결과"
                                Case enStsCd.StsCd_LIS_Modify
                                    strStsCd = "수정"
                                Case enStsCd.StsCd_LIS_Cancel
                                    strStsCd = "취소"
                            End Select
    
                            itmFound.SubItems(1) = ObjDic.Fields("accseq")
                            itmFound.SubItems(2) = strStsCd
                            itmFound.SubItems(3) = ObjDic.Fields("ptid")
                            itmFound.SubItems(4) = ObjDic.Fields("name")
                            itmFound.SubItems(5) = ObjDic.Fields("field5")
                            strTime = Trim(ObjDic.Fields("rcvtm"))
                            itmFound.SubItems(6) = Mid(strTime, 1, 2) & "시 " & Mid(strTime, 3, 2) & "분"
                        End If
                    End If
    
                    I = I + 1
                    objPro.Value = I
                    ObjDic.MoveNext
                Loop
            End With
        End If
        
        
        '## 원본 ==========================================================================================================================
'        strWorkArea = Trim(medGetP(cboWA.Text, 2, Space(60)))
'        Set RS = objSQL.GetAccInfoByTime(strDate, 2, strWorkArea, strTime)
'
'        If Not RS.EOF Then
'            'ProgressBar 설정
'            Set objPro = Nothing
'            Set objPro = New jProgressBar.clsProgress
'            With objPro
'                .Container = Me
'                .Left = lblTitle2.Left
'                .Top = lblTitle2.Top
'                .Width = lblTitle2.Width
'                .Height = lblTitle2.Height
'                .Max = RS.RecordCount
'            End With
'
'            I = 1
'            lvwAccList.ListItems.Clear
'            With lvwAccList
'                Do Until RS.EOF
'                    Set itmFound = .FindItem(Format(I, "0000"))
'                    If itmFound Is Nothing Then
'                        Set itmFound = .ListItems.Add(, , Format(I, "0000"))
'
'                        Select Case Trim(RS.Fields("stscd").Value & "")
'                            Case enStsCd.StsCd_LIS_Order
'                                strStsCd = "처방"
'                            Case enStsCd.StsCd_LIS_Collection
'                                strStsCd = "채혈"
'                            Case enStsCd.StsCd_LIS_Accession
'                                strStsCd = "접수"
'                            Case enStsCd.StsCd_LIS_InProcess
'                                strStsCd = "검사중"
'                            Case enStsCd.StsCd_LIS_MidRst
'                                strStsCd = "중간보고"
'                            Case enStsCd.StsCd_LIS_FinRst
'                                strStsCd = "결과"
'                            Case enStsCd.StsCd_LIS_Modify
'                                strStsCd = "수정"
'                            Case enStsCd.StsCd_LIS_Cancel
'                                strStsCd = "취소"
'                        End Select
'
'                        itmFound.SubItems(1) = RS.Fields("accseq").Value & ""
'                        itmFound.SubItems(2) = strStsCd
'                        itmFound.SubItems(3) = RS.Fields("ptid").Value & ""
'                        itmFound.SubItems(4) = RS.Fields("name").Value & ""
'                        itmFound.SubItems(5) = RS.Fields("field5").Value & ""
'
'                        strTime = Trim(RS.Fields("rcvtm").Value & "")
'                        itmFound.SubItems(6) = Mid(strTime, 1, 2) & "시 " & Mid(strTime, 3, 2) & "분"
'                        itmFound.SubItems(7) = RS.Fields("workarea").Value & "" & "-" & _
'                                               Mid$(RS.Fields("accdt").Value & "", 3) & "-" & _
'                                               RS.Fields("accseq").Value & ""
'
'                    End If
'
'                    I = I + 1
'                    objPro.Value = I
'                    RS.MoveNext
'                Loop
'            End With
'        End If
        '==================================================================================================================================
        
    End If

Skip:
 
    Set RS = Nothing
    Set objPro = Nothing
    Set ObjDic = Nothing
End Sub

Private Sub cmdRefleshCount_Click()
    Dim RS As New Recordset
    Dim itmFound As ListItem
    Dim lngTotal As Long
    
    Set RS = objSQL.GetWorkAccount(Format(dtpDate2.Value, "yyyyMMdd"))
    
    '검체접수현황을 콤보박스에 출력
    lvwCount.ListItems.Clear
    If Not RS.EOF Then
        With lvwCount
            Do Until RS.EOF
                Set itmFound = .FindItem(RS.Fields("field1").Value & "")
                If itmFound Is Nothing Then
                    Set itmFound = .ListItems.Add(, , RS.Fields("field1").Value & "")
                    itmFound.SubItems(1) = Trim(RS.Fields("cnt").Value & "")
                    lngTotal = lngTotal + Val(RS.Fields("cnt").Value & "")
                End If
                
                RS.MoveNext
            Loop
            
            Set itmFound = .ListItems.Add(, , "합계")
            itmFound.SubItems(1) = lngTotal
        End With
    End If
    
    Set RS = Nothing
End Sub

Private Sub cmdWSList_Click()
    If lstWSCode.ListCount = 0 Then
        MsgBox "등록된 WorkSheet 코드가 없습니다.", vbExclamation, "메세지"
        Exit Sub
    End If
    
    lstWSCode.Visible = True
    lstWSCode.ZOrder 0
End Sub

Private Sub cmdClose_Click()
    frmPop.Visible = False
End Sub

Private Sub dtpDate1_Change()
'    dtpFrom.Value = dtpDate1.Value
'    dtpTo.Value = dtpDate1.Value
End Sub

Private Sub dtpDate2_Change()
    Call cmdRefleshCount_Click
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        lstWSCode.Visible = False
    End If
End Sub

Private Sub Form_Load()
    '날짜를 DB시간으로 설정
    dtpDate1.Value = GetSystemDate
    dtpDate2.Value = GetSystemDate
    dtpTime.Value = "00:00"
    
    Call GetWorkArea
    Call cmdRefleshCount_Click
    
    DoEvents
    
    'WorkSheet를 ListView로 불러옴
    Set objPtInfo = New clsPatientInfo
    Call objPtInfo.LoadWorkSheetCode(ObjSysInfo.BuildingCd, lstWSCode)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set objSQL = Nothing
    Set MySql = Nothing
    Set objPtInfo = Nothing
    Set objLab301 = Nothing
End Sub

Private Sub lstWSCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn:
        txtWorkCd.Text = medGetP(lstWSCode.Text, 2, vbTab)
        lstWSCode.Visible = False
        Call txtWorkCd_Validate(False)
    End Select
End Sub

Private Sub lstWSCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call lstWSCode_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvwAccList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer
    
    With lvwAccList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwAccList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim objMyPatient As clsPatient
    Dim objDisease As S2LIS_ReportLib.clsDisease
    
    Dim strWardId As String
    Dim lngAccSeq As Long
    
    Set objMyPatient = New clsPatient
    Set objDisease = New S2LIS_ReportLib.clsDisease
    
    Call ClearLabel
    txtPtId.Text = Item.SubItems(3)
    
    Call ICSPatientMark(txtPtId.Text, enICSNum.ResultReview)
    
    With objMyPatient
        .GETPatient (txtPtId.Text)
        lblName.Caption = .PtNm
        lblSexAge.Caption = .SEXNM & " / " & .Age
        lblDeptNm.Caption = .DeptNm
        
        strWardId = .WardId
        If strWardId <> "" Then
            If .RoomId <> "" Then
                strWardId = strWardId & "-" & .RoomId
            End If
        End If
        lblLocation.Caption = strWardId
        
        objDisease.PTid = txtPtId.Text
        lblDisease.Caption = objDisease.Disease
    End With
    
    'WA별 접수번호
    lngAccSeq = Val(Item.SubItems(1))
    Call GetResultList(lngAccSeq)
    
    tblOrdSheet.SetFocus
    
    OldRow = 0
    Set objMyPatient = Nothing
    Set objDisease = Nothing
End Sub

Private Sub lvwAccList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then      '오른쪽 마우스버튼 클릭시
        'ListView에 조회된 결과가 없거나, 보고일 조회시
        If lvwAccList.ListItems.Count = 0 Or optVfydt.Value = True Then Exit Sub
        
        Set objPop = New clsPopupMenu
        With objPop
            .AddMenu MENU_PRT, "Print"
            .PopupMenus Me.hwnd
        End With
        Set objPop = Nothing
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : ListView에 있는 내용을 출력 - 이상대(2004-12-17)
'-----------------------------------------------------------------------------'
Private Sub objPop_Click(ByVal vMenuID As Long)
    Dim objPrint    As clsLISPrint
    Dim strTitle    As String   'Title
    Dim strHeader   As String   'Header
    Dim strColumns  As String   'Column Header
    Dim strContent  As String   'Content
    Dim strTemp     As String
    Dim I           As Long
    
    If lvwAccList.ListItems.Count = 0 Then Exit Sub
    
    If vMenuID = MENU_PRT Then
        If optAccdt.Value = True Then
            strTemp = "접수일 "
        ElseIf optVfydt.Value = True Then
            strTemp = "보고일 "
        End If
        
        strTemp = strTemp & Format$(dtpDate1.Value, "YYYY-MM-DD") & " " & Format$(dtpTime.Value, "HH:MM")
        
        strTitle = "『접수리스트』"
        strHeader = "※ 조회조건 : " & strTemp & COL_DIV & "5" & COL_DIV & "1"
        strHeader = strHeader & vbTab & "※ 출력일시 : " & Format$(Now, "YYYY-MM-DD") & " " & _
                    Format$(Now, "HH:MM") & COL_DIV & "5" & COL_DIV & "1"
        
        strColumns = "순번" & COL_DIV & "5" & COL_DIV & "0"
        strColumns = strColumns & vbTab & "접수번호" & COL_DIV & "15" & COL_DIV & "0"
        strColumns = strColumns & vbTab & "환자번호" & COL_DIV & "50" & COL_DIV & "0"
        strColumns = strColumns & vbTab & "환자명" & COL_DIV & "70" & COL_DIV & "0"
        strColumns = strColumns & vbTab & "검체명" & COL_DIV & "90" & COL_DIV & "0"
        strColumns = strColumns & vbTab & "상태" & COL_DIV & "110" & COL_DIV & "0"
        strColumns = strColumns & vbTab & "접수시간" & COL_DIV & "135" & COL_DIV & "1"
        
        With lvwAccList
            For I = 1 To .ListItems.Count
                strContent = strContent & .ListItems(I).Text & COL_DIV & "5" & COL_DIV & "0" & COL_DIV & "0"
                strContent = strContent & vbTab & .ListItems(I).SubItems(7) & COL_DIV & "15" & COL_DIV & "0" & COL_DIV & "0"
                strContent = strContent & vbTab & .ListItems(I).SubItems(3) & COL_DIV & "50" & COL_DIV & "0" & COL_DIV & "0"
                strContent = strContent & vbTab & .ListItems(I).SubItems(4) & COL_DIV & "70" & COL_DIV & "0" & COL_DIV & "0"
                strContent = strContent & vbTab & .ListItems(I).SubItems(5) & COL_DIV & "90" & COL_DIV & "0" & COL_DIV & "0"
                strContent = strContent & vbTab & .ListItems(I).SubItems(2) & COL_DIV & "110" & COL_DIV & "0" & COL_DIV & "0"
                strContent = strContent & vbTab & .ListItems(I).SubItems(6) & COL_DIV & "135" & COL_DIV & "1" & COL_DIV & "1" & vbTab
            Next I
        End With
    End If
    
On Error GoTo Errors
    strContent = Mid(strContent, 1, Len(strContent) - 1)
    Set objPrint = New clsLISPrint
    With objPrint
        .PrinterHeader1 = strTitle
        .PrinterHeader2 = strHeader
        .PrinterHeader3 = strColumns
        .PrinterBody = strContent
        Call .CallPrint
    End With
    Exit Sub

Errors:
    MsgBox "출력중 에러가 발생했습니다.", vbCritical, "오류"
End Sub

'Private Sub mnuPrint_Click()
'    Dim strFileNm As String     'CrystalReport.txt 파일
'    Dim strRptNm As String      'Report 파일
'    Dim strTmp As String
'    Dim lngCnt As Long
'    Dim lngFNum As Long
'    Dim strWaWs As String
'    Dim strTime As String
'    Dim I As Integer
'
'    If optAccdt.Value = True Then
'        strWaWs = Trim(medGetP(cboWA.Text, 1, Space(60)))
'    ElseIf optWorkSheet.Value = True Then
'        strWaWs = txtWorkCd.Text
'    End If
'    strTime = Format(dtpTime.Value, "HH:mm")
'
'    strFileNm = Dir(App.Path & "\..\rpt\CrystalReport.txt")
'    If strFileNm = "" Then
'        MsgBox "CrystalReport.txt 파일이 없습니다.", vbCritical, "정보확인"
'        Exit Sub
'    End If
'    strFileNm = App.Path & "\..\Rpt\" & strFileNm
'
'    strRptNm = Dir(App.Path & "\..\Rpt\rptAcctm.rpt")
'    If strRptNm = "" Then
'        MsgBox "rptAcctm.rpt 파일이 없습니다.", vbCritical, "정보확인"
'        Exit Sub
'    End If
'    strRptNm = App.Path & "\..\Rpt\" & strRptNm
'
'    strTmp = ""
'
'    Dim sICSString As String
'
'
'    With lvwAccList
'        For I = 1 To .ListItems.Count
'            strTmp = strTmp & .ListItems(I).Text & vbTab
'            strTmp = strTmp & .ListItems(I).SubItems(1) & vbTab
'            strTmp = strTmp & .ListItems(I).SubItems(3) & vbTab
'
'            sICSString = ICSPatientString(.ListItems(I).SubItems(3), enICSNum.LIS_ALL)
'
'
'            strTmp = strTmp & .ListItems(I).SubItems(4) & sICSString & vbTab
'            strTmp = strTmp & .ListItems(I).SubItems(2) & vbTab
'            strTmp = strTmp & .ListItems(I).SubItems(5) & vbTab
'            strTmp = strTmp & .ListItems(I).SubItems(6) & vbCr
'        Next I
'
'        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
'    End With
'
'    lngFNum = FreeFile
'
'On Error GoTo ErrPrint
'    Open strFileNm For Output As #lngFNum
'    Print #lngFNum, strTmp
'    Close #lngFNum
'
'    With crtReport
'        .ReportFileName = strRptNm
'        .ParameterFields(0) = "WaWsTime;" & strTime & " 이후 " & strWaWs & "별" & ";true"
'        .ParameterFields(1) = "HospNm;" & P_HOSPITALNAME & " 임상병리과" & ";true"
'        .RetrieveDataFiles
'        .WindowState = 2
'        .Destination = crptToPrinter
'        .Action = 1
'        .Reset
'    End With
'    Exit Sub
'
'ErrPrint:
'    MsgBox "출력이 되지 않았습니다.", vbCritical
'
'End Sub

Private Sub optAccdt_Click()
    If optAccdt.Value = True Then
        lblDate.Caption = "접수일자 :"
        lblKindDate.Caption = "접수일"
        
        If lblQKind.Caption = "WorkSheet :" Then
            lblQKind.Caption = "Work Area :"
            txtWorkCd.Visible = False
            cmdWSList.Visible = False
            cboWA.Visible = True
        End If
        
        lblTestCd.Visible = True
        cboTestCd.Visible = True
    End If
End Sub

Private Sub optVfydt_Click()
    If optVfydt.Value = True Then
        lblDate.Caption = "보고일자 :"
        lblKindDate.Caption = "보고일"
        
        If lblQKind.Caption = "WorkSheet :" Then
            lblQKind.Caption = "Work Area :"
            txtWorkCd.Visible = False
            cmdWSList.Visible = False
            cboWA.Visible = True
        End If
        
'        lblTestCd.Visible = False
'        cboTestCd.Visible = False
    End If
End Sub

'*--------------------------------------------------------
'*  기능 : 환자정보 Label들을 Clear
'*--------------------------------------------------------
Private Sub ClearLabel()
    txtPtId.Text = ""
    lblName.Caption = ""
    lblSexAge.Caption = ""
    lblDisease.Caption = ""
    lblDeptNm.Caption = ""
    lblLocation.Caption = ""
    lblVerifierNm.Caption = ""
    lblVerifyDt.Caption = ""
End Sub

'*--------------------------------------------------------
'*  기능 : ComboBox로 WorkArea명를 불러옴
'*--------------------------------------------------------
Private Sub GetWorkArea()
    Dim RS As New Recordset
    
    Set objSQL = New S2LIS_SqlLib.clsLISHospital05
    
    Set RS = objSQL.LoadWorkArea

    cboWA.Clear
    If Not RS.EOF Then
        With RS
            Do Until .EOF
                cboWA.AddItem .Fields("field1").Value & "" & Space(60) & .Fields("cdval1").Value & ""
                .MoveNext
            Loop
        End With
    End If
    
    'cboSpcCd 콤보박스에 검체명, 검체코드 출력
    Set RS = objSQL.LoadSpcCd
    cboSpcCd.Clear
    If Not RS.EOF Then
        cboSpcCd.AddItem "모든검체대상" & Space(60) & "All"
        With RS
            Do Until .EOF
                cboSpcCd.AddItem .Fields("field5").Value & "" & Space(60) & .Fields("cdval1").Value & ""
                .MoveNext
            Loop
        End With
    End If
    
    Set RS = Nothing
    cboSpcCd.ListIndex = 0
    cboWA.ListIndex = 0
End Sub

Private Sub GetResultList(ByVal lngAccSeq As Long)
    Dim objPro As jProgressBar.clsProgress
    Dim tmpRs As New Recordset
    
    Dim strKind     As String   '접수일(rcvdt), 보고일(vfydt) 구분자
    Dim strSQL      As String
    Dim strDate     As String
    Dim strWorkArea As String
    Dim strPtId     As String
    Dim SvKeyDt     As String
    Dim SvSpcNm     As String
    Dim tmpTestNm   As String
    Dim pWorkArea   As String
    Dim pAccDt      As String
    Dim pAccSeq     As String
    Dim strNotice   As String
    Dim strTmp      As String
    Dim ColCnt      As Integer
    Dim I           As Integer
    Dim J           As Integer

    '접수리스트 ComboBox Disable
'    lvwAccList.Enabled = False
    cmdPre.Enabled = False
    cmdNext.Enabled = False
    
    On Error GoTo ErrTrap
    
    'ProgressBar 설정
    Set objPro = Nothing
    Set objPro = New jProgressBar.clsProgress
    With objPro
        .Container = Me
        .Left = picProcess.Left
        .Top = picProcess.Top
        .Width = picProcess.Width
        .Height = picProcess.Height
        .Message = lblName.Caption & " 님의 검사 결과내역을 검색중입니다..."
        .Max = 100
    End With
    
    If optAccdt.Value = True Or optWorkSheet.Value = True Then
        strKind = "x.rcvdt"
    ElseIf optVfydt.Value = True Then
        strKind = "x.examdt"
    End If
    
    '날짜, WA, 환자ID
    strDate = Format(dtpDate1.Value, "yyyyMMdd")
    If optWorkSheet.Value = True Then
        strWorkArea = medGetP(lstWSCode.Text, 3, vbTab)
    Else
        strWorkArea = Trim(medGetP(cboWA.Text, 2, Space(60)))
    End If
    strPtId = txtPtId.Text
    objPro.Value = 20
    
    '## 5.0.4: 이상대(2005-07-12)
    '   - Worksheet별 조회일때는 접수일, 보고일조건이 의미없어 GetResultQueryNewX 함수를
    '     이용 하도록 수정
    If optWorkSheet.Value = False Then
        strSQL = objSQL.GetResultQueryNew(strDate, strKind, strWorkArea, strPtId, lngAccSeq)
    Else
        strSQL = objSQL.GetResultQueryNewX(strKind, strPtId, strWorkArea, lngAccSeq)
    End If
    tmpRs.Open strSQL, DBConn
    SvKeyDt = "": SvSpcNm = ""
    objPro.Value = 40
    
    DoEvents
   
    ReDim aryMesg(0)
    
    Call medClearTable(tblOrdSheet)
    If tmpRs.EOF Then GoTo NoData

    With tblOrdSheet
        Do Until tmpRs.EOF
            If StopFg Then
                StopFg = False
                GoTo NoData
            End If
            
            If objPro.Value >= objPro.Max Then objPro.Max = objPro.Max + 50
            objPro.Value = objPro.Value + 1

            DoEvents
            
            If .MaxRows < .DataRowCnt Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            Else
                .Row = .DataRowCnt + 1
            End If
            
            .RowHeight(.MaxRows) = 11.5
            
            If SvKeyDt <> Trim("" & tmpRs.Fields("KeyDate").Value) Then
                .Col = 1:   .Value = Trim("" & tmpRs.Fields("KeyDate").Value)
                            .FontBold = True: .ForeColor = vbBlack       '-- 처방일
                .Col = 2:   .Value = Trim("" & tmpRs.Fields("SpcNm").Value)
                            .FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                SvKeyDt = Trim("" & tmpRs.Fields("KeyDate").Value)
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            Else
                .Col = 1:   .Value = Trim("" & tmpRs.Fields("KeyDate").Value): .ForeColor = .BackColor
                    If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                        .Col = 2:
                        .Value = Trim("" & tmpRs.Fields("SpcNm").Value)
                        .FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                        SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
                    Else
                        .Col = 2:
                        .Value = Trim("" & tmpRs.Fields("SpcNm").Value): .ForeColor = .BackColor

                        '.FontBold = True: .ForeColor = DCM_LightRed  '-- 검체명
                    End If
            End If
            
            .Col = 32:  .Value = Trim("" & tmpRs.Fields("KeyDate").Value)   '처방일
            .Col = 33:  .Value = Trim("" & tmpRs.Fields("SpcNm").Value)     '검체명
            
            .Col = 3:   '-- 검사명
                        .ForeColor = DCM_MidBlue
                        tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestLongNm").Value), 1, 33)
                        If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                            Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                            Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                            
                            .Value = tmpTestNm & " " & String(35 - Len(tmpTestNm), ".")
                        Else
                            .Value = Space(4) & tmpTestNm & " " & String(35 - Len("  " & tmpTestNm), ".")
                        End If
                        
            .Col = 4:   '-- 결과명(코드일 경우..)
                        .ForeColor = DCM_Brown   '갈색
                        If Trim("" & tmpRs.Fields("VfyDt").Value) = "" Then
                            .Value = "미확"
                            .ForeColor = DCM_MidGray: .FontBold = False:
                        Else
                            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                                .TypeHAlign = TypeHAlignCenter
                                .Value = Trim("" & tmpRs.Fields("RstCd").Value)
                            Else
                                .CellType = CellTypeEdit
                                .TypeHAlign = TypeHAlignLeft
                                .Value = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
                            End If
                            If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then
                                .Value = "Growth"
                            ElseIf Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                                .Value = Space(3)
                            End If
                        End If
                        
            .Col = 5:   '-- 결과단위
                        .Value = Trim("" & tmpRs.Fields("RstUnit").Value)
            
            .Col = 6    '-- High / Low
                        .Value = ""
                        If Trim("" & tmpRs.Fields("VfyDt").Value) <> "" Then
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_HIGH_CD Then .Value = HLDIV_HIGH_FG: .ForeColor = DCM_LightRed
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = HLDIV_LOW_CD Then .Value = HLDIV_LOW_FG:  .ForeColor = DCM_LightBlue
                            If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then .Value = "*": .ForeColor = vbRed
                        End If
            
            .Col = 7:   '-- Delta/Panic
                        .Value = Trim("" & tmpRs.Fields("DPDiv").Value): .ForeColor = vbRed
            
            .Col = 8:   '-- 참고치
                        If Trim("" & tmpRs.Fields("RstDiv").Value) <> "*" And Trim("" & tmpRs.Fields("TestDiv").Value) < "4" Then .Value = CS_QuestionMark
            
            .Col = 9:   '-- More Result...
                        .Value = "": .ForeColor = DCM_LightBlue
                        If Trim("" & tmpRs.Fields("TxtFg").Value) > "0" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.Fields("TxtFg").Value) = "Y" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.Fields("SenFg").Value) = "Y" Then .Value = CS_FingerMark
                        If (Trim("" & tmpRs.Fields("DetailFg").Value) = "" And _
                            Trim("" & tmpRs.Fields("DetailItem").Value) = "") Or _
                            Trim("" & tmpRs.Fields("RstDiv").Value) = "*" Then
                            If Trim("" & tmpRs.Fields("FootNoteFg").Value) = "1" Then .Value = CS_FingerMark
                            If Trim("" & tmpRs.Fields("RmkCd").Value) <> "" Then .Value = CS_FingerMark
                        End If
                        If Trim("" & tmpRs.Fields("DcFg").Value) = "1" Then .Value = .Value & "*"
                        If Trim("" & tmpRs.Fields("TestDiv").Value) = "4" Then .Value = CS_FingerMark    '해부병리
                        If Trim("" & tmpRs.Fields("TestDiv").Value) = "5" Then .Value = CS_FingerMark    '혈액은행
                        If Trim("" & tmpRs.Fields("OrdDiv").Value) = CMT_ORDDIV Then .Value = CS_FingerMark    '종합검증
         
            .Col = 10: .Value = Trim("" & tmpRs.Fields("OrdDate").Value)        '-- 처방일
            .Col = 11: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)          '-- 처방번호
            .Col = 12: .Value = Trim("" & tmpRs.Fields("OrdDoct").Value)        '-- 처방의
            .Col = 13: .Value = Trim("" & tmpRs.Fields("ColDtTm").Value)        '-- 채혈일시
            .Col = 14: .Value = Trim("" & tmpRs.Fields("ColId").Value)          '-- 채혈자
            .Col = 15: .Value = Trim("" & tmpRs.Fields("RcvDtTm").Value)        '-- 접수일시
            .Col = 16: .Value = Trim("" & tmpRs.Fields("RcvId").Value)          '-- 접수자
            .Col = 17: .Value = Trim("" & tmpRs.Fields("WorkArea").Value):  pWorkArea = .Value  'WorkArea
            .Col = 18: .Value = Trim("" & tmpRs.Fields("AccDt").Value):     pAccDt = .Value     'AccDt
            .Col = 19: .Value = Trim("" & tmpRs.Fields("AccSeq").Value):    pAccSeq = .Value    'AccSeq
            .Col = 20: .Value = Trim("" & tmpRs.Fields("LastRst").Value)        '-- 최근결과
            .Col = 21: .Value = Trim("" & tmpRs.Fields("LstVfyDtTm").Value)     '-- 최근결과일시
            .Col = 22: .Value = Trim("" & tmpRs.Fields("LastVfyId").Value)      '-- 최근결과 보고자
            .Col = 23: .Value = Trim("" & tmpRs.Fields("VfyDtTm").Value)        '-- 보고일시
            .Col = 24: .Value = Trim("" & tmpRs.Fields("VfyId").Value)          '-- 보고자
            .Col = 25: .Value = Trim("" & tmpRs.Fields("Sex").Value)            '-- Sex
            .Col = 26: .Value = Trim("" & tmpRs.Fields("AgeDay").Value)         '-- AgeDay
            .Col = 27: .Value = Trim("" & tmpRs.Fields("TestCd").Value)         '-- 검사코드
            .Col = 28: .Value = Trim("" & tmpRs.Fields("SpcCd").Value)          '-- 검체코드
            .Col = 29: .Value = Trim("" & tmpRs.Fields("VfyDt").Value)          '-- 보고일
            .Col = 30: .Value = Trim("" & tmpRs.Fields("TestDiv").Value)        '-- 검사구분
            .Col = 31: .Value = Trim("" & tmpRs.Fields("DeptCd").Value)         '-- 진료과
            .Col = 34: .Value = Trim("" & tmpRs.Fields("TxtFg").Value)          '-- 소견결과여부
            .Col = 35: .Value = Trim("" & tmpRs.Fields("FootNoteFg").Value)     '-- Footnote 여부
            .Col = 36: .Value = Trim("" & tmpRs.Fields("RmkCd").Value)          '-- Remark 코드
            .Col = 37: .Value = Trim("" & tmpRs.Fields("SenFg").Value)          '-- 감수성 여부
            .Col = 38: .Value = Trim("" & tmpRs.Fields("OrdDiv").Value)         '-- 처방구분
            .Col = 39: .Value = Trim("" & tmpRs.Fields("UnitQty").Value)        '-- 수혈수량
            .Col = 40: .Value = Trim("" & tmpRs.Fields("ReqDt").Value)          '-- 수혈예정일
            .Col = 41: .Value = Trim("" & tmpRs.Fields("ReqTm").Value)          '-- 수혈예정시간
            .Col = 42: .Value = Trim("" & tmpRs.Fields("WardId").Value)         '-- 병동
            .Col = 43: .Value = Trim("" & tmpRs.Fields("HosilId").Value)        '-- 호실
            .Col = 44: .Value = Trim("" & tmpRs.Fields("RoomId").Value)        '-- 호실
            .Col = 45: .Value = Trim("" & tmpRs.Fields("Notice").Value)        '-- 호실
            
            ReDim Preserve aryMesg(UBound(aryMesg) + 1)
            aryMesg(UBound(aryMesg)) = Trim("" & tmpRs.Fields("Mesg").Value)    '-- 진료과Remark
         
            If Trim("" & tmpRs.Fields("Notice").Value) <> "" Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 3
                .TypeEditMultiLine = False
                .ForeColor = vbBlack
                .Value = "☞ Clinical Notice "  '& vbCrLf & Trim("" & tmpRs.fields("Notice").value)
                .RowHeight(.MaxRows) = .MaxTextRowHeight(.MaxRows)
                strNotice = Trim("" & tmpRs.Fields("Notice").Value)
                strNotice = Replace(strNotice, vbCr, "")
                strTmp = medShift(strNotice, vbLf)
                While strTmp <> ""
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 3
                    .TypeEditMultiLine = False
                    .ForeColor = &H747474
                    .Value = strTmp
                    strTmp = medShift(strNotice, vbLf)
                Wend
            End If
            tmpRs.MoveNext
        Loop
      
        .MaxRows = .DataRowCnt
        .Row = -1: .Col = 3: .Col2 = 5
        .BlockMode = True
        .AllowCellOverflow = True
        .BlockMode = False
        .ReDraw = True
      
        If chkRefVal.Value = 0 Then GoTo ExitPos
      
        Dim tmpTestCd As String
        Dim tmpSpcCd As String
        Dim tmpVfyDt As String
        Dim tmpSex As String
        Dim tmpAgeDay As String
        Dim tmpRs1 As New Recordset
        Dim tmpRefFromVal As Double
        Dim tmpRefToVal As Double
        Dim tmpRefCd As String
      
        objPro.Value = objPro.Max - 10
        objPro.Message = "임상 참고치를 검색하고 있습니다.."
        DoEvents
        For I = 1 To .MaxRows
            If objPro.Value < objPro.Max Then objPro.Value = objPro.Value + 1
            '참고치 검색
            .Row = I
            .Col = 8: If .Value <> CS_QuestionMark Then GoTo RefSkip
            
            .Col = 25:  tmpSex = Trim(.Value)
            .Col = 26:  tmpAgeDay = Trim(.Value)
            .Col = 27:  tmpTestCd = Trim(.Value)
            .Col = 28:  tmpSpcCd = Trim(.Value)
            .Col = 29:  tmpVfyDt = Trim(.Value)
                        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
         
            strSQL = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = Nothing
            Set tmpRs1 = New Recordset
            tmpRs1.Open strSQL, DBConn
            If tmpRs1.EOF Then
                '"B"(Both)에 해당하는 참고치가 없는 경우 환자성별에 해당하는 데이타 검색
                '--> 거의 Both로 등록됨.
                strSQL = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = Nothing
                Set tmpRs1 = New Recordset
                tmpRs1.Open strSQL, DBConn
            End If
            If tmpRs1.EOF Then
                tmpRefCd = Space(5)
            Else
                tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
                If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
                   tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
            End If
            Set tmpRs1 = Nothing
            For J = I To .MaxRows
                .Row = J
                .Col = 27   '참고치
                If Trim(.Value) = tmpTestCd Then _
                    .Col = 8: .Value = tmpRefCd: .ForeColor = DCM_Green
            Next
         
            DoEvents

RefSkip:
        Next
      
ExitPos:
        objPro.Value = objPro.Max
        DoEvents
'        medSleep (500)
        If .MaxRows < 27 Then .MaxRows = 27
    End With
   
NoData:
    Me.Enabled = True
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    
'    lvwAccList.Enabled = True
    cmdPre.Enabled = True
    cmdNext.Enabled = True
    Set objPro = Nothing
    Exit Sub

ErrTrap:
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
    MouseDefault
    Set objPro = Nothing
    
End Sub

Private Sub optWorkSheet_Click()
    If optWorkSheet.Value = True Then
        lblDate.Caption = "작성일 :"
        lblKindDate.Caption = "접수일"
        lblQKind.Caption = "WorkSheet :"
        
'        lblTestCd.Visible = False
'        cboTestCd.Visible = False
        cboWA.Visible = False
        txtWorkCd.Visible = True
        cmdWSList.Visible = True
        dtpTime.SetFocus
    End If
End Sub

Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
    Dim pWorkArea As String
    Dim pAccDt As String
    Dim pAccSeq As String
    Dim pTestDiv As String
    Dim strOrdDiv As String
    Dim strWardId As String
    Dim strHosilId As String
    Dim tmpResult As New clsLISResultReview
    Dim strRoomId As String '추가사항(병실)

    If Row = 0 Then Exit Sub
    If OldRow = Row Then Exit Sub

    With tblOrdSheet
        .Row = Row
        .Col = 46:
                If .Value = "" Or .Value = "0" Then
                    .Value = "1"
                ElseIf .Value = "1" Then
                    .Value = "0"
                End If
        .Col = 3:  If .Value = "" Then Exit Sub

        .Col = 17: pWorkArea = .Value
        .Col = 18: pAccDt = .Value
        .Col = 19: pAccSeq = .Value
        .Col = 30: pTestDiv = .Value
        .Col = 38: strOrdDiv = .Value
        .Col = 42: strWardId = .Value
        .Col = 43: strHosilId = .Value
        .Col = 44: strRoomId = .Value   '추가사항

        If strWardId <> "" Then
            lblLocation.Caption = strWardId & " - " & strHosilId
            If Trim(strRoomId) <> "" Then lblLocation.Caption = lblLocation.Caption & " - " & strRoomId
        Else
            lblLocation.Caption = ""
        End If

        If (pWorkArea = "" Or pAccDt = "" Or pAccSeq = "") And strOrdDiv <> BBS_ORDDIV And strOrdDiv <> POC_ORDDIV And strOrdDiv <> CMT_ORDDIV Then
            MsgBox "접수번호가 없습니다. (전산실로 연락바람 ☎" & ObjSysInfo.HelpLine & ")", vbExclamation, "오류발생"
            Exit Sub
        End If

        If OldRow > 0 Then
            .Row = OldRow
            .Col = -1
            .BackColor = OldBackColor
        End If

        .Row = Row
        .Col = -1
        OldRow = Row
        OldBackColor = .BackColor
        .BackColor = &HD9ECFF ' &HFCEFE9   ' &HF5FFF4       '연두색


        .Col = 8: '참고치
        If Trim(.Value) = CS_QuestionMark Then Call GetRefValue(Row)

        .Col = 23:  lblVerifyDt.Caption = .Value                        '보고일시
        .Col = 24:  lblVerifierNm.Caption = GetEmpNm(.Value)  '보고자
        .Col = 31:  lblDeptNm.Caption = GetDeptNm(.Value)     '진료과
        
        Call ResultClear
        .Col = 33:   lblSpecimenNm.Caption = .Value '검체
        
        DoEvents
'        If pTestDiv = "2" Or pTestDiv = "1" Then
            Call DisplayLISResult(pWorkArea, pAccDt, Val(pAccSeq), pTestDiv)
'        End If
        Screen.MousePointer = vbDefault

        Set tmpResult = Nothing
    End With
End Sub

Private Sub GetRefValue(ByVal iRow As Integer)
    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New Recordset
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    Dim SqlStmt As String
      
    With tblOrdSheet
        '기준치 검색
        .Row = iRow
        .Col = 8: If .Value <> CS_QuestionMark Then Exit Sub
        
        .Col = 25:    tmpSex = Trim(.Value)
        .Col = 26:    tmpAgeDay = Trim(.Value)
        .Col = 27:    tmpTestCd = Trim(.Value)
        .Col = 28:    tmpSpcCd = Trim(.Value)
        .Col = 29:    tmpVfyDt = Trim(.Value)
                      If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
        
        SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
        Set tmpRs1 = Nothing
        Set tmpRs1 = New Recordset
        tmpRs1.Open SqlStmt, DBConn
        
        If tmpRs1.EOF Then
           SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
           Set tmpRs1 = Nothing
           Set tmpRs1 = New Recordset
           tmpRs1.Open SqlStmt, DBConn
        End If
        
        If tmpRs1.EOF Then
           tmpRefCd = Space(5)
        Else
           tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
           tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
           tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
           If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
        End If
        
        Set tmpRs1 = Nothing
        .Col = 8: .ForeColor = &H8000&
        If Trim(tmpRefCd) = "" Then
            .Value = "없음"
        Else
            .Value = tmpRefCd:
        End If
    End With
    
    Set tmpRs1 = Nothing
End Sub

Private Sub tblOrdSheet_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim tmpToolTip As String
    Dim MyResult As New clsLISResultReview
    Dim strSQL As String
    Dim rsMod As Recordset
    Dim tmpColNm As String
    Dim pWorkArea As String
    Dim pAccDt As String
    Dim pAccSeq As String
    
    tmpToolTip = vbCrLf
   
    With tblOrdSheet
        .Row = Row
       
        If Col = 4 Then
            .Col = 4
            If Len(.Value) > 20 Then
                MultiLine = 1
                TipWidth = 4000
                tmpToolTip = vbCrLf & Space(3) & .Value & Space(3) & vbCrLf
                TipText = tmpToolTip
                ShowTip = True
                Exit Sub
            End If
        End If
       
        .Col = 3:    If Trim(.Value) = "" Then Exit Sub
       
        If chkToolTip.Value = 0 Then GoTo Skip
       
        .Col = 10:  tmpToolTip = tmpToolTip & "  처    방 : " & .Value       '처방일
        .Col = 12:  tmpToolTip = tmpToolTip & "  by  " & GetDoctNm(.Value)                 '처방의
        .Col = 11:  tmpToolTip = tmpToolTip & " ( # " & Format(.Value, "##") & " )" & vbCrLf        '처방번호
        .Col = 13:  tmpToolTip = tmpToolTip & "  채    혈 : " & .Value       '채혈일시
        .Col = 14:  tmpColNm = GetDoctNm(.Value)       '채혈자-간호사
                    If Trim(tmpColNm) = "" Then
                        tmpColNm = GetEmpNm(.Value)    '채혈자-병리사
                    End If
                    tmpToolTip = tmpToolTip & "  by  " & tmpColNm & vbCrLf   '채혈자
        .Col = 15:  tmpToolTip = tmpToolTip & "  접    수 : " & .Value       '접수일시
        .Col = 16:  tmpToolTip = tmpToolTip & "  by  " & GetEmpNm(.Value) & vbCrLf   '접수자
        .Col = 4:
                    If .Value <> "미확" Then
                        .Col = 23:   tmpToolTip = tmpToolTip & "  결과보고 : " & .Value       '보고일시
                        .Col = 24:   tmpToolTip = tmpToolTip & "  by  " & GetEmpNm(.Value) & vbCrLf   '보고자
                    End If
        .Col = 17: pWorkArea = .Value
        .Col = 18: pAccDt = .Value
        .Col = 19: pAccSeq = .Value
            tmpToolTip = tmpToolTip & "  접수번호 : " & pWorkArea & "-" & pAccDt & "-" & pAccSeq & vbCrLf
       .Col = 20:
                    If .Value <> "" Then
                        tmpToolTip = tmpToolTip & vbCrLf & "  최근결과 : [ " & .Value & " ] " '& vbCrLf        '최근결과
                        '.Col = 21:   tmpToolTip = tmpToolTip & "             " & .Value  '최근결과일시
                        .Col = 21
                        tmpToolTip = tmpToolTip & Mid(.Value, 1, 9) '최근결과일시
                        .Col = 22
                        tmpToolTip = tmpToolTip & "  by " & GetEmpNm(.Value) & vbCrLf  '최근결과 보고자
                    End If
       '수정전 결과...
       .Col = 38:
                    Dim strModRst As String
'                    Dim pWorkArea As String
'                    Dim pAccDt As String
'                    Dim pAccSeq As String
                  
'                    .Col = 17: pWorkArea = .Value
'                    .Col = 18: pAccDt = .Value
'                    .Col = 19: pAccSeq = .Value
                  
                    .Col = 27
                    strSQL = MySql.SqlGetOldResult(pWorkArea, pAccDt, pAccSeq, .Value)
                    Set rsMod = New Recordset
                    rsMod.Open strSQL, DBConn
                    If Not rsMod.EOF Then
                        tmpToolTip = tmpToolTip & vbCrLf & "  수정전결과 : " & vbCrLf
                        'While (Not rsMod.EOF)
                            strModRst = "             [ " & Trim(rsMod.Fields("RstCd").Value) & " ] "
                            strModRst = strModRst & Format(Mid(rsMod.Fields("vfydt").Value, 3, 6), "0#-##-##") & Space(2)
                            strModRst = strModRst & " by " & rsMod.Fields("EmpNm").Value & vbCrLf
                            tmpToolTip = tmpToolTip & strModRst
                        '    rsMod.MoveNext
                        'Wend
                    End If
                    Set rsMod = Nothing
       
Skip:
        If UBound(aryMesg) >= Row Then
           If aryMesg(Row) <> "" Then tmpToolTip = tmpToolTip & vbCrLf & "  " & aryMesg(Row) & vbCrLf
        End If
     
Skip1:
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 5000
        .TextTipDelay = 500
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        If chkToolTip.Value = 1 Then
            ShowTip = True
        Else
            ShowTip = False
        End If
       
    End With
End Sub

Private Sub DisplayLISResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long, _
                          ByVal pTestDiv As String, Optional pQuery As Boolean = True)
   
    Dim I As Integer, J As Integer
    Dim MyResult As New clsLISResultReview
    Dim ResultBuffer As String
    Dim RstTxtBuffer As String
    Dim SamTxtBuffer As String
       
    With MyResult
      
        MouseRunning
        
        Call .ResultMore(pWorkArea, pAccDt, pAccSeq, pTestDiv)
      
        If .ResultCnt = 0 Then
            MouseDefault
            Exit Sub
        End If
      
        lblDeptNm.Caption = .DeptNm
      
        For I = 1 To .RstRow
            tblResult.Row = I + .OffSet
            For J = 1 To 8
                tblResult.Col = J
                If .Get_ForeColor(J, I) <> 0 Then tblResult.ForeColor = .Get_ForeColor(J, I)
            Next
        Next
      
        '결과내역 Display
        tblResult.Row = 1
        tblResult.Row2 = tblResult.MaxRows
        tblResult.Col = 2
        tblResult.Col2 = tblResult.MaxCols
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.Clip = .ResultClipText    '& .SenClipText             'ResultBuffer
        tblResult.BlockMode = False
      
        '미생물 감수성 결과의 경우 항생제명 순으로 Sort / Align Left
        'If .SortFg Then
        If .SortFg Then
            For I = 1 To .SensiCount
                tblResult.SortBy = SortByRow
                tblResult.SortKey(1) = 2  '항생제명
                tblResult.SortKeyOrder(1) = SortKeyOrderAscending
                tblResult.Col = -1
                tblResult.Row = .AntiSortStartRow(I)   '+ .OffSet
                tblResult.Row2 = .AntiSortEndRow(I)    '+ .OffSet
                tblResult.Action = ActionSort
                tblResult.Row = .SortStartRow - 1 '+ .OffSet
                tblResult.Col = 2
                tblResult.FontUnderline = True
            Next
        Else
            tblResult.Col = 6
            tblResult.Row = -1
            tblResult.ForeColor = DCM_LightRed
            tblResult.FontBold = True
        End If
        If .TestDiv = TST_MicTest Then
            '미생물 결과 : 균명컬럼 Align Left
            tblResult.Row = -1
            tblResult.Col = -1
            tblResult.BlockMode = True
            tblResult.AllowCellOverflow = True
            tblResult.TypeHAlign = TypeHAlignLeft
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 17
            
            For I = 1 To 5
                If .MicFg(I) Then
                    tblResult.ColWidth(I + 2) = 9
                Else
                    tblResult.ColWidth(I + 2) = 4
                End If
            Next
            
            tblResult.ColWidth(8) = 20
            tblResult.Col = 3: tblResult.Col2 = 7
            tblResult.Row = -1
            tblResult.BlockMode = True
            tblResult.FontBold = False
            tblResult.BlockMode = False
        Else
            '일반결과 : 결과컬럼 Align Center
            tblResult.Row = 1: tblResult.Row2 = tblResult.MaxRows
            tblResult.Col = 3: tblResult.Col2 = 7
            tblResult.BlockMode = True
            tblResult.TypeHAlign = TypeHAlignCenter
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 13
            tblResult.ColWidth(3) = 9
            tblResult.ColWidth(4) = 9
            tblResult.ColWidth(5) = 3
            tblResult.ColWidth(6) = 5
            tblResult.ColWidth(7) = 13
        End If
      
        '검체리마크 & 풋노트 Display
        If .CommentFg Then
            txtSamCmt.Text = .SamTextBuffer
            chkSamCmt.Value = 1
            chkSamCmt.Enabled = True
            Call HighlightText(txtSamCmt, "<< Remark >>", True)
            Call HighlightText(txtSamCmt, "<< Foot Note >>", False)
        Else
            chkSamCmt.Value = 0
            chkSamCmt.Enabled = False
        End If
      
        '특수검사 결과 Display
        If .SpecialFg Then
            rtfResult.TextRTF = .SpeTextBuffer
            rtfResult.Tag = rtfResult.Tag & COL_DIV & .SpeRstTitle
            Call rtfResult_DblClick
        End If
        
        '미생물결과 Or FootNote가 있을경우에만 표시
        If pTestDiv = "2" Or .CommentFg Then
            frmPop.Visible = True
        End If
    End With
   
   
    With tblResult
        .Col = 2: .Col2 = 5 '.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        txtRstCmt1.Text = .Clip
        .BlockMode = False
    End With
    Call HighlightText(txtRstCmt1, "<< 검사 소견 >>", True)
    Call HighlightText(txtRstCmt1, "<< Supplemental Report >>", False)
    Call HighlightText(txtRstCmt1, "[ Susceptibility test ]", False)
    Call HighlightText(txtRstCmt1, "Antibiotics", False, , &HDF6A3E)
    Call HighlightText(txtRstCmt1, "1      ", False, , &HDF6A3E)
    Call HighlightText(txtRstCmt1, "2      ", False, , &HDF6A3E)
    Call HighlightText(txtRstCmt1, "3      ", False, , &HDF6A3E)
   
    MouseDefault
End Sub

Private Sub rtfResult_DblClick()
    Dim sLabNo()
    Dim strTag As String
    Dim strLabNo As String
    Dim aryLabNo As Variant
    
    Screen.MousePointer = vbArrowHourglass
    DoEvents
    
    strTag = rtfResult.Tag
    strLabNo = medGetP(strTag, 1, COL_DIV)
    aryLabNo = Split(strLabNo, "-")
    
    frmAPS905.rtfResultText.Visible = True
    frmAPS905.OrdDiv = "L"
    frmAPS905.Caption = medGetP(strTag, 2, COL_DIV)
    frmAPS905.rtfResultText.TextRTF = rtfResult.TextRTF
    Screen.MousePointer = vbDefault

    DoEvents
    
    frmAPS905.WindowState = 0
    frmAPS905.Show vbModal
End Sub

Private Sub ResultClear()
    txtSamCmt.Text = ""
    txtRstCmt1.Text = ""
    lblSpecimenNm.Caption = ""
    rtfResult.Text = ""
            
'    fraLisResult.Visible = True
'    tblResult.Visible = True
'    picFootNote.Visible = True
'    rtfResult.Visible = False
   
    With tblResult
        '결과테이블 Clear
        .Row = -1:  .Col = -1
        .BlockMode = True
        .FontBold = False
        .Action = ActionClearText
        .ForeColor = &H747474
        .BlockMode = False
        
        '검사명/결과 컬럼 Bold
        .Row = -1: .Col = 2: .Col2 = 3
        .BlockMode = True
        .FontBold = True
        .BlockMode = False
        
        'High/Low field font 지정
        .Row = -1: .Col = 5: .Col2 = 5
        .BlockMode = True
        .FontName = "돋움"
        .BlockMode = False
        .RowsFrozen = 0
    End With
End Sub

Private Sub txtWorkCd_GotFocus()
    Call FocusMe(Me.txtWorkCd)
End Sub

Private Sub txtWorkCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstWSCode.ListCount = 0 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDown:
            lstWSCode.Visible = True
            lstWSCode.ListIndex = 0
            lstWSCode.ZOrder 0
            lstWSCode.SetFocus
        Case vbKeyEscape:
            lstWSCode.Visible = False
    End Select
End Sub

Private Sub txtWorkCd_KeyPress(KeyAscii As Integer)
    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    
    If KeyAscii = vbKeyReturn Then
        Call lstWSCode_KeyDown(vbKeyReturn, 0)
        lstWSCode.Visible = False
        Call cmdQuery_Click
        Exit Sub
    ElseIf KeyAscii = vbKeyEscape Then
        lstWSCode.Visible = False
        Exit Sub
    End If
    
    If lstWSCode.ListCount > 0 Then
        lstWSCode.Visible = True
        lstWSCode.ZOrder 0
        Call medCodeHelp(KeyAscii, lstWSCode, txtWorkCd.Text, txtWorkCd, dtpDate1)
    End If
End Sub

Private Sub txtWorkCd_Validate(Cancel As Boolean)
    Dim strWorkCd As String
    
    strWorkCd = medGetP(lstWSCode.Text, 1, vbTab)
    If txtWorkCd.Text = "" Then Exit Sub

    If Trim(strWorkCd) = "" Then
        Cancel = True
        txtWorkCd.SetFocus
        Exit Sub
    End If
    
    Set objLab301 = New clsWSBuild
    If objLab301.IsWorkCd(strWorkCd) = False Then
        MsgBox "코드 입력 Error!", vbCritical
        Cancel = True
        txtWorkCd.SetFocus
        Exit Sub
    End If
    
    Set objLab032 = New clsComcode032
    With objLab032
        .LoadTable LC3_WorkSheetName, , strWorkCd
        .MoveFirst
        If .RecordCount > 0 Then
            If Trim(ObjSysInfo.BuildingCd) <> Trim(.Field2) Then
                MsgBox "현재 건물에서는 사용할 수 없는 코드입니다.", vbCritical
                Cancel = True
                Call FocusMe(Me.txtWorkCd)
                Exit Sub
            End If
        End If
    End With
    
    Set objLab032 = Nothing
    lstWSCode.Visible = False
End Sub

Public Sub FocusMe(ctlName As Control)
    With ctlName
        .SelStart = 0
        .SelLength = Len(ctlName)
    End With
End Sub
