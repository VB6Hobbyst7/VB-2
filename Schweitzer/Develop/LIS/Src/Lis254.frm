VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm253MReading 
   BackColor       =   &H00DBE6E6&
   Caption         =   "배양 양성자 출력"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14670
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14670
   Tag             =   "25200"
   WindowState     =   2  '최대화
   Begin VB.ComboBox cboMonth 
      BackColor       =   &H00F1F5F4&
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
      ItemData        =   "Lis254.frx":0000
      Left            =   5715
      List            =   "Lis254.frx":0002
      Style           =   2  '드롭다운 목록
      TabIndex        =   50
      Top             =   390
      Width           =   840
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BackColor       =   &H00FFF2EC&
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   180
      Left            =   11865
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "☞ Bold : 이전결과가 Nogrowth"
      Top             =   900
      Width           =   2580
   End
   Begin VB.CheckBox chkStain 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Stain Worksheet"
      ForeColor       =   &H005B679D&
      Height          =   300
      Left            =   10965
      TabIndex        =   39
      Top             =   60
      Width           =   1770
   End
   Begin VB.PictureBox picAddSpc 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   2580
      ScaleHeight     =   2355
      ScaleWidth      =   4575
      TabIndex        =   22
      Top             =   2970
      Visible         =   0   'False
      Width           =   4575
      Begin DRcontrol1.DrFrame fraAddSpc 
         Height          =   2340
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   4128
         Title           =   "= 추가할 검체의 Lab No.를 입력하세요. ="
         BackColor       =   16776439
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton cmdCancel 
            Caption         =   "취소"
            Height          =   345
            Left            =   2595
            TabIndex        =   28
            Top             =   1785
            Width           =   750
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "추가"
            Height          =   345
            Left            =   3360
            Style           =   1  '그래픽
            TabIndex        =   27
            Top             =   1785
            Width           =   780
         End
         Begin VB.TextBox txtAccSeq 
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BorderStyle     =   0  '없음
            Height          =   225
            Left            =   3285
            MaxLength       =   5
            TabIndex        =   26
            Text            =   "10013"
            Top             =   645
            Width           =   615
         End
         Begin VB.TextBox txtAccDt 
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BorderStyle     =   0  '없음
            Height          =   240
            Left            =   2460
            MaxLength       =   4
            TabIndex        =   25
            Text            =   "9906"
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox txtWorkArea 
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BorderStyle     =   0  '없음
            Height          =   225
            Left            =   1755
            MaxLength       =   2
            TabIndex        =   24
            Text            =   "41"
            Top             =   645
            Width           =   375
         End
         Begin MedControls1.LisLabel LisLabel11 
            Height          =   360
            Left            =   315
            TabIndex        =   29
            Top             =   570
            Width           =   1290
            _ExtentX        =   2275
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
            Caption         =   "접수 번호"
            Appearance      =   0
         End
         Begin VB.Label lblRstCd 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   1755
            TabIndex        =   44
            Top             =   1875
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblSpcCd 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   2925
            TabIndex        =   43
            Top             =   1515
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lblTestFg 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   1005
            TabIndex        =   38
            Top             =   1890
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblTestCd 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   300
            TabIndex        =   37
            Top             =   1905
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblTestNm 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   240
            Left            =   315
            TabIndex        =   36
            Top             =   1560
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.Label lblSpcNm 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   2910
            TabIndex        =   35
            Top             =   1215
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lblAgeSex 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   2175
            TabIndex        =   34
            Top             =   1215
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblPtNm 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   1170
            TabIndex        =   33
            Top             =   1215
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblPtId 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   315
            TabIndex        =   32
            Top             =   1215
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   31
            Top             =   645
            Width           =   195
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   30
            Top             =   645
            Width           =   195
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00F1F5F4&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            Height          =   360
            Index           =   1
            Left            =   1620
            Shape           =   4  '둥근 사각형
            Top             =   570
            Width           =   2460
         End
      End
   End
   Begin VB.CommandButton cmdDelWorksheet 
      BackColor       =   &H00F8F8FE&
      Caption         =   "Worksheet 삭제"
      Height          =   480
      Left            =   2100
      Style           =   1  '그래픽
      TabIndex        =   21
      Tag             =   "25206"
      Top             =   8535
      Width           =   1920
   End
   Begin VB.CommandButton cmdAddWorksheet 
      BackColor       =   &H00FFF9F7&
      Caption         =   "Worksheet 추가"
      Height          =   480
      Left            =   75
      Style           =   1  '그래픽
      TabIndex        =   20
      Tag             =   "25206"
      Top             =   8535
      Width           =   1920
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출  력"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   18
      Tag             =   "25206"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "25206"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdWSList 
      BackColor       =   &H00DEDBDD&
      Caption         =   "▼"
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
      Left            =   10575
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   45
      Width           =   345
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00DBE6E6&
      Height          =   7230
      Left            =   75
      TabIndex        =   2
      Top             =   1275
      Width           =   14400
      Begin VB.CommandButton cmdEx1 
         BackColor       =   &H00CDE7FA&
         Caption         =   ">>"
         Height          =   350
         Left            =   9465
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   2670
         Width           =   550
      End
      Begin VB.CommandButton cmdIn1 
         BackColor       =   &H00CDE7FA&
         Caption         =   "<<"
         Height          =   350
         Left            =   9465
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   3150
         Width           =   550
      End
      Begin FPSpread.vaSpread ssTable 
         Height          =   6525
         Left            =   180
         TabIndex        =   4
         Tag             =   "25211"
         Top             =   600
         Width           =   9255
         _Version        =   196608
         _ExtentX        =   16325
         _ExtentY        =   11509
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   14
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis254.frx":0004
         UserResize      =   0
         VisibleCols     =   6
         VisibleRows     =   500
         TextTip         =   2
      End
      Begin FPSpread.vaSpread ssHTable 
         Height          =   6495
         Left            =   10050
         TabIndex        =   19
         Tag             =   "25211"
         Top             =   600
         Width           =   4095
         _Version        =   196608
         _ExtentX        =   7223
         _ExtentY        =   11456
         _StockProps     =   64
         AutoCalc        =   0   'False
         DisplayColHeaders=   0   'False
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   14
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis254.frx":1FAD
         UserResize      =   0
         VisibleCols     =   6
         VisibleRows     =   500
      End
      Begin VB.Label lblWarnCnt 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   13650
         TabIndex        =   41
         Top             =   330
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblHCount 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   13605
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblCount 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   8565
         TabIndex        =   12
         Top             =   315
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00DBE6E6&
         Caption         =   "◈  배양 양성자 리스트"
         Height          =   225
         Left            =   10065
         TabIndex        =   11
         Top             =   300
         Width           =   2235
      End
      Begin VB.Label Label9 
         BackColor       =   &H00DBE6E6&
         Caption         =   "◈  배양 대상 리스트"
         Height          =   225
         Left            =   300
         TabIndex        =   10
         Top             =   315
         Width           =   2505
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  '점
         X1              =   9735
         X2              =   9735
         Y1              =   525
         Y2              =   6960
      End
   End
   Begin VB.TextBox txtWSUnit 
      Alignment       =   2  '가운데 맞춤
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
      Left            =   7710
      TabIndex        =   1
      Text            =   "19990005"
      Top             =   45
      Width           =   2850
   End
   Begin VB.ComboBox cboWSCode 
      BackColor       =   &H00F1F5F4&
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
      ItemData        =   "Lis254.frx":3ECE
      Left            =   5715
      List            =   "Lis254.frx":3ED0
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   45
      Width           =   2010
   End
   Begin VB.ListBox lstWSUnit 
      BackColor       =   &H00F7FFF7&
      Height          =   2220
      Left            =   7725
      TabIndex        =   9
      Top             =   405
      Visible         =   0   'False
      Width           =   2850
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   3810
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   45
      Width           =   1875
      _ExtentX        =   3307
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
      Caption         =   "Work Sheet Unit"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   810
      Width           =   1665
      _ExtentX        =   2937
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
      Caption         =   "접수 마감일/시"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   2
      Left            =   195
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   810
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "총 검체수"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   3
      Left            =   6180
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   810
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "Worksheet 작성일/시"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   4
      Left            =   3810
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   390
      Width           =   1875
      _ExtentX        =   3307
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
      Caption         =   "조회기간"
      Appearance      =   0
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "WARNING"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   10500
      TabIndex        =   40
      ToolTipText     =   "선택된 환자의 이전 결과는 NoGrowth입니다."
      Top             =   885
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Shape shpWarning 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Left            =   10440
      Shape           =   4  '둥근 사각형
      Top             =   795
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblTCount 
      BackStyle       =   0  '투명
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   240
      Left            =   1575
      TabIndex        =   16
      Top             =   855
      Width           =   555
   End
   Begin VB.Label lblRcvDT 
      BackStyle       =   0  '투명
      Caption         =   "Feb 03 1999 10:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   255
      Left            =   4005
      TabIndex        =   15
      Top             =   855
      Width           =   1935
   End
   Begin VB.Label lblBltDate 
      BackStyle       =   0  '투명
      Caption         =   "Feb 03 1999 10:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   225
      Left            =   8310
      TabIndex        =   14
      Top             =   855
      Width           =   1965
   End
   Begin VB.Label lblInsResult 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Insert Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   405
      TabIndex        =   7
      Tag             =   "25205"
      Top             =   7965
      Width           =   1620
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F1F5F5&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      Height          =   525
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   720
      Width           =   14415
   End
End
Attribute VB_Name = "frm253MReading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objRstDic As New clsDictionary
Private objMicRst As New clsLISMicResult
Private objMicCul As New clsLISMicCulture
Private objMicLib As New clsLISMicroLib

Private fWorkSheet() As tpMicWorkSheet
Private fNGCode() As String

Private Const fSCItem = &H8080FF          ' Worksheet List 에서 선택된 Lab-No
Private fGCItem As Long

Private Sub cboWSCode_Click()
    
    Dim i As Integer
    
    If cboWSCode.ListIndex < 0 Then Exit Sub
    
    txtWSUnit = ""
    lstWSUnit.Clear
    lstWSUnit.Visible = False
    txtWSUnit.SetFocus

    lblTCount = "": lblRcvDT = "": lblBltDate = ""
    lblCount = "": ssTable.MaxRows = 0
    picAddSpc.Visible = False
    
End Sub

Private Sub chkStain_Click()
    
    If chkStain.Value = 0 Then
        objMicRst.LoadWorksheetCode MWS_ForCulture, cboWSCode, fWorkSheet
    Else
        objMicRst.LoadWorksheetCode MWS_ForStain, cboWSCode, fWorkSheet
    End If
    cboWSCode.ListIndex = -1
    txtWSUnit.Text = ""
    ssHTable.MaxRows = 0
    ScreenClear
    
End Sub

Private Sub cmdAdd_Click()
    
    Dim objMicWS As New clsLISMicWorksheet
    Dim blnAdd As Boolean
    Dim iWSIndex As Integer
    Dim tmpAccDt As String
    
    If Trim(txtWorkArea.Text) = "" Then Exit Sub
    If Trim(txtAccDt.Text) = "" Then Exit Sub
    If Trim(txtAccSeq.Text) = "" Then Exit Sub
    
    tmpAccDt = IIf(Mid(txtAccDt.Text, 1, 1) = "9", "19" & txtAccDt.Text, "20" & txtAccDt.Text)
    
    
    iWSIndex = cboWSCode.ListIndex
    blnAdd = objMicWS.SetWorksheetB(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text, _
                                    txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, _
                                    lblTestFg.Caption)
    If Not blnAdd Then
        MsgBox "검체 추가시 오류가 발생했습니다.", vbInformation, "검체추가"
        Exit Sub
    End If
    
    blnAdd = objMicWS.SetStatus(txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, "'" & lblTestCd.Caption & "'")
    If Not blnAdd Then
        MsgBox "검체 추가시 오류가 발생했습니다.", vbInformation, "검체추가"
        Exit Sub
    End If
    
    With objMicLib
        .WorkArea = txtWorkArea.Text
        .AccDt = Mid(Format(GetSystemDate, "yyyyMMdd"), 1, 2) & txtAccDt.Text
        .AccSeq = txtAccSeq.Text
        .PtId = lblPtId.Caption
        .TestCd = "'" & lblTestCd.Caption & "'"
        .SpcCd = lblSpcCd.Caption
    End With
    
    With ssTable
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        'ColNo.cnAccNo=1
        .Col = ColNo.cnAccNo:  .Text = txtWorkArea.Text & "-" & txtAccDt.Text & "-" & txtAccSeq.Text
        .Col = ColNo.cnPtid:   .Text = lblPtId.Caption
        .Col = ColNo.cnPtNm:   .Text = lblPtNm.Caption
        .Col = ColNo.cnSA:     .Text = lblAgeSex.Caption
        .Col = ColNo.cnSpcNm:  .Text = lblSpcNm.Caption
        .Col = ColNo.cnLstRst: .Text = objMicLib.GetNoGrowthLatestRst
        .Col = ColNo.cnCurRst: .Text = objMicLib.GetNoGrowthRst(lblRstCd.Caption)
        .Col = ColNo.cnMic:    .Text = lblTestFg.Caption
        .Col = ColNo.cnTestCd: .Text = lblTestCd.Caption
        .Col = ColNo.cnWsCd:   .Text = fWorkSheet(iWSIndex).WsCode
        .Col = ColNo.cnWsUnit: .Text = txtWSUnit.Text
        .Col = ColNo.cnHold:
        .Col = ColNo.cnSpcCd:  .Text = lblSpcCd.Caption
        .Col = ColNo.cnWarn:
        lblCount.Caption = .MaxRows
    End With
    
    txtAccSeq.Text = ""
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblAgeSex.Caption = ""
    lblSpcNm.Caption = ""
    lblTestNm.Caption = ""
    lblTestCd.Caption = ""
    
    txtAccSeq.SetFocus
    
'    picAddSpc.Visible = False
'    fraMain.Enabled = True

End Sub

Private Sub cmdAddWorksheet_Click()
    
    Dim strWorkArea As String
    Dim strAccDt As String
    
    ssTable.Row = 1
    ssTable.Col = ColNo.cnAccNo
    strWorkArea = medGetP(ssTable.Value, 1, "-")
    strAccDt = medGetP(ssTable.Value, 2, "-")
    
    txtWorkArea.Text = IIf(Trim(strWorkArea) = "", MIC_WorkArea, strWorkArea)
    txtAccDt.Text = IIf(Trim(strAccDt) = "", Format(Now, "YYMM"), strAccDt)
    
    txtAccSeq.Text = ""
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblAgeSex.Caption = ""
    lblSpcNm.Caption = ""
    lblTestNm.Caption = ""
    lblTestCd.Caption = ""
    
    picAddSpc.Visible = True
    picAddSpc.ZOrder 0
    
    fraMain.Enabled = False
    
    txtAccSeq.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    picAddSpc.Visible = False
    fraMain.Enabled = True
End Sub

Private Sub cmdClear_Click()
    cboWSCode.ListIndex = -1:
    txtWSUnit = ""
    ssHTable.MaxRows = 0
    Call ScreenClear
    
    shpWarning.Visible = False
    lblWarning.Visible = False
End Sub

Private Sub cmdDelWorksheet_Click()
    
    Dim Resp As VbMsgBoxResult
    Dim objMicWS As New clsLISMicWorksheet
    Dim strDelList As String
    Dim blnDel  As Boolean
    Dim i As Long
    Dim iWSIndex As Integer
    
    If picAddSpc.Visible Then Exit Sub
    
    Resp = MsgBox("선택된 검체를 해당 Worksheet에서 삭제하시겠습니까?", vbQuestion + vbYesNo, "검체삭제")
    If Resp = vbNo Then Exit Sub
    
    MouseRunning
    
    iWSIndex = cboWSCode.ListIndex
    
    strDelList = ""
    With ssTable
        For i = 1 To .MaxRows
            .Row = i
            .Col = ColNo.cnCOL0
            If .Text = "X" Then
                .Col = ColNo.cnAccNo
                strDelList = .Text & COL_DIV
            End If
        Next
    End With
    
    If Trim(strDelList) = "" Then
        MsgBox "선택된 검체가 없습니다.", vbInformation, "검체삭제"
        Exit Sub
    End If
    
    blnDel = objMicWS.DelSpcFROMWorksheet(strDelList, fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text)
    
    MouseDefault
    
    If Not blnDel Then
        MsgBox "검체삭제시 오류가 발생했습니다.", vbExclamation, "오류"
        Exit Sub
    End If
        
    With ssTable
        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = ColNo.cnCOL0
            If .Text = "X" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
            End If
        Next
        lblCount.Caption = .MaxRows
    End With
    
End Sub

Private Sub cmdPrint_Click()
    
    If Printers.Count = 0 Then
        MsgBox "출력할 프린터가 설정되지 않았습니다.", vbExclamation, "프린터"
        Exit Sub
    End If
    
    Dim MyReport As clsWorkListM
    Dim pParas As String
    Dim i As Integer
    Dim svWsCd As String, svWsUnit As String
    
    If ssHTable.MaxRows <= 0 Then Exit Sub
    
    pParas = "": svWsCd = "": svWsUnit = ""
    For i = 1 To ssHTable.MaxRows
        ssHTable.Row = i
        ssHTable.Col = ColNo.cnWsCd
        If svWsCd <> ssHTable.Value Then
            pParas = pParas & ssHTable.Value & "-"
            svWsCd = ssHTable.Value
            ssHTable.Col = ColNo.cnWsUnit
            pParas = pParas & ssHTable.Value & ";"
            svWsUnit = ssHTable.Value
        Else
            ssHTable.Col = ColNo.cnWsUnit
            If svWsUnit <> ssHTable.Value Then
                ssHTable.Col = ColNo.cnWsCd
                pParas = pParas & ssHTable.Value & "-"
                ssHTable.Col = ColNo.cnWsUnit
                pParas = pParas & ssHTable.Value & ";"
                svWsUnit = ssHTable.Value
            End If
        End If
            
    Next
    
    'MyReport.WS2Keys = pParas

    Dim strParas As String, strTmp As String
    Dim strSpcGrp As String, strWsUnit As String
    
    '2000.08.08 추가 : Nogrowth Batch등록에서 보류리스트의 Worksheet을 출력할 경우...
    strParas = pParas
    strTmp = medShift(strParas, ";")
    While (Trim(strTmp) <> "")
        strSpcGrp = medGetP(strTmp, 1, "-")
        strWsUnit = medGetP(strTmp, 2, "-")
        
        Set MyReport = New clsWorkListM
        '수정된 로직 Modify By legends 2003/08/08
        MyReport.Worksheet2 = False
        '
        
        If chkStain.Value = 0 Then
'            MyReport.Worksheet2 = True
            MyReport.StainWorksheet = False
        Else
'            MyReport.Worksheet2 = False
            MyReport.StainWorksheet = True
        End If
        Call MyReport.GetInputData(strSpcGrp, strWsUnit, "")
        Call MyReport.PrintReport
        Set MyReport = Nothing
        
        strTmp = medShift(strParas, ";")
    Wend
        
End Sub

Private Sub Form_Load()

    ssTable.Row = 1: ssTable.Col = 1: fGCItem = ssTable.ForeColor
    
    chkStain.Value = 0
    objMicRst.LoadWorksheetCode MWS_ForCulture, cboWSCode, fWorkSheet
    'objMicRst.LoadWorkSheetCode MWS_ForAll, cboWSCode, fWorkSheet
    
    cboWSCode.ListIndex = -1: Erase fNGCode
    txtWSUnit = ""
    ssHTable.MaxRows = 0
    ScreenClear
    
    cboMonth.Clear
    cboMonth.AddItem "1"
    cboMonth.AddItem "2"
    cboMonth.AddItem "3"
    cboMonth.AddItem "4"
    cboMonth.AddItem "5"
    cboMonth.AddItem "6"
    cboMonth.ListIndex = 0
End Sub

Private Sub ScreenClear()

    lblTCount = "": lblRcvDT = "": lblBltDate = ""
    lblCount = "": ssTable.MaxRows = 0
    picAddSpc.Visible = False
    fraMain.Enabled = True
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    Set objRstDic = Nothing
    Set objMicRst = Nothing
    Set objMicCul = Nothing
    Set frm253MReading = Nothing

End Sub

Private Sub cmdWSList_Click()
   
    Dim sWsCd As String
    Dim sMonth As String

    If cboWSCode.ListIndex < 0 Then Exit Sub

    sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
    sMonth = cboMonth.Text
    
'    objMicRst.LoadMicWorkList sWsCd, lstWSUnit
    objMicRst.LoadMicWorkList_New sWsCd, sMonth, lstWSUnit
    If lstWSUnit.ListCount <= 0 Then Exit Sub
    
    lstWSUnit.ListIndex = 0
    lstWSUnit.Visible = True
    lstWSUnit.ZOrder
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRstDic = Nothing
    Set objMicRst = Nothing
    Set objMicCul = Nothing
    Set objMicLib = Nothing
    Set frm253MReading = Nothing
End Sub

Private Sub lstWSUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim iListIndex As Integer, iWSIndex As Integer
    
    MouseRunning
    
    iWSIndex = cboWSCode.ListIndex
    iListIndex = lstWSUnit.ListIndex
    
    If Button = vbLeftButton And iListIndex >= 0 Then
        txtWSUnit.Text = medGetP(lstWSUnit.List(iListIndex), 1, " ")
        lstWSUnit.Clear
        lstWSUnit.Visible = False
        DoEvents
        Call DisplayData(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text, fWorkSheet(iWSIndex).WsRstType)
    End If
    
    lstWSUnit.Clear
    lstWSUnit.Visible = False
    
    MouseDefault
End Sub

Private Sub DisplayData(ByVal pWsCd As String, ByVal pWsUnit As String, ByVal sRTs As String)

    Dim strBuildDtTm As String, strRcvDtTm As String

    ScreenClear

    Call objMicRst.DispWorksheetInfo(pWsCd, pWsUnit, strBuildDtTm, strRcvDtTm)
    lblBltDate.Caption = strBuildDtTm
    lblRcvDT.Caption = strRcvDtTm
    
    lblCount.Caption = objMicCul.DispNogrowthList(ssTable, pWsCd, pWsUnit, sRTs)
    
    lblHCount.Caption = objMicCul.DispHoldingList(ssHTable, pWsCd, pWsUnit, sRTs, True)

    Call DispWarning
End Sub

Private Sub cmdIn1_Click()
    
    Dim i As Integer, sCnt As Integer
    ssHTable.Col = 1: sCnt = 0
    
    For i = ssHTable.MaxRows To 1 Step -1
        ssHTable.Row = i
        If ssHTable.ForeColor = fSCItem Then
            sCnt = sCnt + 1
            Call AddWorkSheet(ssHTable, i)
        End If
    Next i

    lblHCount = Val(lblHCount) - sCnt
    lblCount = Val(lblCount) + sCnt

End Sub
'
Private Sub ssHTable_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

    Call AddWorkSheet(ssHTable, Row)
    lblHCount = Val(lblHCount) - 1
    lblCount = Val(lblCount) + 1
    
    Call DispWarning
    
    If ssHTable.DataRowCnt = 0 Then
        shpWarning.Visible = False
        lblWarning.Visible = False
    End If
End Sub

Private Sub AddWorkSheet(ByVal pObj As Object, ByVal pRow As Integer)
    
    Dim sAccBuf As String

    ssTable.MaxRows = ssTable.MaxRows + 1
    
    pObj.Col = 1: pObj.COL2 = pObj.MaxCols
    pObj.Row = pRow: pObj.Row2 = pRow
    ssTable.Col = 1: ssTable.COL2 = ssTable.MaxCols
    ssTable.Row = ssTable.MaxRows: ssTable.Row2 = ssTable.MaxRows
    ssTable.Clip = pObj.Clip
    
    Call SaveOneRow(pRow, ssHTable, MWS_Ready)
    
    pObj.Row = pRow
    pObj.Action = ActionDeleteRow
    pObj.MaxRows = pObj.MaxRows - 1
    
End Sub

Private Sub ssTable_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim tmpcolor As Long
    
    If Col > 0 And Row = 0 Then
        
        ssTable.Col = -1: ssTable.Row = -1
        ssTable.ForeColor = fGCItem
        
        ssTable.SortBy = SortByRow
        ssTable.SortKey(1) = Col
        ssTable.SortKey(2) = 1
        ssTable.SortKeyOrder(1) = SortKeyOrderAscending
        ssTable.SortKeyOrder(2) = SortKeyOrderAscending
        ssTable.Col = 1
        ssTable.COL2 = ssTable.MaxCols
        ssTable.Row = 1
        ssTable.Row2 = ssTable.MaxRows
        ssTable.Action = ActionSort
        
    End If
    
    If Col > 0 And Row > 0 Then
    
        ssTable.Col = -1: ssTable.Row = Row
        tmpcolor = ssTable.ForeColor
        
        If tmpcolor = fSCItem Then
            ssTable.ForeColor = fGCItem
        Else
            ssTable.ForeColor = fSCItem
        End If
        
    End If
    
    If Col = 0 And Row > 0 Then
        ssTable.Row = Row
        ssTable.Col = ColNo.cnCOL0
        If ssTable.Text = "X" Then
            ssTable.Text = Row
        Else
            ssTable.Text = "X"
        End If
        ssTable.ForeColor = vbRed
    End If
    
End Sub

Private Sub cmdEx1_Click()
    
    Dim i As Integer, sCnt As Integer
    ssTable.Col = 1: sCnt = 0
    
    For i = ssTable.MaxRows To 1 Step -1
        ssTable.Row = i
        If ssTable.ForeColor = fSCItem Then
            sCnt = sCnt + 1
            MovetoETable i
        End If
    Next i

    lblCount = Val(lblCount) - sCnt
    lblHCount = Val(lblHCount) + sCnt
    
    Call DispWarning
End Sub


Private Sub ssTable_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

    MovetoETable Row
    lblCount = Val(lblCount) - 1
    lblHCount = Val(lblHCount) + 1
    
    Call DispWarning
End Sub

Private Sub MovetoETable(ByVal pRow As Integer)
    
    Dim sAccBuf As String
    Dim strKey1 As String
    Dim strKey2 As String
    Dim strKey3 As String
    Dim i As Long
    Dim varTestCd As Variant
    Dim varSpcCd As Variant

    shpWarning.Visible = False
    lblWarning.Visible = False
    
    
    Call ssTable.GetText(ColNo.cnSpcCd, pRow, varSpcCd)
    Call ssTable.GetText(ColNo.cnTestCd, pRow, varTestCd)
    
    ssHTable.MaxRows = ssHTable.MaxRows + 1
    
    ssTable.Col = 1: ssTable.Row = pRow
    strKey1 = ssTable.Text
    strKey2 = fWorkSheet(cboWSCode.ListIndex).WsCode
    strKey3 = txtWSUnit.Text
    
    With ssHTable
        For i = 1 To .MaxRows
            .Row = i
            .Col = ColNo.cnAccNo
            If .Value = strKey1 Then
                .Col = ColNo.cnWsCd
                If .Value = strKey2 Then
                    .Col = ColNo.cnWsUnit
                    If .Value = strKey3 Then Exit Sub
                End If
            End If
        Next
    End With
    
    ssTable.Col = 1: ssTable.COL2 = ssTable.MaxCols
    ssTable.Row = pRow: ssTable.Row2 = pRow
    ssHTable.Col = 1: ssHTable.COL2 = ssTable.MaxCols
    ssHTable.Row = ssHTable.MaxRows: ssHTable.Row2 = ssHTable.MaxRows
    ssHTable.Clip = ssTable.Clip
    ssHTable.Col = ColNo.cnWsCd     '검체군
    ssHTable.Value = fWorkSheet(cboWSCode.ListIndex).WsCode
    ssHTable.Col = ColNo.cnWsUnit    'Worksheet Unit
    ssHTable.Value = txtWSUnit.Text
    
    '보류리스트로 옮기는 동시에 Status도 Update한다.
    Call SaveOneRow(pRow, ssTable, MWS_Holding)
    
    ssTable.Row = pRow
    ssTable.Action = ActionDeleteRow
    ssTable.MaxRows = ssTable.MaxRows - 1
    
    With ssHTable
        .ReDraw = False
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
    
        .SortBy = SortByRow
        .SortKey(1) = ColNo.cnWsCd
        .SortKey(2) = ColNo.cnWsUnit
        .SortKey(3) = ColNo.cnAccNo
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending
        .SortKeyOrder(3) = SortKeyOrderAscending
        .Action = ActionSort
        .ReDraw = True
    End With
    
    Call GetWarningGrowth(strKey1, varTestCd, varSpcCd)
End Sub

Private Sub SaveOneRow(ByVal iRow As Long, ByVal ssObj As Object, ByVal pStatus As String)
        
    Dim sAccNo As String, sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim sWsCd As String
    
    'pStatus = MWS_Holding - 보류, MWS_Ready - Worksheet
               
    With ssObj
        
        .Col = ColNo.cnAccNo: .Row = iRow: sAccNo = .Text
        sWorkArea = medGetP(sAccNo, 1, "-")
        sAccDt = medGetP(sAccNo, 2, "-")
        sAccSeq = medGetP(sAccNo, 3, "-")
        sAccDt = IIf(Mid$(sAccDt, 1, 1) = "9", "19", "20") & sAccDt
        
        sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
        Call objMicCul.SaveOneStatus(sWsCd, txtWSUnit.Text, sWorkArea, sAccDt, sAccSeq, pStatus)
    
    End With

End Sub

Private Sub ssTable_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If Row = 0 Then Exit Sub
' barcode의 Col 이 17번인가 ? 아니 이게 왜 14번이지 ?
    With ssTable
        .Row = Row: .Col = 14
        If .Value = "" Then Exit Sub
        MultiLine = 1
        TipText = vbCRLF & "  " & .Text & vbCRLF
        TipWidth = 3000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("돋움체", 16, True, False, &HEEFDF2, DCM_Red)
        ShowTip = True
    End With
End Sub

Private Sub txtAccDt_Change()
    If Not picAddSpc.Visible Then Exit Sub
    If Len(txtAccDt.Text) = txtAccDt.MaxLength Then txtAccSeq.SetFocus
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtAccSeq.SetFocus
End Sub

Private Sub txtAccSeq_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call txtAccSeq_LostFocus
End Sub

Private Sub txtAccSeq_LostFocus()

    Dim objMicWS As New clsLISMicWorksheet
    Dim strSpcInfo As String
    Dim iWSIndex As Integer
    Dim tmpAccDt As String
    Dim chkG As String, chkS As String, sTestFlag As String
    
    If Screen.ActiveControl.Name = cmdCancel.Name Then Exit Sub
    
    tmpAccDt = IIf(Mid(txtAccDt.Text, 1, 1) = "9", "19" & txtAccDt.Text, "20" & txtAccDt.Text)
    
    iWSIndex = cboWSCode.ListIndex
    strSpcInfo = objMicWS.GetAddSpcInfo(txtWorkArea.Text, tmpAccDt, txtAccSeq.Text, _
                                        fWorkSheet(iWSIndex).WsCode, fWorkSheet(iWSIndex).WsRstType)
            
    If strSpcInfo = "" Then
        MsgBox "추가할 수 없는 검체입니다.", vbInformation, "검체추가"
    Else
        lblPtId.Caption = medGetP(strSpcInfo, 1, COL_DIV)
        lblPtNm.Caption = medGetP(strSpcInfo, 2, COL_DIV)
        lblAgeSex.Caption = medGetP(strSpcInfo, 3, COL_DIV) & " / " & _
                            medGetP(strSpcInfo, 4, COL_DIV) \ 365
        lblSpcNm.Caption = medGetP(strSpcInfo, 5, COL_DIV)
        lblTestCd.Caption = medGetP(strSpcInfo, 6, COL_DIV)
        lblTestNm.Caption = medGetP(strSpcInfo, 7, COL_DIV)
        lblTestFg.Caption = medGetP(strSpcInfo, 8, COL_DIV)
        lblSpcCd.Caption = medGetP(strSpcInfo, 10, COL_DIV)
        lblRstCd.Caption = medGetP(strSpcInfo, 11, COL_DIV)
    End If
    Set objMicWS = Nothing
    
End Sub

Private Sub txtWorkArea_Change()
    If Not picAddSpc.Visible Then Exit Sub
    If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then txtAccDt.SetFocus
End Sub

Private Sub txtWSUnit_KeyPress(KeyAscii As Integer)
    
    Dim iWSIndex As Integer
        
    If KeyAscii = vbKeyReturn Then
    
        iWSIndex = cboWSCode.ListIndex

        If ExistWS(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text) Then
            Call DisplayData(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text, fWorkSheet(iWSIndex).WsRstType)
        Else
            Call ScreenClear
        End If
        
    End If

End Sub

Private Sub DispWarning()
    Dim varWarning As Variant
    Dim Cnt As Long
    Dim i As Long
    
    For i = 1 To ssHTable.DataRowCnt
        Call ssHTable.GetText(ssHTable.MaxCols, i, varWarning)
        ssHTable.Col = -1: ssHTable.Row = i
        If varWarning = "1" Then
            ssHTable.ForeColor = vbActiveTitleBar
            ssHTable.FontBold = True
            Cnt = Cnt + 1
        Else
            ssHTable.ForeColor = vbBlack
            ssHTable.FontBold = False
        End If
    Next
    
    If Cnt > 0 Then
        lblWarnCnt.Visible = True: lblWarnCnt.Caption = Cnt
    Else
        lblWarnCnt.Visible = False: lblWarnCnt.Caption = ""
    End If
End Sub

Private Sub GetWarningGrowth(ByVal pAccNo As String, ByVal pTestCd As String, ByVal pSpcCd As String)
    Dim blnWarning As Boolean
    
    objMicLib.WorkArea = medGetP(pAccNo, 1, "-")
    objMicLib.AccDt = Mid(Format(GetSystemDate, "yyyyMMdd"), 1, 2) & medGetP(pAccNo, 2, "-")
    objMicLib.AccSeq = medGetP(pAccNo, 3, "-")
    objMicLib.TestCd = pTestCd
    objMicLib.SpcCd = pSpcCd
    
    If objMicLib.ChkInPatient = False Then
        blnWarning = False
    Else
        If objMicLib.GetNoGrowthLatestRst = "" Then
            blnWarning = False
        Else
            blnWarning = True
        End If
    End If
    
    If blnWarning Then
        shpWarning.Visible = True
        lblWarning.Visible = True
        
        Dim varAccNo As Variant
        Dim i As Long
        
        For i = 1 To ssHTable.DataRowCnt
            Call ssHTable.GetText(ColNo.cnAccNo, i, varAccNo)
            
            If varAccNo = pAccNo Then
                Call ssHTable.SetText(ssHTable.MaxCols, i, "1")
                Exit For
            End If
        Next
    Else
        shpWarning.Visible = False
        lblWarning.Visible = False
    End If
End Sub
