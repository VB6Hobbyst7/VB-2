VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm252MBatch 
   BackColor       =   &H00DBE6E6&
   Caption         =   "No Growth 배치결과등록"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14670
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
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
      ItemData        =   "Lis252.frx":0000
      Left            =   5775
      List            =   "Lis252.frx":0002
      Style           =   2  '드롭다운 목록
      TabIndex        =   65
      Top             =   390
      Width           =   1995
   End
   Begin VB.CommandButton cmdRPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   7815
      Style           =   1  '그래픽
      TabIndex        =   63
      Tag             =   "132"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00EBE0EB&
      Caption         =   "Worksheet 출력"
      Height          =   510
      Left            =   3240
      Style           =   1  '그래픽
      TabIndex        =   55
      Tag             =   "25206"
      Top             =   8535
      Width           =   1575
   End
   Begin VB.CheckBox chkStain 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Stain Worksheet"
      ForeColor       =   &H005B679D&
      Height          =   300
      Left            =   11070
      TabIndex        =   54
      Top             =   255
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
      TabIndex        =   20
      Top             =   2970
      Visible         =   0   'False
      Width           =   4575
      Begin DRcontrol1.DrFrame fraAddSpc 
         Height          =   2340
         Left            =   0
         TabIndex        =   21
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
         Begin VB.TextBox txtAccDt 
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BorderStyle     =   0  '없음
            Height          =   240
            Left            =   2460
            MaxLength       =   4
            TabIndex        =   26
            Text            =   "9906"
            Top             =   645
            Width           =   525
         End
         Begin VB.TextBox txtAccSeq 
            Appearance      =   0  '평면
            BackColor       =   &H00F1F5F4&
            BorderStyle     =   0  '없음
            Height          =   225
            Left            =   3270
            MaxLength       =   5
            TabIndex        =   28
            Text            =   "10011"
            Top             =   645
            Width           =   615
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "추가"
            Height          =   345
            Left            =   3360
            Style           =   1  '그래픽
            TabIndex        =   23
            Top             =   1785
            Width           =   780
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "취소"
            Height          =   345
            Left            =   2595
            TabIndex        =   22
            Top             =   1785
            Width           =   750
         End
         Begin MedControls1.LisLabel LisLabel11 
            Height          =   360
            Left            =   315
            TabIndex        =   25
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
            Left            =   1740
            TabIndex        =   58
            Top             =   1890
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblSpcCd 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   2955
            TabIndex        =   57
            Top             =   1515
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label7 
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
            TabIndex        =   36
            Top             =   645
            Width           =   195
         End
         Begin VB.Label Label6 
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
            TabIndex        =   35
            Top             =   645
            Width           =   195
         End
         Begin VB.Label lblPtId 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   315
            TabIndex        =   34
            Top             =   1215
            Visible         =   0   'False
            Width           =   825
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
         Begin VB.Label lblAgeSex 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   2175
            TabIndex        =   32
            Top             =   1215
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblSpcNm 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   2910
            TabIndex        =   31
            Top             =   1215
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lblTestNm 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   240
            Left            =   315
            TabIndex        =   30
            Top             =   1560
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.Label lblTestCd 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   300
            TabIndex        =   29
            Top             =   1905
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblTestFg 
            BackColor       =   &H00EFF5F8&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            ForeColor       =   &H00C76456&
            Height          =   255
            Left            =   1005
            TabIndex        =   27
            Top             =   1890
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00F1F5F4&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            Height          =   360
            Left            =   1620
            Shape           =   4  '둥근 사각형
            Top             =   570
            Width           =   2460
         End
      End
   End
   Begin VB.CommandButton cmdAddWorksheet 
      BackColor       =   &H00FFF9F7&
      Caption         =   "Worksheet 추가"
      Height          =   510
      Left            =   75
      Style           =   1  '그래픽
      TabIndex        =   19
      Tag             =   "25206"
      Top             =   8535
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelWorksheet 
      BackColor       =   &H00F8F8FE&
      Caption         =   "Worksheet 삭제"
      Height          =   510
      Left            =   1650
      Style           =   1  '그래픽
      TabIndex        =   18
      Tag             =   "25206"
      Top             =   8535
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   9165
      Style           =   1  '그래픽
      TabIndex        =   17
      Tag             =   "25206"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.ComboBox cboResult 
      BackColor       =   &H00FFF8EE&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   11235
      Style           =   2  '드롭다운 목록
      TabIndex        =   14
      Top             =   810
      Width           =   3075
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
      Left            =   10395
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
      TabIndex        =   5
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdFinEnter 
      BackColor       =   &H00F4F0F2&
      Caption         =   "최종 결과(&F)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "25207"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdMidEnter 
      BackColor       =   &H00F4F0F2&
      Caption         =   "중간 결과(&M)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "25206"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00DBE6E6&
      Height          =   7230
      Left            =   75
      TabIndex        =   2
      Top             =   1260
      Width           =   14385
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체선택(&A)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7455
         TabIndex        =   56
         Tag             =   "137"
         Top             =   300
         Width           =   1410
      End
      Begin VB.PictureBox picExtra 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Height          =   6705
         Left            =   9465
         ScaleHeight     =   6705
         ScaleWidth      =   4800
         TabIndex        =   37
         Top             =   465
         Width           =   4800
         Begin VB.CommandButton cmdIn1 
            BackColor       =   &H00CDE7FA&
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            Style           =   1  '그래픽
            TabIndex        =   43
            Top             =   5880
            Width           =   550
         End
         Begin VB.CommandButton cmdEx1 
            BackColor       =   &H00CDE7FA&
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            Style           =   1  '그래픽
            TabIndex        =   42
            Top             =   5520
            Width           =   550
         End
         Begin VB.CommandButton cmdIn2 
            BackColor       =   &H00CDE7FA&
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            Style           =   1  '그래픽
            TabIndex        =   41
            Top             =   3795
            Width           =   550
         End
         Begin VB.ListBox lstGList 
            BackColor       =   &H00F5F2FF&
            Height          =   2400
            Left            =   570
            TabIndex        =   40
            Top             =   405
            Width           =   1900
         End
         Begin VB.ListBox lstFList 
            BackColor       =   &H00F5FEFE&
            Height          =   2400
            Left            =   2730
            TabIndex        =   39
            Top             =   420
            Width           =   1900
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00F4F0F2&
            Caption         =   "Worksheet2"
            Height          =   300
            Left            =   2940
            Style           =   1  '그래픽
            TabIndex        =   38
            Top             =   2940
            Visible         =   0   'False
            Width           =   1290
         End
         Begin FPSpread.vaSpread ssETable 
            Height          =   1455
            Left            =   600
            TabIndex        =   44
            Tag             =   "25211"
            Top             =   5160
            Width           =   4095
            _Version        =   196608
            _ExtentX        =   7223
            _ExtentY        =   2566
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
            SpreadDesigner  =   "Lis252.frx":0004
            UserResize      =   0
            VisibleCols     =   6
            VisibleRows     =   500
         End
         Begin FPSpread.vaSpread ssHTable 
            Height          =   1455
            Left            =   600
            TabIndex        =   45
            Tag             =   "25211"
            Top             =   3300
            Width           =   4095
            _Version        =   196608
            _ExtentX        =   7223
            _ExtentY        =   2566
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
            SpreadDesigner  =   "Lis252.frx":1F70
            UserResize      =   0
            VisibleCols     =   6
            VisibleRows     =   500
         End
         Begin VB.Line Line2 
            BorderStyle     =   3  '점
            X1              =   270
            X2              =   270
            Y1              =   105
            Y2              =   6540
         End
         Begin VB.Label Label2 
            BackColor       =   &H00DBE6E6&
            Caption         =   "☞ Growth"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   570
            TabIndex        =   53
            Top             =   135
            Width           =   1080
         End
         Begin VB.Label Label10 
            BackColor       =   &H00DBE6E6&
            Caption         =   "☞ 보류 리스트"
            Height          =   225
            Left            =   585
            TabIndex        =   52
            Top             =   3045
            Width           =   2595
         End
         Begin VB.Label Label11 
            BackColor       =   &H00DBE6E6&
            Caption         =   "☞ 제외 리스트"
            Height          =   225
            Left            =   615
            TabIndex        =   51
            Top             =   4920
            Width           =   2235
         End
         Begin VB.Label lblGCount 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00DBE6E6&
            Caption         =   "000"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   1935
            TabIndex        =   50
            Top             =   165
            Width           =   495
         End
         Begin VB.Label lblHCount 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00DBE6E6&
            Caption         =   "000"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   4050
            TabIndex        =   49
            Top             =   3015
            Width           =   495
         End
         Begin VB.Label lblECount 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00DBE6E6&
            Caption         =   "000"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   4035
            TabIndex        =   48
            Top             =   4890
            Width           =   495
         End
         Begin VB.Label Label4 
            BackColor       =   &H00DBE6E6&
            Caption         =   "☞ Final"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2730
            TabIndex        =   47
            Top             =   150
            Width           =   1200
         End
         Begin VB.Label lblFCount 
            Alignment       =   1  '오른쪽 맞춤
            BackColor       =   &H00DBE6E6&
            Caption         =   "000"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   4080
            TabIndex        =   46
            Top             =   180
            Width           =   495
         End
      End
      Begin FPSpread.vaSpread ssTable 
         Height          =   6495
         Left            =   90
         TabIndex        =   6
         Tag             =   "25211"
         Top             =   615
         Width           =   9360
         _Version        =   196608
         _ExtentX        =   16510
         _ExtentY        =   11456
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         EditModePermanent=   -1  'True
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
         MaxCols         =   18
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis252.frx":3EC4
         UserResize      =   0
         VisibleCols     =   6
         VisibleRows     =   500
         TextTip         =   2
      End
      Begin VB.Label lblCount 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   8565
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "◈  배치 결과 등록 대상 리스트"
         Height          =   180
         Left            =   300
         TabIndex        =   10
         Top             =   315
         Width           =   2520
      End
   End
   Begin VB.TextBox txtWSUnit 
      Alignment       =   2  '가운데 맞춤
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
      Left            =   7755
      TabIndex        =   1
      Text            =   "19990005"
      Top             =   45
      Width           =   2625
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
      ItemData        =   "Lis252.frx":6038
      Left            =   5775
      List            =   "Lis252.frx":603A
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   45
      Width           =   1995
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   3810
      TabIndex        =   59
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
   Begin VB.ListBox lstWSUnit 
      BackColor       =   &H00FFF9F7&
      Height          =   2220
      ItemData        =   "Lis252.frx":603C
      Left            =   7740
      List            =   "Lis252.frx":603E
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   2625
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   825
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
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   825
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
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   825
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
      TabIndex        =   64
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
   Begin VB.Label lblTCount 
      BackStyle       =   0  '투명
      Caption         =   "000"
      ForeColor       =   &H00C76456&
      Height          =   210
      Left            =   1575
      TabIndex        =   16
      Top             =   900
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "결과"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C76456&
      Height          =   270
      Left            =   10530
      TabIndex        =   15
      Top             =   840
      Width           =   675
   End
   Begin VB.Label lblRcvDT 
      BackStyle       =   0  '투명
      Caption         =   "Feb 03 1999 10:00"
      ForeColor       =   &H00C76456&
      Height          =   180
      Left            =   4095
      TabIndex        =   13
      Top             =   915
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
      TabIndex        =   12
      Top             =   855
      Width           =   1965
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F1F5F5&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      Height          =   540
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   720
      Width           =   14325
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
End
Attribute VB_Name = "frm252MBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objRstDic As New clsDictionary
Private objMicRst As New clsLISMicResult
Private objMicCul As New clsLISMicCulture
Private objMicLib As New clsLISMicroLib '최근결과,Warning서식관련 클래스

Private fWorkSheet() As tpMicWorkSheet
Private fNGCode() As Variant

Private Const fSCItem = &H8080FF          ' Worksheet List 에서 선택된 Lab-No
Private fGCItem As Long

Private Sub cboWSCode_Click()
    
    Dim i As Integer
    
    If cboWSCode.ListIndex < 0 Then Exit Sub
    
    txtWSUnit = ""
    lstWSUnit.Clear
    lstWSUnit.Visible = False
    txtWSUnit.SetFocus

'    If chkStain.Value = 0 Then
        Call objMicRst.LoadNGRstCd(fWorkSheet(cboWSCode.ListIndex).WsType, cboResult, fNGCode)
        cboResult.ListIndex = -1
'    End If
    
    lblTCount = "": lblRcvDT = "": lblBltDate = ""
    
    lblGCount = "": lstGList.Clear
    lblFCount = "": lstFList.Clear
    lblCount = "":  ssTable.MaxRows = 0
    lblHCount = "": ssHTable.MaxRows = 0
    lblECount = "": ssETable.MaxRows = 0

End Sub

Private Sub chkSelAll_Click()
    ssTable.Col = ColNo.cnHold: ssTable.COL2 = ColNo.cnHold
    ssTable.Row = 1:  ssTable.Row2 = ssTable.DataRowCnt
    ssTable.BlockMode = True
    ssTable.Value = chkSelAll.Value
    ssTable.BlockMode = False
End Sub

Private Sub chkStain_Click()
    
    cboResult.Clear: Erase fNGCode
    If chkStain.Value = 0 Then
        Call objMicRst.LoadWorkSheetCode(MWS_ForCulture, cboWSCode, fWorkSheet)
        cboResult.Enabled = True
        picExtra.Enabled = True
        cmdMidEnter.Enabled = True
        cmdFinEnter.Enabled = True
    Else
        Call objMicRst.LoadWorkSheetCode(MWS_ForStain, cboWSCode, fWorkSheet)
        cboResult.Enabled = True
        picExtra.Enabled = True
        cmdMidEnter.Enabled = True
        cmdFinEnter.Enabled = True
    End If
    cboWSCode.ListIndex = -1
    txtWSUnit.Text = ""
    ScreenClear
    
End Sub

Private Sub cmdAdd_Click()
    
    Dim objMicWS As New clsLISMicWorksheet
    Dim blnAdd   As Boolean
    Dim iWSIndex As Integer
    Dim tmpAccDt As String
    
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
        .Col = ColNo.cnAccNo:   .Text = txtWorkArea.Text & "-" & txtAccDt.Text & "-" & txtAccSeq.Text
        .Col = ColNo.cnPtid:    .Text = lblPtId.Caption
        .Col = ColNo.cnPtNm:    .Text = lblPtNm.Caption
        .Col = ColNo.cnSA:      .Text = lblAgeSex.Caption
        .Col = ColNo.cnSpcNm:   .Text = lblSpcNm.Caption
        .Col = ColNo.cnLstRst:  .Text = objMicLib.GetNoGrowthLatestRst
        .Col = ColNo.cnCurRst:  .Text = objMicLib.GetNoGrowthRst(lblRstCd.Caption)
        .Col = ColNo.cnMic:     .Text = lblTestFg.Caption
        .Col = ColNo.cnTestCd:  .Text = lblTestCd.Caption
        .Col = ColNo.cnWsCd:    .Text = fWorkSheet(iWSIndex).WsCode
        .Col = ColNo.cnWsUnit:  .Text = txtWSUnit.Text
        .Col = ColNo.cnSpcCd:   .Text = lblSpcCd.Caption
        
        '-- 추가 By M.G.Choi
        .Col = 15:  .Text = objMicLib.mVfyDt
        .Col = 16:  .Text = objMicLib.mVfyTm
        .Col = 17:  .Text = objMicLib.mVfyID
        
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
End Sub

Private Sub cmdAddWorksheet_Click()
    
    Dim strWorkArea As String
    Dim strAccDt As String
    
    ssTable.Row = 1
    ssTable.Col = 1
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
    Call ScreenClear
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
            .Col = ColNo.cnHold
            If .Text = "1" Then
                .Col = ColNo.cnAccNo
                strDelList = .Text & COL_DIV
            End If
        Next
    End With
    
    If Trim(strDelList) = "" Then
        MsgBox "선택된 검체가 없습니다.", vbInformation, "검체삭제"
        Exit Sub
    End If
    
    blnDel = objMicWS.DelSpcFromWorksheet(strDelList, fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text)
    
    MouseDefault
    
    If Not blnDel Then
        MsgBox "검체삭제시 오류가 발생했습니다.", vbExclamation, "오류"
        Exit Sub
    End If
        
    With ssTable
        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = ColNo.cnHold
            If .Text = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
            End If
        Next
        lblCount.Caption = .MaxRows
    End With
    
End Sub

Private Sub cmdFinEnter_Click()
    Me.MousePointer = vbHourglass
    Call VerifyResult(enStsCd.StsCd_LIS_FinRst)
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
    
    Dim MyReport As New clsWorkListM
    
    If ssHTable.MaxRows <= 0 Then Exit Sub
        
    MyReport.Worksheet2 = True
    Call MyReport.GetInputData(fWorkSheet(cboWSCode.ListIndex).WsCode, txtWSUnit.Text, cboWSCode.Text)
    Call MyReport.PrintReport
    Set MyReport = Nothing
        
End Sub

Private Sub cmdReprint_Click()

    ' 클래스를 이용하여 출력
    Dim MyReport As New clsWorkListM
     
    If txtWSUnit.Text <> "" Then
        MyReport.Worksheet2 = False
        Call MyReport.GetInputData(fWorkSheet(cboWSCode.ListIndex).WsCode, txtWSUnit.Text, cboWSCode.Text)
        Call MyReport.PrintReport
        Set MyReport = Nothing
    End If

End Sub

Private Sub Form_Load()

    ssTable.Row = 1: ssTable.Col = 1: fGCItem = ssTable.ForeColor

    Call objMicRst.LoadWorkSheetCode(MWS_ForCulture, cboWSCode, fWorkSheet)
'    SetNGCode
    
    cboWSCode.ListIndex = -1: cboResult.Clear: Erase fNGCode
    txtWSUnit.Text = ""
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
    cboResult.ListIndex = -1
'    cboResult.Enabled = False
    
    lblCount = "": ssTable.MaxRows = 0
    lblGCount = "": lstGList.Clear
    lblFCount = "": lstFList.Clear
    lblHCount = "": ssHTable.MaxRows = 0
    lblECount = "": ssETable.MaxRows = 0
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    Set objRstDic = Nothing
    Set objMicRst = Nothing
    Set objMicCul = Nothing
    Set frm252MBatch = Nothing
    
End Sub

'### 조회조건 추가
'### 온승호
'### 2010년 5월 13일
Private Sub cmdWSList_Click()
        
    Dim sWsCd As String
    Dim sMonth As String

    If cboWSCode.ListIndex < 0 Then Exit Sub

    sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
    sMonth = cboMonth.Text
    
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
    Set objMicLib = Nothing '최근결과,Warning서식관련 클래스
End Sub

Private Sub lstWSUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim iListIndex As Integer, iWSIndex As Integer
    
    MouseRunning
    
    iWSIndex = cboWSCode.ListIndex
    iListIndex = lstWSUnit.ListIndex
    
    If Button = vbLeftButton And iListIndex >= 0 Then
        txtWSUnit.Text = medGetP(lstWSUnit.List(iListIndex), 1, " ")
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
    
    lblTCount = 0
    lblCount = objMicCul.DispNogrowthList(ssTable, pWsCd, pWsUnit, sRTs)
    lblTCount = Val(lblTCount) + Val(lblCount)
    
    lblGCount = objMicCul.DispGrowthList(lstGList, pWsCd, pWsUnit)
    lblTCount = Val(lblTCount) + Val(lblGCount)
    
    lblFCount = objMicCul.DispFinalList(lstFList, pWsCd, pWsUnit)
    lblTCount = Val(lblTCount) + Val(lblFCount)
    
    lblHCount = objMicCul.DispHoldingList(ssHTable, pWsCd, pWsUnit, sRTs)
    lblTCount = Val(lblTCount) + Val(lblHCount)

    lblECount = 0: ssETable.MaxRows = 0
    
    If chkStain.Value = 0 Then cboResult.Enabled = True
    
    ssTable.RowHeight(-1) = 20

End Sub

Private Sub ssETable_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim tmpcolor As Long
    
    If Col >= 0 And Row > 0 Then
    
        ssETable.Col = -1: ssETable.Row = Row
        tmpcolor = ssETable.ForeColor
        
        If tmpcolor = fSCItem Then
            ssETable.ForeColor = fGCItem
        Else
            ssETable.ForeColor = fSCItem
        End If
        
    End If

End Sub

Private Sub ssETable_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub '

    Call AddWorkSheet(ssETable, Row)
    lblECount = Val(lblECount) - 1
    lblCount = Val(lblCount) + 1

End Sub

Private Sub cmdIn1_Click()
    
    Dim i As Integer, sCnt As Integer
    ssETable.Col = 1: sCnt = 0
    
    For i = ssETable.MaxRows To 1 Step -1
        ssETable.Row = i
        If ssETable.ForeColor = fSCItem Then
            sCnt = sCnt + 1
            Call AddWorkSheet(ssETable, i)
        End If
    Next i

    lblECount = Val(lblECount) - sCnt
    lblCount = Val(lblCount) + sCnt

End Sub

Private Sub ssHTable_Click(ByVal Col As Long, ByVal Row As Long)
Dim tmpcolor As Long

    If Col >= 0 And Row > 0 Then
    
        ssHTable.Col = -1: ssHTable.Row = Row
        tmpcolor = ssHTable.ForeColor
        
        If tmpcolor = fSCItem Then
            ssHTable.ForeColor = fGCItem
        Else
            ssHTable.ForeColor = fSCItem
        End If
        
    End If

End Sub

Private Sub ssHTable_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

    Call AddWorkSheet(ssHTable, Row)
    lblHCount = Val(lblHCount) - 1
    lblCount = Val(lblCount) + 1

End Sub

Private Sub cmdIn2_Click()
    
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

Private Sub AddWorkSheet(ByVal pObj As Object, ByVal pRow As Integer)
    
    Dim sAccBuf As String

    ssTable.MaxRows = ssTable.MaxRows + 1
    
    pObj.Col = 1: pObj.COL2 = pObj.MaxCols
    pObj.Row = pRow: pObj.Row2 = pRow
    ssTable.Col = 1: ssTable.COL2 = ssTable.MaxCols
    ssTable.Row = ssTable.MaxRows: ssTable.Row2 = ssTable.MaxRows
    ssTable.Clip = pObj.Clip
    
    '## 5.1.7: 이상대(2005-05-30)
    '   - 과거결과를 빨강색으로 표시
    ssTable.Col = 6: ssTable.ForeColor = vbRed
    
    pObj.Row = pRow
    pObj.Action = ActionDeleteRow
    pObj.MaxRows = pObj.MaxRows - 1
End Sub

Private Sub ssTable_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim tmpcolor As Long
    
    If Col = ColNo.cnHold And Row > 0 Then
        ssTable.Row = Row
        ssTable.Col = Col
        ssTable.Value = (Val(ssTable.Value) + 1) Mod 2
        Exit Sub
    End If
    
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
    lblECount = Val(lblECount) + sCnt

End Sub

Private Sub ssTable_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub
    If chkStain.Value = 1 Then Exit Sub
    
    If Col = ColNo.cnHold And Row > 0 Then
        ssTable.Row = Row
        ssTable.Col = Col
        ssTable.Value = (Val(ssTable.Value) + 1) Mod 2
        Exit Sub
    End If

    MovetoETable Row
    lblCount = Val(lblCount) - 1
    lblECount = Val(lblECount) + 1

End Sub

Private Sub MovetoETable(ByVal pRow As Integer)
    
    Dim sAccBuf As String

    ssETable.MaxRows = ssETable.MaxRows + 1
    
    ssTable.Col = 1: ssTable.COL2 = ssTable.MaxCols
    ssTable.Row = pRow: ssTable.Row2 = pRow
    ssETable.Col = 1: ssETable.COL2 = ssETable.MaxCols
    ssETable.Row = ssETable.MaxRows: ssETable.Row2 = ssETable.MaxRows
    ssETable.Clip = ssTable.Clip
    
    ssTable.Row = pRow
    ssTable.Action = ActionDeleteRow
    ssTable.MaxRows = ssTable.MaxRows - 1
    
End Sub

Private Sub cmdMidEnter_Click()
    Me.MousePointer = vbHourglass
    Call VerifyResult(enStsCd.StsCd_LIS_MidRst)
    Me.MousePointer = vbDefault
End Sub
        
Private Sub VerifyResult(ByVal pStatus As String)

    Dim i As Integer
    Dim sPtid As String
    Dim sAccNo As String, sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim sTestCds As String, sWsCd As String, sRTs As String
    Dim sDate As String, sTime As String
    Dim tmpRs As Recordset
    Dim strSQL As String
    Dim strDate As String
    Dim tmpRS1  As Recordset
    Dim strTmp  As String
    
    If (ssTable.MaxRows + ssETable.MaxRows) = 0 Then Exit Sub
    
    If cboWSCode.ListIndex < 0 Or txtWSUnit = "" Then Exit Sub
    If ssTable.MaxRows > 0 And cboResult.ListIndex < 0 Then
        MsgBox "아직 결과를 선택하지 않았습니다. 배치 등록할 결과를 선택하세요", vbInformation, "배치등록"
        Exit Sub
    End If
   
    If ssTable.MaxRows >= 1 Then
    
        sDate = Format(GetSystemDate, CS_DateDbFormat)
        sTime = Format(GetSystemDate, CS_TimeDbFormat)
    
        sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
        sRTs = fWorkSheet(cboWSCode.ListIndex).WsRstType
'
        ' 접수번호 건당 처리하여 처리도중 에러 발생하더라도 이후부터 작업
        For i = 1 To ssTable.MaxRows
            
            ssTable.Row = i
            ssTable.Col = ColNo.cnHold
            If ssTable.Value = 0 Then   '보류(1)는 제외
                ssTable.Col = ColNo.cnAccNo: sAccNo = ssTable.Text
                sWorkArea = medGetP(sAccNo, 1, "-"): sAccDt = medGetP(sAccNo, 2, "-"): sAccSeq = medGetP(sAccNo, 3, "-")
                sAccDt = IIf(Mid(sAccDt, 1, 1) = "9", "19" & sAccDt, "20" & sAccDt)
                ssTable.Col = ColNo.cnTestCd: sTestCds = ssTable.Text
                ssTable.Col = 2: sPtid = ssTable.Text
'                strDate = TO_DATE(TO_CHAR(GetSystemDate, "yyyymmdd"), "yyyymmdd")
                strDate = Format(GetSystemDate, "yyyyMMdd")
                    
                If pStatus = 5 Then
                    strSQL = ""
                    strSQL = strSQL + "SELECT a.ptid,                     "
                    strSQL = strSQL + "       a.orddoct,a.rcvdt,          "
                    strSQL = strSQL + "       a.deptcd,                   "
                    strSQL = strSQL + "       b.testcd,                   "
                    strSQL = strSQL + "       b.orddt,                    "
                    strSQL = strSQL + "       b.ordno,                    "
                    strSQL = strSQL + "       b.ordseq                    "
                    strSQL = strSQL + "  FROM S2LAB201 a, S2LAB404 b      "
                    strSQL = strSQL + " WHERE a.workarea = '" & sWorkArea & "'  "
                    strSQL = strSQL + "   AND a.accdt = '" & sAccDt & "'        "
                    strSQL = strSQL + "   AND a.accseq = '" & sAccSeq & "'      "
                    strSQL = strSQL + "   AND a.WORKAREA = b.WORKAREA     "
                    strSQL = strSQL + "   AND a.ACCDT = b.ACCDT           "
                    strSQL = strSQL + "   AND a.ACCSEQ = b.ACCSEQ         "
                    strSQL = strSQL + "   AND b.TESTCD = 'B4062'          "
                    
                    Set tmpRs = New Recordset
                    tmpRs.Open strSQL, DBConn
                                    
                    If tmpRs.RecordCount > 0 Then
                        strSQL = ""
                        strSQL = " SELECT * FROM MDEXMORT WHERE PATNO = '" & "" & tmpRs.Fields("PtId") & "' AND ORDDATE = '" & Format("" & tmpRs.Fields("ORDDT").Value, "####-##-##") & "' "
                            
                        Set tmpRS1 = New Recordset
                        tmpRS1.Open strSQL, DBConn
                        
                        If tmpRS1.RecordCount > 0 Then
                            strTmp = "" & tmpRS1.Fields("PATSITE").Value
                        End If
                        
                        If strTmp = "O" And tmpRs.Fields("orddoct").Value <> "" Then
                            strSQL = ""
                            strSQL = strSQL & " INSERT INTO aclisrpt (PATNO,MEDDATE,MEDDEPT,MEDDR,SUGACODE,ORDSEQNO,RCPYN,EDITID,EDITIP,EDITDATE,WORKAREA,execdate )"
                            strSQL = strSQL & " values('" & tmpRs.Fields("ptid").Value & "' ,"
                            strSQL = strSQL & "        '" & tmpRS1.Fields("meddate").Value & "' ,"
                            strSQL = strSQL & "        '" & tmpRs.Fields("deptcd").Value & "' ,"
                            strSQL = strSQL & "        '" & tmpRs.Fields("orddoct").Value & "' ,"
                            strSQL = strSQL & "        '" & tmpRs.Fields("testcd").Value & "' ,"
                            strSQL = strSQL & "        '" & tmpRs.Fields("ordno").Value & "' ,"
                            strSQL = strSQL & "        '' ,"
                            strSQL = strSQL & "        '" & ObjSysInfo.EmpId & "' ,"
                            strSQL = strSQL & "        '' ,"
                            strSQL = strSQL & "        '" & strDate & "' ,"
                            strSQL = strSQL & "        '" & sAccNo & "' ,"
                            strSQL = strSQL & "        '" & Format(tmpRs.Fields("rcvdt").Value, "####-##-##") & "')"
                            
                            DBConn.Execute strSQL
                        End If
                    End If
                End If
                
                Set tmpRs = Nothing
                Set tmpRS1 = Nothing
                
                On Error GoTo DBExecError
        
                DBConn.BeginTrans
                If Not objMicCul.SaveNogrowthBatch(sWorkArea, sAccDt, sAccSeq, sTestCds, fNGCode(cboResult.ListIndex), _
                                                   pStatus, sDate, sTime, ObjSysInfo.EmpId, sWsCd) Then GoTo DBExecError
                                                   
                '===========================================================================================
                '결과보고대기내역 생성(2002/09/05)
                Dim tmpBussDiv  As String
                Dim tmpDept     As String
                Dim tmpDoct     As String
                Dim tmpPtid     As String
                
                strTmp = ""
                tmpBussDiv = "": tmpDept = "": tmpDoct = ""
                
                
'                '---------------------------------------------------------
'                ' 그룹검사의 결과 Verify를 위해서 사용한다.
'                ' 일단 삭제후 테스트 해본후에 살리자(성모자애에서는 안쓴다.
'                '---------------------------------------------------------
'
'                SELECT Case
'                    Case "04", "05"
'                        strTmp = objMicRst.Get_OrderInFo(sWorkArea, sAccDt, sAccSeq)
'
'                        '-- 원본 ==========================================================
'                        tmpBussDiv = medGetP(strTmp, 1, COL_DIV)
'                        tmpDoct = medGetP(strTmp, 4, COL_DIV)
'                        tmpPtid = medGetP(strTmp, 5, COL_DIV)
'                        SELECT Case tmpBussDiv
'                            Case "1": tmpDept = medGetP(strTmp, 3, COL_DIV)
'                            Case "2": tmpDept = medGetP(strTmp, 2, COL_DIV)
'                            Case Else
'                                tmpDept = medGetP(strTmp, 2, COL_DIV)
'                                If tmpDept = "" Then tmpDept = medGetP(strTmp, 3, COL_DIV)
'                        End SELECT
'                        '==================================================================
'
'                        If Not objMicRst.SubmitVerifyList(tmpDept, sDate, sTime, tmpPtid, pStatus, ObjMyUser.EmpId, tmpDoct, tmpBussDiv) Then GoTo DBExecError
'                End SELECT
                '============================================================================================================================
                                                   
                DBConn.CommitTrans
            End If
            
        Next i
    End If
    
    ' 보류 리스트 처리
    On Error GoTo DBExecError
   
    DBConn.BeginTrans
    sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
    If lblECount.Caption > 0 Then Call objMicCul.SaveHoldingList(ssETable, sWsCd, txtWSUnit.Text)
    DBConn.CommitTrans



'    '감염관리
    Call ICSNoGroWthSave(ssTable, cboResult.Text)
    

'    If icsresultchk = True Then
'        Dim objICS      As New clsICSResultChk
'        Dim arytmp()    As String
'        Dim jj          As Integer
'
'        For i = 1 To ssTable.MaxRows
'            ssTable.Row = i
'            ssTable.Col = ColNo.cnHold
'            If ssTable.Value = 0 Then   '보류(1)는 제외
'                ssTable.Col = 1: sAccNo = ssTable.Text
'                sWorkArea = medGetP(sAccNo, 1, "-"): sAccDt = medGetP(sAccNo, 2, "-"): sAccSeq = medGetP(sAccNo, 3, "-")
'                sAccDt = IIf(Mid(sAccDt, 1, 1) = "9", "19" & sAccDt, "20" & sAccDt)
'                ssTable.Col = 9: sTestCds = ssTable.Text
'                sTestCds = Replace(sTestCds, "'", "")
'                arytmp() = Split(sTestCds, ",")
'
'                ssTable.Col = 2: sPtid = ssTable.Text
'                For jj = LBound(arytmp()) To UBound(arytmp())
'                    Call objICS.ICSNoGroWthSave(sPtid, sWorkArea, sAccDt, sAccSeq, arytmp(jj), cboResult.Text)
'                Next
'
'            End If
'        Next
'        Set objICS = Nothing
'    End If
        
    
    
    Call txtWSUnit_KeyPress(vbKeyReturn)
    
    Exit Sub

DBExecError:
    DBConn.RollbackTrans
End Sub

Private Sub ssTable_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim strVfyDt        As String
    Dim strVfyTm        As String
    Dim strVfyID        As String
    Dim strResult       As String
    Dim strRcpNo        As String
    
    Dim strLVfyDt       As String
    Dim strLVfyTm       As String
    Dim strLVfyID       As String
    Dim strLResult      As String
    
    Dim RS          As Recordset
    Dim tmpToolTip  As String
    Dim SSQL        As String
    Dim WorkArea       As String
    Dim AccDt      As String
    Dim AccSeq      As Integer


    If Col = 12 Then Exit Sub
    If Row < 1 Then Exit Sub
    With ssTable
        If Col = ColNo.cnLstRst Then
            .Row = Row: .Col = Col
            If .Value = "" Then Exit Sub
            strResult = .Value
            MultiLine = 1
            
            .Col = 15:  strVfyDt = .Value '보고일자
            .Col = 16:  strVfyTm = .Value '보고일시
            .Col = 17:  strVfyID = .Value '보고자
        
            .Col = 1
            strRcpNo = Trim(.Value)
            WorkArea = Mid(strRcpNo, 1, 2)
            AccDt = Mid(Now, 1, 2) & Mid(strRcpNo, 4, 4)
            AccSeq = CInt(Mid(strRcpNo, 9))
        
'            TipText = vbCRLF & " (과거) 결 과 값 : " & strResult & vbCRLF & _
                               " (과거) 결과일시 : " & strVfyDt & " " & _
                                                       strVfyTm & vbCRLF & _
                               " (과거) 보 고 자 : " & GetEmpNm(strVfyID) & vbCRLF
            
            SSQL = " SELECT lastrst,lastvfydt,lastvfytm,lastvfyid " & _
                   "  FROM " & T_LAB404 & _
                   " WHERE " & DBW("workarea=", WorkArea) & _
                   "   AND " & DBW("accdt=", AccDt) & _
                   "   AND " & DBW("accseq=", AccSeq)

            Set RS = New Recordset
            RS.Open SSQL, DBConn
            If Not RS.EOF Then
                Do Until RS.EOF
                    If Not IsNull(RS.Fields("lastvfydt").Value) Then
                        strLResult = RS.Fields("lastrst").Value & ""
                        
                        tmpToolTip = vbCRLF & " (과거) 결 과 값 : " & strLResult & " " & vbCRLF & _
                                              " (과거) 결과일시 : " & Format(RS.Fields("lastvfydt").Value & "", "0###-##-##") & " " & _
                                                                      Format(Mid(RS.Fields("lastvfytm").Value & "", 1, 4), "0#:##") & vbCRLF & _
                                              " (과거) 보 고 자 : " & GetEmpNm(RS.Fields("lastvfyid").Value & "") & vbCRLF
                
                    End If
                    RS.MoveNext
                Loop
            End If
            TipWidth = 5000
            Call .SetTextTipAppearance("돋움체", 10, False, False, &HEEFDF2, &H996666)
        Else
' 2008.10.24. 바코드를 보여주기위해 수정함.

            .Row = Row: .Col = 17
            If .Value = "" Then Exit Sub
            MultiLine = 1
'            TipText = vbCRLF & "  " & .Text & vbCRLF
            TipWidth = 3000
            tmpToolTip = "   " & .Text
            Call .SetTextTipAppearance("돋움체", 16, True, False, &HEEFDF2, DCM_Red)
        End If
        
        TipText = tmpToolTip
        MultiLine = 1
        .TextTipDelay = 1000
'        Call .SetTextTipAppearance("돋움체", 10, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
    
    Set RS = Nothing

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

Private Sub cmdRPrint_Click()

    With ssTable
        
        If .DataRowCnt < 1 Then
            Exit Sub
        End If
        
        .PrintMarginTop = 100
        .PrintJobName = "No Growth 결과"
        
        .PrintAbortMsg = "No Growth 결과를 출력중입니다. "

        .PrintOrientation = PrintOrientationLandscape
        If Printer.PaperSize = vbPRPSA4 Then
            .PrintMarginLeft = 1700
            .PrintMarginRight = 100
            .PrintMarginTop = 800
            .PrintMarginBottom = 800
        Else
            .PrintMarginTop = 300
            .PrintMarginBottom = 500
            .PrintMarginLeft = 250
            .PrintMarginRight = 100
        End If
        .PrintColor = False
        .PrintFirstPageNumber = 1
        
        .PrintHeader = "/n/n/l/fb1 " & "♧ No Growth 결과 - " & "접수 마감일/시:" & lblRcvDT.Caption & "   " & _
                                       "Worksheet 작성일/시" & lblBltDate.Caption & " /c/fb1/n/n"
        
        .PrintFooter = "/c/p/fb1"
        
        .PrintGrid = False
        .PrintShadows = False
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintPageEnd = 2
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        '.PrintGrid = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll
         
        .Action = ActionPrint
        .GridSolid = True
    End With
    
    'If chkGraph.Value = 1 Then Call PrintGraph

End Sub

