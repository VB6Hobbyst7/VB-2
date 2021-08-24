VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm155Accession 
   BackColor       =   &H00DBE6E6&
   Caption         =   "검체접수"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14790
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   14790
   WindowState     =   2  '최대화
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
      TabIndex        =   32
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
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
      TabIndex        =   31
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   75
      ScaleHeight     =   390
      ScaleWidth      =   7440
      TabIndex        =   27
      Top             =   1890
      Width           =   7440
      Begin MedControls1.LisLabel lblMsg 
         Height          =   360
         Left            =   15
         TabIndex        =   28
         Top             =   15
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   635
         BackColor       =   13434879
         ForeColor       =   5584725
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
         Caption         =   "이미 접수된 검체입니다 !!"
         LeftGab         =   100
      End
   End
   Begin VB.Frame fraReceive 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   75
      TabIndex        =   6
      Tag             =   "15502"
      Top             =   2205
      Width           =   7455
      Begin VB.CommandButton cmdOrderView 
         BackColor       =   &H00F4F0F2&
         Caption         =   "처방별조회(&C)"
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
         Left            =   5850
         Style           =   1  '그래픽
         TabIndex        =   45
         Top             =   180
         Width           =   1500
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   1425
         TabIndex        =   7
         Top             =   510
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   1410
         TabIndex        =   8
         Top             =   870
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel lblPtId 
         Height          =   300
         Left            =   1425
         TabIndex        =   9
         Top             =   180
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   529
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblWard 
         Height          =   315
         Left            =   1410
         TabIndex        =   16
         Top             =   1215
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   225
         TabIndex        =   34
         Top             =   165
         Width           =   1140
         _ExtentX        =   2011
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
         Caption         =   "환자   ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReceptNo 
         Height          =   315
         Left            =   225
         TabIndex        =   35
         Top             =   510
         Width           =   1140
         _ExtentX        =   2011
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
         Caption         =   "성      명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   225
         TabIndex        =   36
         Top             =   870
         Width           =   1140
         _ExtentX        =   2011
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
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   225
         TabIndex        =   37
         Top             =   1215
         Width           =   1140
         _ExtentX        =   2011
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
         Caption         =   "병     동"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   75
      TabIndex        =   30
      Top             =   45
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   556
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
      Caption         =   "접수대상검체 확인"
      LeftGab         =   100
   End
   Begin VB.Frame fraInput 
      BackColor       =   &H00DBE6E6&
      Height          =   1605
      Left            =   75
      TabIndex        =   5
      Top             =   270
      Width           =   14385
      Begin VB.CommandButton cmdExecute 
         BackColor       =   &H00F4F0F2&
         Caption         =   "일괄접수 실행(&S)"
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
         Left            =   7485
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   1035
         Width           =   1740
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "개별접수"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   495
         TabIndex        =   20
         Top             =   1065
         Width           =   1260
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "일괄접수"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   495
         TabIndex        =   19
         Top             =   405
         Width           =   1260
      End
      Begin VB.PictureBox picLabNo 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   4020
         ScaleHeight     =   330
         ScaleWidth      =   2865
         TabIndex        =   15
         Top             =   1095
         Width           =   2925
         Begin VB.TextBox txtWorkArea 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   15
            MaxLength       =   2
            TabIndex        =   2
            Top             =   60
            Width           =   600
         End
         Begin VB.TextBox txtAccDt 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            MaxLength       =   6
            TabIndex        =   3
            Top             =   60
            Width           =   1080
         End
         Begin VB.TextBox txtAccNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2115
            TabIndex        =   4
            Top             =   60
            Width           =   705
         End
         Begin VB.Line Line2 
            X1              =   1950
            X2              =   2100
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Line Line1 
            X1              =   660
            X2              =   810
            Y1              =   180
            Y2              =   180
         End
      End
      Begin VB.TextBox txtBarcode 
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
         Height          =   390
         Left            =   4020
         TabIndex        =   1
         Top             =   615
         Width           =   2910
      End
      Begin VB.CheckBox chkReader 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&Barcode Reader 사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2670
         TabIndex        =   0
         Tag             =   "15501"
         Top             =   255
         Width           =   2535
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   375
         Index           =   3
         Left            =   2655
         TabIndex        =   41
         Top             =   615
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
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
         Caption         =   "Barcode"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   375
         Index           =   4
         Left            =   2655
         TabIndex        =   42
         Top             =   1095
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   661
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
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   2205
         X2              =   2205
         Y1              =   225
         Y2              =   1600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   2220
         X2              =   2220
         Y1              =   225
         Y2              =   1600
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   33
      Top             =   3825
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   556
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
      Caption         =   "병동 선택"
      LeftGab         =   100
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4995
      Left            =   75
      TabIndex        =   10
      Top             =   4065
      Width           =   7440
      Begin VB.TextBox txtMesg 
         BackColor       =   &H00F7FDF8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   44
         ToolTipText     =   "검사 리마크를 입력하세요."
         Top             =   3690
         Width           =   7050
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   2445
         Left            =   180
         TabIndex        =   14
         Top             =   855
         Width           =   6960
         _Version        =   196608
         _ExtentX        =   12277
         _ExtentY        =   4313
         _StockProps     =   64
         BackColorStyle  =   1
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
         GridColor       =   14737632
         MaxCols         =   4
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "Lis155.frx":0000
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   1470
         TabIndex        =   11
         Top             =   510
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblStoreNm 
         Height          =   315
         Left            =   4725
         TabIndex        =   12
         Top             =   510
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLabNo 
         Height          =   315
         Left            =   1470
         TabIndex        =   17
         Top             =   165
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   195
         TabIndex        =   38
         Top             =   165
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "접수번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   315
         Left            =   195
         TabIndex        =   39
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "검     체"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   3450
         TabIndex        =   40
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "검체보관방법"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   315
         Left            =   195
         TabIndex        =   43
         Top             =   3315
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "Remark"
         Appearance      =   0
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F7F3F8&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   1260
         Index           =   0
         Left            =   210
         Top             =   3660
         Width           =   7110
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "응 급"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3780
         TabIndex        =   13
         Tag             =   "105"
         Top             =   210
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Shape shpStat 
         BackColor       =   &H000000FF&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   360
         Left            =   3435
         Top             =   150
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame fraMulti 
      BackColor       =   &H00DBE6E6&
      Height          =   6660
      Left            =   7575
      TabIndex        =   18
      Top             =   1785
      Width           =   6885
      Begin VB.CommandButton cmdClearRow 
         BackColor       =   &H00EDE2ED&
         Caption         =   "Clear Row"
         Height          =   330
         Left            =   1365
         Style           =   1  '그래픽
         TabIndex        =   29
         Top             =   210
         Width           =   1185
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Clear List"
         Height          =   330
         Left            =   165
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   210
         Width           =   1185
      End
      Begin MSComctlLib.ListView lstAccList 
         Height          =   5940
         Left            =   150
         TabIndex        =   22
         Top             =   555
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   10478
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15857140
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblErrCnt 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "150"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   5835
         TabIndex        =   25
         Top             =   300
         Width           =   525
      End
      Begin VB.Label lblTotCnt 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "150"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4410
         TabIndex        =   24
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "총                건,  오류                 건"
         Height          =   225
         Left            =   4155
         TabIndex        =   23
         Top             =   330
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frm155Accession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tmpAccDt As String
Private objMySql As New clsLISSqlAccession
Private blnExeFg As Boolean

Private Const CS_AccSuccess = "정상"
Private Const lngMaxRows = 9
Private Const lngRowHeight = 12.5


'% 바코드 리더기를 사용할 것인지에 대한 여부
Private Sub chkReader_Click()
    Call ClearRtn
    If chkReader.value = 1 Then
        txtBarcode.Locked = False
        txtBarcode.BackColor = DCM_White    '&H80000005
        picLabNo.Enabled = False
        picLabNo.BackColor = DCM_LightGray
        txtWorkArea.BackColor = DCM_LightGray
        txtAccDt.BackColor = DCM_LightGray
        txtAccNo.BackColor = DCM_LightGray
        txtBarcode.SetFocus
    Else
        txtBarcode.Locked = True
        txtBarcode.BackColor = DCM_LightGray
        picLabNo.Enabled = True
        picLabNo.BackColor = DCM_White
        txtWorkArea.BackColor = DCM_White
        txtAccDt.BackColor = DCM_White
        txtAccNo.BackColor = DCM_White
        txtWorkArea.SetFocus
    End If
End Sub

Private Sub cmdClear_Click()
    Dim intTemp As Integer
    
    '## 5.1.7: 이상대(2005-05-31)
    '   - 바코드만 리딩하고 접수하지 않은 검체가 있을경우 확인 메시지를 출력하도록 수정
    If lstAccList.ListItems.Count > 0 Then
        If lstAccList.ListItems(1).SubItems(2) = "" Then
            intTemp = MsgBox("접수되지 않은 검체가 있습니다. 화면을 지울까요?", vbYesNo + vbDefaultButton2 + vbQuestion, "확인")
            If intTemp = vbNo Then Exit Sub
        End If
    End If
    
    Call ClearRtn
    Call cmdReset_Click
    optOption(0).value = True
    txtWorkArea.Text = ""
    txtAccDt.Text = ""
    txtAccNo.Text = ""
    txtBarcode.Text = ""
    If chkReader.value = 1 Then
        txtBarcode.SetFocus
    Else
        txtWorkArea.SetFocus
    End If
End Sub

Private Sub cmdClearRow_Click()

    Dim i As Long
    
    For i = lstAccList.ListItems.Count To 1 Step -1
        If lstAccList.ListItems(i).Checked Then
            lstAccList.ListItems.Remove (i)
        End If
    Next
    For i = 1 To lstAccList.ListItems.Count
        lstAccList.ListItems(i).Text = CStr(i)
    Next

End Sub

Private Sub cmdExecute_Click()
    
    If lstAccList.ListItems.Count <= 0 Then
        MsgBox "바코드가 한 건도 리딩되지 않았습니다.", vbExclamation, "검체접수"
        Exit Sub
    End If
    
    optOption(0).Enabled = False
    optOption(1).Enabled = False
    txtBarcode.Enabled = False
    txtBarcode.BackColor = DCM_LightGray
    cmdExecute.Enabled = False
    cmdExit.Enabled = False
    cmdClear.Enabled = False
    'fraInput.Enabled = False
    
    Dim blnAccFg As Boolean
    Dim i As Integer
        
    blnExeFg = True
    For i = 1 To lstAccList.ListItems.Count
        txtBarcode.Text = lstAccList.ListItems(i).SubItems(1)
        
        Call ClearRtn
        
        blnAccFg = DisplayOrder(0, i)
        If blnAccFg Then Call DoAccession(i)
        
        txtBarcode.Text = ""
        txtWorkArea.Text = ""
        txtAccDt.Text = ""
        txtAccNo.Text = ""
        DoEvents
    Next
    
'    For i = lstAccList.ListItems.Count To 1 Step -1
'        If lstAccList.ListItems(i).SubItems(2) = CS_AccSuccess Then '"정상"
'            lstAccList.ListItems.Remove (i)
'        Else
'            lblErrCnt.Caption = Val(lblErrCnt.Caption) + 1
'        End If
'        DoEvents
'    Next
    blnExeFg = False
    
    cmdExit.Enabled = True
    cmdClear.Enabled = True

End Sub

'% 종료
Private Sub cmdExit_Click()
    Set objMySql = Nothing
    Unload Me
    Set frm155Accession = Nothing
End Sub

Private Sub cmdOrderView_Click()
'    Dim i As Integer
'    Dim pFrmName As String
'    If Len(lblPtId.Caption) < 2 Then GoTo End2Stop
'
'    pFrmName = "frm401ResultView"
'
''    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
''    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied
'
'    medMain.lblSubMenu.Caption = "처방결과조회" 'medGetP(Button.Tag, 1, "(")
'
'    frmLisReview.ButtonKey = "LIS501" 'Button.Key
'    frmLisReview.PtId = lblPtId.Caption
'    frmLisReview.Show
'    frmLisReview.ZOrder 0
'    frmLisReview.ShowThisForm
'
'    Exit Sub
'
'PermissionDenied:
'
''    blnFormShow = False
'    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
'End2Stop:

    Dim i As Integer
    Dim pFrmName As String
    If Len(lblPtId.Caption) < 2 Then GoTo End2Stop

    pFrmName = "frm401ResultView"
    
    medMain.lblSubMenu.Caption = "처방결과조회" 'medGetP(Button.Tag, 1, "(")
    
    frmLisReview.ButtonKey = "LIS155A" 'Button.Key
    frmLisReview.PtId = lblPtId.Caption
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

    Exit Sub

PermissionDenied:
   
'    blnFormShow = False
    MsgBox "이 화면을 사용할 수 있는 권한이 없습니다.", vbExclamation, "Security Check!"
End2Stop:

End Sub

Private Sub cmdReset_Click()
    lstAccList.ListItems.Clear
    'fraInput.Enabled = True
    optOption(0).Enabled = True
    optOption(1).Enabled = True
    txtBarcode.Enabled = True
    cmdExecute.Enabled = True
    fraInput.Enabled = True
    txtBarcode.BackColor = DCM_White
    lblMsg.Caption = ""
    lblTotCnt.Caption = ""
    lblErrCnt.Caption = ""
    txtBarcode.SetFocus
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

'% 폼 로드
Private Sub Form_Load()
    Me.Show
    chkReader.value = 1
    medInitLvwHead lstAccList, "Seq,검체번호,Message,SeqNo", "-1000,-300,1000,300"

    optOption(0).value = True
    Call cmdReset_Click
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub optOption_Click(Index As Integer)
    cmdExecute.Enabled = optOption(0).value
    If optOption(0).value Then
        chkReader.value = 1
        chkReader.Enabled = False
        fraMulti.Enabled = True
    Else
        chkReader.Enabled = True
        fraMulti.Enabled = False
    End If
    If chkReader.value = 1 Then
        txtBarcode.SetFocus
    Else
        txtWorkArea.SetFocus
    End If
End Sub

Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    With tblOrdSheet
        .Row = Row: .Col = 4
        txtMesg.Text = .value
    End With
End Sub

Private Sub txtAccDt_Change()
    Dim strDt As String
    Dim strYYYY As String
    
    strDt = Mid(txtAccDt.Text, 1, 2) & "-01-01"
    strYYYY = Format(strDt, "yyyy")
    tmpAccDt = strYYYY & Mid(txtAccDt.Text, 3)
    
    If Len(txtAccDt.Text) = txtAccDt.MaxLength Then
       If txtAccNo.Enabled Then txtAccNo.SetFocus
    End If
End Sub

Private Sub txtAccDt_GotFocus()
    With txtAccDt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccDt_KeyPress(KeyAscii As Integer)
    If txtAccDt.Text = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then txtAccNo.SetFocus
End Sub

Private Sub txtAccNo_GotFocus()
    With txtAccNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'% 접수번호를 입력했을 경우
Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = vbKeyReturn Then
    
        If txtAccNo.Text = "" Then Exit Sub
           
        Call ClearRtn
        
        Dim blnAccFg As Boolean
        blnAccFg = DisplayOrder(1)
        If blnAccFg Then Call DoAccession
        
        txtBarcode.Text = ""
        txtWorkArea.Text = ""
        txtAccDt.Text = ""
        txtAccNo.Text = ""
        If chkReader.value = 1 Then
            txtBarcode.SetFocus
        Else
            txtWorkArea.SetFocus
        End If
    
    End If

End Sub

'% 리더기로 검체번호를 리딩했을 경우...
Private Sub txtBarcode_KeyPress(KeyAscii As Integer)



    If KeyAscii = vbKeyReturn Then
        Call medClearTable(tblOrdSheet)
        If txtBarcode.Text = "" Then Exit Sub
         Dim blnAccFg As Boolean
        If optOption(0).value Then
            
            '일괄접수
            If blnExeFg Then Exit Sub
            blnAccFg = DisplayOrder(0)
            'lstAccList.ListItems.Add , , txtBarcode.Text
            lstAccList.ListItems.Add , , lstAccList.ListItems.Count + 1
            lstAccList.ListItems(lstAccList.ListItems.Count).SubItems(1) = txtBarcode.Text
            lblTotCnt.Caption = lstAccList.ListItems.Count
            lstAccList.SetFocus
            SendKeys "^{END}"
            DoEvents
            txtBarcode.Text = ""
            txtBarcode.SetFocus
        
        Else
            
            '개별접수
            Call ClearRtn
           
            blnAccFg = DisplayOrder(0)
            If blnAccFg Then Call DoAccession
            
            txtBarcode.Text = ""
            txtWorkArea.Text = ""
            txtAccDt.Text = ""
            txtAccNo.Text = ""
            If chkReader.value = 1 Then
                txtBarcode.SetFocus
            Else
                txtWorkArea.SetFocus
            End If
        End If
    End If

End Sub


'% 접수번호 또는 검체번호를 기준으로 발생한 검사내역을 검색한다.
Private Function DisplayOrder(ByVal QueryOption As Integer, Optional ByVal ii As Integer) As Boolean

    Dim objRs As Recordset
    Dim tmpSQL As String
    Dim tmpBarcode As String
    Dim i As Long

    
    If QueryOption = 1 Then
        tmpSQL = objMySql.SqlOrdersForAccess(1, txtWorkArea.Text, tmpAccDt, txtAccNo.Text)
    Else
        '** 원본 ============================================================
        'tmpBarcode = CStr(Mid(txtBarcode.Text, 1, Len(txtBarcode.Text) - 1))
        '====================================================================
        
        '** 예수병원 ========================================================
        tmpBarcode = CStr(Mid(txtBarcode.Text, 1, Len(txtBarcode.Text)))
        '====================================================================
        
        tmpSQL = objMySql.SqlOrdersForAccess(2, Mid(tmpBarcode, 1, P_SpcYyLength), Val(Mid(tmpBarcode, P_SpcYyLength + 1)))
    End If
    
    
    Set objRs = New Recordset
    objRs.Open tmpSQL, DBConn
    
    If objRs.EOF Then
        DisplayOrder = False
        lblMsg.Caption = "해당 데이타가 없습니다 !!"
        If ii > 0 Then lstAccList.ListItems(ii).SubItems(2) = "해당 데이타가 없습니다 !!"
        'MsgBox "해당 데이타가 없습니다 !!", vbOKOnly + vbExclamation, "Message"
        GoTo NoData
    End If
    
    txtWorkArea.Text = "" & objRs.Fields("WorkArea").value
    txtAccDt.Text = Mid("" & objRs.Fields("AccDt").value, 3)
    txtAccNo.Text = "" & objRs.Fields("AccSeq").value
    
    lblLabNo.Caption = "" & objRs.Fields("WorkArea").value & "-" & _
                        Mid(objRs.Fields("AccDt").value, 3) & "-" & _
                        objRs.Fields("AccSeq").value
    lblPtId.Caption = "" & objRs.Fields("PtId").value
    
    '감염관리
    Call ICSPatientMark(objRs.Fields("ptid").value & "", enICSNum.LIS_ALL)
    
    
    lblPtNm.Caption = "" & objRs.Fields("PtNm").value
    lblDeptNm.Caption = "" & objRs.Fields("DeptNm").value
    lblWard.Caption = "" & objRs.Fields("Location").value
    If objRs.Fields("StatFg").value = "1" Then
        shpStat.Visible = True
        lblStat.Visible = True
    Else
        shpStat.Visible = False
        lblStat.Visible = False
    End If
    lblSpcNm.Caption = "" & objRs.Fields("SpcNm").value
    lblStoreNm.Caption = "" & objRs.Fields("StoreCd").value
    
    If objRs.Fields("StsCd").value >= enStsCd.StsCd_LIS_Accession Then
        DisplayOrder = False
        lblMsg.Caption = "이미 접수된 검체입니다 !!"
        If ii > 0 Then lstAccList.ListItems(ii).SubItems(2) = "이미 접수된 검체입니다(" & lblLabNo.Caption & ")"
        'MsgBox "이미 접수된 검체입니다 !!", vbOKOnly + vbExclamation, "Message"
    Else
        lblMsg.Caption = ""
        DisplayOrder = True
    End If
    
    With tblOrdSheet
        If objRs.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
        Else
            .MaxRows = objRs.RecordCount
        End If
        For i = 1 To objRs.RecordCount
           .Row = i
           .Col = 1: .value = objRs.Fields("OrdDt").value & ""
           .Col = 2: .value = objRs.Fields("TestNm").value & ""
                     .ForeColor = DCM_LightBlue        '약간 파란색
           .Col = 3: .value = objRs.Fields("OrdCd").value & ""
           .Col = 4: .value = objRs.Fields("mesg").value & ""
           objRs.MoveNext
        Next
        .RowHeight(-1) = lngRowHeight
    End With
    Call tblOrdSheet_Click(1, 1)
NoData:
    Set objRs = Nothing

End Function

'% 접수Procedure를 수행한다.
Private Sub DoAccession(Optional ByVal ii As Integer = 0)

    Dim objAccess  As New clsLISAccession
    Dim blnSuccess As Boolean
    Dim strRcvDt   As String
    Dim strRcvTm   As String
      
    '데이타베이스의 날짜/시간으로 System Date/Time을 셋팅...
    Date = GetSystemDate
    strRcvDt = Format(GetSystemDate, "yyyymmdd")
    Time = GetSystemDate
    strRcvTm = Format(GetSystemDate, "hhmmss")
      
    MouseRunning  '13
      
    With objAccess
'        Call .SetDatabase(DbConn)
        blnSuccess = .DoAccession(txtWorkArea.Text, tmpAccDt, txtAccNo.Text, ObjMyUser.EmpId)
        If blnSuccess Then
            'lblMsg.Caption = "정상적으로 접수되었습니다 !!"
            If ii > 0 Then lstAccList.ListItems(ii).SubItems(2) = "정상"
            
            '리스트뷰에 구해온번호 넣어놓구...
            '-----------------------------------------------------------------
            '접수가 정상적으로 수행되면 WorkArea 별 접수 순번을 번호 정보 테이블에 부여한다.
            '-- Parameter (WorkArea, 접수일자(RcvDt))
            If ii > 0 Then
                lblMsg.Caption = "정상적으로 접수되었습니다 !!"
                lstAccList.ListItems(ii).SubItems(3) = GetSeqNo(txtWorkArea.Text, strRcvDt, tmpAccDt, txtAccNo.Text)
            Else
                lblMsg.Caption = "정상적으로 접수되었습니다 !!" & "  일련번호:" & GetSeqNo(txtWorkArea.Text, strRcvDt, tmpAccDt, txtAccNo.Text)
            End If
            '-----------------------------------------------------------------
        Else
            lblMsg.Caption = "오류 발생 !! (" & lblLabNo.Caption & ")"
            If ii > 0 Then lstAccList.ListItems(ii).SubItems(2) = "오류 발생 !!"
        End If
    End With
    
    '-- 업무나열서 순번(Rack 위치 확인 위한 일련번호)
    ' - 추가작업 By M.G.Choi
'    Call RackNo_Seq_Insert(txtWorkArea.Text, tmpAccDt, txtAccNo.Text, strRcvDt, strRcvTm)
    
    Set objAccess = Nothing
    
    MouseDefault
    
End Sub

'Private Sub RackNo_Seq_Insert(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccNo As String, _
'                              ByVal pRcvDt As String, ByVal pRcvTm As String)
'    Dim ObjAccess As New clsLISSqlAccession
'    Dim strSQL    As String
'
'    strSQL = ObjAccess.SqlRackNoInsert(pWorkArea, pAccDt, pAccNo, pRcvDt, pRcvTm, "")
'
'    Call dbconn.Execute(strSQL)
'
'    Set ObjAccess = Nothing
'
'End Sub
'
'% 워크애리어별 순번을 부여한다.
Private Function GetSeqNo(ByVal pWorkArea As String, ByRef pRcvDt As String, _
                          ByVal pAccDt As String, ByVal pAccSeq As String) As String

    Dim objSeq As clsLISSqlCollection
    Dim tmpRs As New Recordset
    Dim tmpSQL As String
    Dim LabDiv As String
    Dim tmpStr As String
    Dim tmpRng1 As Integer, tmpRng2 As Integer
    Dim tmpSpcGrp As String
    Dim tmpAccNo  As String

    GetSeqNo = 0

    tmpRng1 = 1
    tmpRng2 = 9999
    tmpSpcGrp = "0"

    Set objSeq = New clsLISSqlCollection
    
    tmpAccNo = pWorkArea & pAccDt & pAccSeq
    
    tmpSQL = objSeq.CreateSql_SeqNo(pWorkArea, pRcvDt, tmpSpcGrp, 4, tmpAccNo)
    
    On Error GoTo Err_Trap

    DBConn.BeginTrans
    DBConn.Execute tmpSQL   'Lock 걸림

    '// Sql문장 생성
    tmpSQL = objSeq.CreateSql_SeqNo(pWorkArea, pRcvDt, tmpSpcGrp, 5, tmpAccNo)
    '//
    tmpRs.Open tmpSQL, DBConn
    
    If tmpRs.EOF Then
        GetSeqNo = tmpRng1
        tmpSQL = objSeq.CreateSql_SeqNo(pWorkArea, pRcvDt, GetSeqNo, 2, tmpAccNo, GetSeqNo)
    Else
        GetSeqNo = Val(tmpRs.Fields("Seq").value & "")
        If GetSeqNo < tmpRng1 Then
            GetSeqNo = tmpRng1
        Else
            GetSeqNo = GetSeqNo + 1
        End If
        If GetSeqNo > tmpRng2 Then
            MainFrm.stsBar.Panels(2).Text = "접수순번의 범의 (" & tmpRng1 & "-" & tmpRng2 & ")를 벗어났읍니다. : " & GetSeqNo
            GoTo Err_Trap
        End If
        tmpSQL = objSeq.CreateSql_SeqNo(pWorkArea, pRcvDt, GetSeqNo, 2, tmpAccNo, GetSeqNo)
    End If
    Set tmpRs = Nothing

    DBConn.Execute tmpSQL
    DBConn.CommitTrans

    Exit Function

Err_Trap:
    DBConn.RollbackTrans
    Set tmpRs = Nothing
    Set objSeq = Nothing
    GetSeqNo = 0
    Exit Function

End Function


'% Clear 루틴
Sub ClearRtn()
    txtMesg.Text = ""
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblDeptNm.Caption = ""
    lblWard.Caption = ""
    lblSpcNm.Caption = ""
    lblStoreNm.Caption = ""
    lblLabNo.Caption = ""
    With tblOrdSheet
        .Row = -1
        .Col = -1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    lblMsg.Caption = ""
    
     Call ICSPatientMark
End Sub

Private Sub txtWorkArea_Change()
    If Len(txtWorkArea.Text) = txtWorkArea.MaxLength Then
        If txtAccDt.Enabled Then txtAccDt.SetFocus
    End If
End Sub

Private Sub txtWorkArea_GotFocus()
    With txtWorkArea
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWorkArea_KeyPress(KeyAscii As Integer)
    Call ClearRtn
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If txtWorkArea = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then txtAccDt.SetFocus
End Sub
