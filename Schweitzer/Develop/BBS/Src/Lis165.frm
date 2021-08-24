VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm165OutCol 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   5685
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin VB.Frame Frame3 
      Height          =   8655
      Left            =   0
      TabIndex        =   84
      Top             =   -160
      Width           =   4335
      Begin VB.Frame fraSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   645
         Left            =   40
         TabIndex        =   88
         Tag             =   "136"
         Top             =   960
         Width           =   4220
         Begin VB.TextBox txtSearchKey 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            MaxLength       =   10
            TabIndex        =   91
            Top             =   240
            Width           =   1830
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   2505
            TabIndex        =   90
            Tag             =   "15305"
            Top             =   285
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00DBE6E6&
            Caption         =   "&ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1995
            TabIndex        =   89
            Tag             =   "15304"
            Top             =   300
            Width           =   510
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  '단색
            Height          =   285
            Index           =   1
            Left            =   3465
            Shape           =   4  '둥근 사각형
            Top             =   255
            Width           =   675
         End
         Begin VB.Label lblReset 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  '투명
            Caption         =   "Reset"
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
            Left            =   3570
            MouseIcon       =   "Lis165.frx":0000
            MousePointer    =   99  '사용자 정의
            TabIndex        =   92
            Top             =   285
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   480
         Left            =   40
         TabIndex        =   85
         Tag             =   "136"
         Top             =   480
         Width           =   4220
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "채혈 대상자만"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   135
            TabIndex        =   87
            Tag             =   "15304"
            Top             =   165
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "전체 환자"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   2505
            TabIndex        =   86
            Tag             =   "15305"
            Top             =   150
            Width           =   1260
         End
      End
      Begin MSComctlLib.ListView lvwPtList 
         Height          =   6915
         Left            =   45
         TabIndex        =   93
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   12197
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16775406
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   285
         Left            =   80
         TabIndex        =   94
         Top             =   200
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   503
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
         Caption         =   "환자검색"
         LeftGab         =   100
      End
   End
   Begin VB.CommandButton cmdRRmkVisible 
      BackColor       =   &H00E0E0E0&
      Caption         =   "맞춤채혈등록 (&S)"
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
      Left            =   5775
      Style           =   1  '그래픽
      TabIndex        =   35
      Tag             =   "0"
      Top             =   8130
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame capFrame1 
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
      Height          =   2295
      Left            =   4425
      TabIndex        =   13
      Tag             =   "104"
      Top             =   285
      Width           =   9975
      Begin VB.CommandButton cmdRRmk 
         BackColor       =   &H00F7F3F8&
         Caption         =   "맞춤채혈확인"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2700
         Picture         =   "Lis165.frx":08CA
         Style           =   1  '그래픽
         TabIndex        =   71
         Top             =   210
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdContent 
         BackColor       =   &H00DBE6E6&
         Caption         =   "기존채혈내역"
         Height          =   345
         Left            =   8310
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   1185
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
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
         Height          =   330
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   17
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox txtReceptNo 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   16
         Top             =   555
         Width           =   1410
      End
      Begin VB.ComboBox cboOrdDate 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   210
         Width           =   2055
      End
      Begin VB.TextBox txtMesg 
         BackColor       =   &H00F7FDF8&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1290
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   14
         ToolTipText     =   "검사 리마크를 입력하세요."
         Top             =   1620
         Width           =   8250
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   4740
         TabIndex        =   18
         Top             =   210
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         BackColor       =   15857140
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
         Height          =   300
         Left            =   7575
         TabIndex        =   19
         Top             =   540
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   529
         BackColor       =   15857140
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
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   1575
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         BackColor       =   16249848
         ForeColor       =   0
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
         Caption         =   "◈ Remark"
      End
      Begin MedControls1.LisLabel lblDisease 
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Top             =   1200
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   556
         BackColor       =   15988216
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
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   150
         TabIndex        =   60
         Top             =   240
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
         Left            =   150
         TabIndex        =   61
         Top             =   555
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
         Caption         =   "영수증번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   3690
         TabIndex        =   62
         Top             =   210
         Width           =   1020
         _ExtentX        =   1799
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
         Index           =   8
         Left            =   3690
         TabIndex        =   63
         Top             =   540
         Width           =   1020
         _ExtentX        =   1799
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
         Height          =   315
         Index           =   11
         Left            =   6390
         TabIndex        =   64
         Top             =   210
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
         Caption         =   "처 방 일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   12
         Left            =   6390
         TabIndex        =   65
         Top             =   540
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   20
         Left            =   150
         TabIndex        =   66
         Top             =   1200
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
         Caption         =   "상 병 명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColID 
         Height          =   300
         Left            =   7575
         TabIndex        =   69
         Top             =   870
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   529
         BackColor       =   15857140
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
         Index           =   13
         Left            =   6390
         TabIndex        =   70
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
         Caption         =   "전문채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSang 
         Height          =   315
         Index           =   14
         Left            =   150
         TabIndex        =   82
         Top             =   860
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   255
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
         Caption         =   "감염주의"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDiseaseSang 
         Height          =   315
         Left            =   1320
         TabIndex        =   83
         Top             =   880
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "새굴림"
            Size            =   9
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
      Begin VB.Label lblSex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4740
         TabIndex        =   24
         Top             =   555
         Width           =   480
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5445
         TabIndex        =   23
         Top             =   555
         Width           =   480
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6150
         TabIndex        =   22
         Top             =   585
         Width           =   60
      End
      Begin VB.Label lblOrdDtCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   9675
         TabIndex        =   21
         Top             =   315
         Width           =   105
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F7F3F8&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   690
         Index           =   0
         Left            =   165
         Top             =   1560
         Width           =   9420
      End
      Begin VB.Label Label8 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         Caption         =   "        /"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4725
         TabIndex        =   25
         Top             =   540
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00CDC481&
      Caption         =   "채혈+접수(&A)"
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
      Index           =   1
      Left            =   7830
      Style           =   1  '그래픽
      TabIndex        =   81
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel lblRIMsg 
      Height          =   405
      Left            =   4455
      TabIndex        =   80
      Top             =   8520
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   714
      BackColor       =   14411494
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Alignment       =   1
      Caption         =   "핵의학 처방이 존재합니다."
      Appearance      =   0
   End
   Begin VB.CommandButton cmdRA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "핵의학(&R)"
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
      Left            =   9165
      Style           =   1  '그래픽
      TabIndex        =   79
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraPtRmk 
      BackColor       =   &H00F1F5F4&
      Height          =   5280
      Left            =   4425
      TabIndex        =   31
      Top             =   870
      Visible         =   0   'False
      Width           =   9570
      Begin VB.ComboBox cboRmk 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   68
         Text            =   "Combo1"
         Top             =   2175
         Width           =   5400
      End
      Begin VB.CommandButton cmdRSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "저장"
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
         Left            =   6840
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   4710
         Width           =   1320
      End
      Begin VB.TextBox txtRMesg 
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
         Height          =   1815
         Left            =   4095
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   44
         ToolTipText     =   "검사 리마크를 입력하세요."
         Top             =   2850
         Width           =   5370
      End
      Begin VB.TextBox txtRTitle 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4095
         MaxLength       =   50
         TabIndex        =   43
         Top             =   2520
         Width           =   5370
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5610
         Style           =   1  '그래픽
         TabIndex        =   40
         Tag             =   "PtID"
         Top             =   1515
         Width           =   315
      End
      Begin MedControls1.LisLabel lblRDeptNm 
         Height          =   330
         Left            =   5940
         TabIndex        =   38
         Top             =   1185
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   582
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin VB.CommandButton cmdHelpList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5610
         Style           =   1  '그래픽
         TabIndex        =   37
         Tag             =   "PtID"
         Top             =   1185
         Width           =   315
      End
      Begin VB.TextBox txtRDept 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4095
         MaxLength       =   10
         TabIndex        =   36
         Top             =   1185
         Width           =   1515
      End
      Begin VB.CommandButton cmdRClose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "닫기"
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
         Left            =   8175
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   4710
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   330
         Index           =   2
         Left            =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   150
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   582
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
         Caption         =   "환자 특이사항 등록 및 조회"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin MSComCtl2.DTPicker dtpEntDt 
         Height          =   330
         Left            =   4095
         TabIndex        =   34
         Top             =   855
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88997889
         CurrentDate     =   37505
      End
      Begin MedControls1.LisLabel lblRColNm 
         Height          =   315
         Left            =   5940
         TabIndex        =   41
         Top             =   1530
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRPtnm 
         Height          =   315
         Left            =   5940
         TabIndex        =   42
         Top             =   525
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblREntNm 
         Height          =   315
         Left            =   5940
         TabIndex        =   45
         Top             =   855
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin FPSpread.vaSpread tblRData 
         Height          =   4125
         Left            =   60
         TabIndex        =   46
         Top             =   525
         Width           =   2685
         _Version        =   196608
         _ExtentX        =   4736
         _ExtentY        =   7276
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   9
         MaxRows         =   17
         OperationMode   =   2
         ScrollBars      =   0
         SelectBlockOptions=   0
         ShadowColor     =   15663103
         SpreadDesigner  =   "Lis165.frx":0E54
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F1F5F4&
         Height          =   435
         Left            =   4095
         TabIndex        =   47
         Top             =   1740
         Width           =   1845
         Begin VB.OptionButton optExp 
            BackColor       =   &H00F1F5F4&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   1
            Left            =   990
            TabIndex        =   49
            Top             =   165
            Width           =   795
         End
         Begin VB.OptionButton optExp 
            BackColor       =   &H00F1F5F4&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   48
            Top             =   165
            Width           =   960
         End
      End
      Begin MedControls1.LisLabel lblRPtid 
         Height          =   315
         Left            =   4095
         TabIndex        =   51
         Top             =   525
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblRSeq 
         Height          =   300
         Left            =   5955
         TabIndex        =   52
         Top             =   1860
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BackColor       =   15857140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   2805
         TabIndex        =   53
         Top             =   525
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "환자정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   2805
         TabIndex        =   54
         Top             =   855
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "등록정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   2805
         TabIndex        =   55
         Top             =   1185
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "맞춤채혈부서"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   2805
         TabIndex        =   56
         Top             =   1515
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "맞춤 채혈자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   2805
         TabIndex        =   57
         Top             =   1845
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "폐기여부"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   2805
         TabIndex        =   58
         Top             =   2520
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "Title"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   1800
         Index           =   10
         Left            =   2805
         TabIndex        =   59
         Top             =   2850
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   3175
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   2805
         TabIndex        =   67
         Top             =   2175
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "Templete"
         Appearance      =   0
      End
      Begin VB.TextBox txtRColid 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4095
         MaxLength       =   10
         TabIndex        =   39
         Top             =   1515
         Width           =   1515
      End
      Begin VB.Shape Shape2 
         Height          =   4185
         Index           =   1
         Left            =   30
         Top             =   495
         Width           =   2745
      End
      Begin VB.Shape Shape2 
         Height          =   4185
         Index           =   0
         Left            =   2760
         Top             =   495
         Width           =   6735
      End
   End
   Begin VB.CheckBox chkPay 
      BackColor       =   &H00800000&
      Caption         =   "전체처방조회(미수납포함)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5850
      TabIndex        =   26
      Top             =   60
      Visible         =   0   'False
      Width           =   3345
   End
   Begin MedControls1.LisLabel lblBar 
      Height          =   315
      Left            =   4425
      TabIndex        =   1
      Top             =   2580
      Width           =   9960
      _ExtentX        =   17568
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
      Caption         =   "검체 채취 리스트"
      LeftGab         =   100
   End
   Begin VB.Frame fraOrder 
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
      Height          =   5655
      Left            =   4425
      TabIndex        =   5
      Top             =   2820
      Width           =   9990
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5070
         Left            =   45
         TabIndex        =   12
         Tag             =   "10114"
         Top             =   495
         Width           =   9885
         _Version        =   196608
         _ExtentX        =   17436
         _ExtentY        =   8943
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14936810
         GridColor       =   14737632
         MaxCols         =   36
         MaxRows         =   5
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis165.frx":1399
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   500
      End
      Begin VB.PictureBox picOrdDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2145
         ScaleHeight     =   300
         ScaleWidth      =   3690
         TabIndex        =   9
         Top             =   180
         Width           =   3690
         Begin VB.Shape shp1 
            BackColor       =   &H00553755&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   1
            Left            =   135
            Shape           =   3  '원형
            Top             =   60
            Width           =   330
         End
         Begin VB.Label lblLIS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "임상병리"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   420
            TabIndex        =   11
            Top             =   60
            Width           =   720
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00496835&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   2
            Left            =   1185
            Shape           =   3  '원형
            Top             =   60
            Width           =   330
         End
         Begin VB.Label lblBBS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "혈액은행"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   1485
            TabIndex        =   10
            Top             =   60
            Width           =   720
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00E0E0E0&
            Height          =   300
            Index           =   0
            Left            =   0
            Shape           =   4  '둥근 사각형
            Top             =   0
            Width           =   2385
         End
      End
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체 선택(&A)"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   1470
      End
      Begin VB.CheckBox chkChangeColTm 
         BackColor       =   &H00DBE6E6&
         Caption         =   "채혈시간변경 : "
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   300
         Left            =   6465
         TabIndex        =   6
         Top             =   180
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker dtpColDtTm 
         Height          =   300
         Left            =   8010
         TabIndex        =   8
         Top             =   165
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         CustomFormat    =   "yyyy-MM-dd H:mm"
         Format          =   88997891
         UpDown          =   -1  'True
         CurrentDate     =   36851.6291666667
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "채   혈 (&S)"
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
      Index           =   0
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   3
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종 료(&X)"
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
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   285
      Left            =   4425
      TabIndex        =   0
      Top             =   45
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   503
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
      Caption         =   "환자 기본정보"
      LeftGab         =   100
   End
   Begin VB.Frame fraContent 
      BackColor       =   &H00DBE6E6&
      Height          =   3975
      Left            =   4425
      TabIndex        =   29
      Top             =   1605
      Visible         =   0   'False
      Width           =   9585
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00DBE6E6&
         Height          =   3210
         Left            =   2910
         ScaleHeight     =   3150
         ScaleWidth      =   6555
         TabIndex        =   75
         Top             =   195
         Width           =   6615
         Begin MedControls1.LisLabel LisLabel7 
            Height          =   300
            Index           =   1
            Left            =   45
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   45
            Width           =   6480
            _ExtentX        =   11430
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
            Caption         =   "접수번호별 검사항목 리스트"
            Appearance      =   0
            LeftGab         =   200
         End
         Begin FPSpread.vaSpread tblLabno 
            Height          =   2745
            Left            =   45
            TabIndex        =   77
            Tag             =   "10114"
            Top             =   375
            Width           =   6465
            _Version        =   196608
            _ExtentX        =   11404
            _ExtentY        =   4842
            _StockProps     =   64
            AutoCalc        =   0   'False
            AutoClipboard   =   0   'False
            BackColorStyle  =   1
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FormulaSync     =   0   'False
            MaxCols         =   8
            MoveActiveOnFocus=   0   'False
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            ShadowText      =   0
            SpreadDesigner  =   "Lis165.frx":214E
            StartingColNumber=   2
            VirtualRows     =   24
            VisibleCols     =   5
            VisibleRows     =   500
         End
      End
      Begin VB.PictureBox fraLabNo 
         BackColor       =   &H00DBE6E6&
         Height          =   3210
         Left            =   105
         ScaleHeight     =   3150
         ScaleWidth      =   2730
         TabIndex        =   72
         Top             =   195
         Width           =   2790
         Begin MedControls1.LisLabel LisLabel7 
            Height          =   300
            Index           =   0
            Left            =   30
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   30
            Width           =   2670
            _ExtentX        =   4710
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
            Caption         =   "접수번호리스트"
            Appearance      =   0
            LeftGab         =   200
         End
         Begin MSComctlLib.ListView lvwLabNo 
            Height          =   2745
            Left            =   15
            TabIndex        =   74
            Top             =   360
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   4842
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "접수번호"
               Object.Width           =   2824
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "상태"
               Object.Width           =   1766
            EndProperty
         End
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "확인"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   8190
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   3405
         Width           =   1320
      End
   End
   Begin MedControls1.LisLabel lblRIFlag 
      Height          =   375
      Left            =   4440
      TabIndex        =   78
      Top             =   8265
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      BackColor       =   15857140
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
End
Attribute VB_Name = "frm165OutCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private objPatient  As clsPatient
Private objSql      As clsLISSqlCollection
Private objCollect  As clsLISCollectioin

Private blnOrdFg    As Boolean
Private blnSelAllFg As Boolean
Private blnMsgFg    As Boolean
Private blnClearFg  As Boolean

Private lngBackEven As Long
Private lngBackOdd  As Long
Private lngForeEven As Long
Private lngForeOdd  As Long

Private strDeptCd   As String

Private blnInitFg   As Boolean
Private mvarPtID    As String
Private mvarOrddt   As String

Private strWrkDiv   As String

Private Const lngMaxRows = 20
Private Const lngRowHeight = 12.5

Public Event LastFormUnload()

Public Property Let Ptid(ByVal vData As String)
    mvarPtID = vData
End Property

Public Property Let OrdDt(ByVal vData As String)
    mvarOrddt = vData
End Property

Private Sub cboOrdDate_Click()
   
    '** 예수병원 추가 변수========
    Dim Message     As String
    Dim tmpDate     As String
    Dim strTemp     As String
    '=============================
       
    If txtPtId.Text = "" Then Exit Sub
    If Screen.ActiveForm.name <> Me.name Then Exit Sub
    
    '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
    '=====================================================================================
    If cboOrdDate.ListIndex > -1 Then
        tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
        strTemp = objSql.RI_Collection_Flag(txtPtId.Text, tmpDate)
        
        lblRIFlag.Caption = medGetP(strTemp, 1, COL_DIV)
        strWrkDiv = medGetP(strTemp, 2, COL_DIV)
        If lblRIFlag.Caption = "Y" Then
            lblRIMsg.Caption = "핵의학 처방이 있습니다."
        Else
            lblRIMsg.Caption = ""
        End If
    Else
        lblRIFlag.Caption = "": strWrkDiv = ""
        lblRIMsg.Caption = ""
    End If
    '=====================================================================================
    
    Call DisplayOrder
    
    If blnOrdFg Then
        cmdSave(0).Enabled = True
        cmdSave(1).Enabled = True
        tblOrdSheet.SetFocus
    Else
        If cboOrdDate.ListCount > 1 Then
            With tblOrdSheet
                .Row = 1: .Row2 = .MaxRows
                .Col = 1: .COL2 = .MaxCols
                .BlockMode = True
                .Action = ActionClearText
                .BlockMode = False
            End With
            cboOrdDate.SetFocus
            
            '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
            '=====================================================================================
            If lblRIFlag.Caption = "Y" Then
                '-- 핵의학 채취화면 Call 여부
                Message = MsgBox("핵의학 검사가 존재합니다. 계속진행하시겠습니까?", vbExclamation + vbYesNo, "채취확인")
                
                If Message = vbYes Then
                    If strWrkDiv = "3" Or strWrkDiv = "4" Then
                        '** 핵의학 화면 Call
                        Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
                    Else
                        '** 핵의학 화면 Call
                        Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
                        '사용자ID : ObjSysInfo.EmpId
                    End If
                End If
            End If
            '=====================================================================================
        Else
            '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
            '=====================================================================================
            If lblRIFlag.Caption = "Y" Then
                '-- 핵의학 채취화면 Call 여부
                Message = MsgBox("핵의학 검사가 존재합니다. 계속진행하시겠습니까?", vbExclamation + vbYesNo, "채취확인")
                
                If Message = vbYes Then
                    If strWrkDiv = "3" Or strWrkDiv = "4" Then
                        '** 핵의학 화면 Call
                        Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
                    Else
                        '** 핵의학 화면 Call
                        Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
                    End If
                    
                End If
            End If
            '=====================================================================================
            
            Call ClearRtn
            txtPtId.Text = ""
            txtPtId.SetFocus
            
        End If
    End If

End Sub

Private Sub cboRmk_Click()
    txtRMesg.Text = ""
    If cboRmk.ListIndex <> 0 Then
        txtRMesg.Text = medGetP(cboRmk.Text, 1, vbTab)
    End If
End Sub


Private Sub chkPay_Click()
    If txtPtId.Text <> "" Then
        Call txtPtId_LostFocus
    End If
End Sub

Private Sub chkSelAll_Click()
    Dim ii  As Integer
    
    blnSelAllFg = True
    With tblOrdSheet
        .Col = 1: .COL2 = 1
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .value = chkSelAll.value
        .BlockMode = False
        If P_PayDtUsed Then
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = enCOLLIST.tcPAYDT
                If .value = "" Then
                    .Col = enCOLLIST.tcCHECK: .value = 0
                End If
            Next
        End If
    End With
    
    
    blnSelAllFg = False
   
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn
    txtPtId.Text = ""
    txtReceptNo.Text = ""
    txtPtId.SetFocus
End Sub

Private Sub cmdContent_Click()
    Call LastCollectFg(True)
End Sub

Private Sub cmdExit_Click()
    Unload Me
'    If IsLastForm Then RaiseEvent LastFormUnload
    Set objPatient = Nothing
    Set objSql = Nothing
    Set objCollect = Nothing
End Sub

Private Sub cboOrdDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cboOrdDate_Click
End Sub

Private Sub cmdOK_Click()
    fraContent.Visible = False
End Sub

Private Sub cmdRA_Click()
    Dim tmpDate     As String
    
    '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
    '=====================================================================================
    If lblRIFlag.Caption = "Y" Then
        tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
        
        If strWrkDiv = "3" Or strWrkDiv = "4" Then
            '** 핵의학 화면 Call
            Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
        Else
            '** 핵의학 화면 Call
            Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
            '사용자ID : ObjSysInfo.EmpId
        End If
                                
        'Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
    End If
    '=====================================================================================
End Sub

Private Sub cmdSave_Click(Index As Integer)
    Dim objPrgBar       As clsProgress
    Dim objDIC          As clsDictionary
    Dim BBSColSuccess   As Boolean
    Dim LISColSuccess   As Boolean
    
    Dim iCheckOrder     As Integer

    Dim ii              As Integer
    
    If CollectionTargetChk = False Then
'       MsgBox "채혈할 항목을 선택하세요..", vbInformation, "항목선택"
       tblOrdSheet.SetFocus
       Exit Sub
    End If
    
    iCheckOrder = objCollect.CheckSameOrder(tblOrdSheet, 1)     '중복처방 Check
    If iCheckOrder > 0 Then GoTo OrdCheck1
    
    Call MouseRunning
     
    Set objPrgBar = New clsProgress
    With objPrgBar
        .Container = Me
        .Left = lblBar.Left + 5
        .Top = lblBar.Top + 5
        .Width = lblBar.Width - 10
        .Height = lblBar.Height - 10
        .Message = "선택된 검사항목에 대해 채혈처리중입니다..."
        .Max = 90
        .value = 10
        DoEvents
    End With

    DoEvents

    '----------------------------------------------------------
    '업무별 구분을 위해서 업무별로 불럭을 구분한다.(2001/06/08)
    '----------------------------------------------------------

    Set objDIC = New clsDictionary
    objDIC.Clear
    objDIC.FieldInialize "orddiv", "first,last,coldt,coltm"
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = enCOLLIST.tcORDDIV
            Select Case .value
                Case BBS_ORDDIV
                    If objDIC.Exists(.value) Then
                        objDIC.KeyChange BBS_ORDDIV
                        objDIC.Fields("last") = .Row
                    Else
                        .Col = enCOLLIST.tcREQDTTM
                        objDIC.AddNew BBS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & _
                                      Format(.Text, "yyyymmdd") & COL_DIV & Format(.Text, "HHmm")
                    End If
                Case LIS_ORDDIV
                    If objDIC.Exists(.value) Then
                        objDIC.KeyChange LIS_ORDDIV
                        objDIC.Fields("last") = .Row
                    Else
                        objDIC.AddNew LIS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & "" & COL_DIV & ""
                    End If
            End Select
        Next
        objDIC.MoveFirst
        Do Until objDIC.EOF
            If objDIC.Fields("last") = "" Then
                objDIC.Fields("last") = objDIC.Fields("first")
            End If
            objDIC.MoveNext
        Loop
    End With
    With objDIC
        .MoveFirst
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case LIS_ORDDIV: iCheckOrder = objCollect.ChkSpcnm(tblOrdSheet, .Fields("first"), .Fields("last"))
            End Select
            If iCheckOrder > 0 Then GoTo OrdCheck2
            .MoveNext
        Loop
    End With
    
    '-------------------------------------------------------------
    '업무별로 채혈을 수행한다(혈액은행은 지정검체 체크가 필요없음)
    '-------------------------------------------------------------
    With objDIC
        .MoveFirst
        BBSColSuccess = True:  LISColSuccess = True
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case BBS_ORDDIV: BBSColSuccess = CollectForBBS(.Fields("first"), .Fields("last"), _
                                                                    Format(GetSystemDate, "yyyymmdd"), _
                                                                    Format(GetSystemDate, "HHmmss"), objPrgBar)
                Case LIS_ORDDIV: LISColSuccess = CollectForLIS(.Fields("first"), .Fields("last"), objPrgBar, Index)
            End Select
            .MoveNext
        Loop
    End With
    
    If Not BBSColSuccess And LISColSuccess Then
        Set objPrgBar = Nothing
        MsgBox "채혈처리중 오류가 발생했습니다 !!" & vbCrLf & _
               "재실행하신 후 오류가 계속되면 전산실 혹은 임상병리과로 연락바랍니다.", _
               vbCritical, "오류"
    End If
    
    Call MouseDefault
    Set objPrgBar = Nothing
    Set objDIC = Nothing
ExitPos:
    '** 예수병원 추가 루틴 By M.G.Choi 2004.11.09
    '=====================================================================================
    If lblRIFlag.Caption = "Y" Then
        Dim tmpDate     As String
        
        tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
        
        If strWrkDiv = "3" Or strWrkDiv = "4" Then
            '** 핵의학 화면 Call
            Call Shell("C:\uniHIS\EXE\MCOLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
        Else
            '** 핵의학 화면 Call
            Call Shell("C:\uniHIS\EXE\COLLECT.EXE" & " " & txtPtId.Text & " " & tmpDate & " " & ObjSysInfo.EmpId & " " & "1", vbNormalFocus)
            '사용자ID : ObjSysInfo.EmpId
        End If
        
    End If
    '=====================================================================================
    
    '-- Clear Modify By M.G.Choi 200.06.28
    Call cboOrdDate_Click
    
    If tblOrdSheet.DataRowCnt < 1 Then
        Call cmdClear_Click
    End If
    txtPtId.SetFocus
    Set objDIC = Nothing
    
    Exit Sub

OrdCheck1:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "중복처방입니다. 확인하고 다시 채혈하십시오.", vbExclamation, "중복처방"
    tblOrdSheet.SetFocus
    
    Exit Sub

OrdCheck2:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "지정검체 정보가 없습니다. 전산실 혹은 임상병리과로 연락하세요.", vbInformation + vbOKOnly, "오류"
    tblOrdSheet.SetFocus
    Set objDIC = Nothing
    Exit Sub

End Sub


'** 혈액은행 채혈루틴
Private Function CollectForBBS(ByVal FRowCnt As Integer, ByVal LRowCnt As Integer, _
                                ByVal ColDt As String, ByVal ColTm As String, _
                                ByRef objProgress As Object) As Boolean

    
    Dim dicBBS      As clsDictionary
    Dim objBAR      As clsDictionary
    Dim objCollect  As clsBBSCollection
    
    Dim tmpClipData As String
    Dim strStatFg   As String
    
    Dim tmpTotData  As Variant
    Dim tmpRowData  As Variant
    
    Dim lngColCnt   As Integer
    
    Dim i           As Long
    
    lngColCnt = 0
    
    Set objCollect = New clsBBSCollection
    
    If objCollect.Blood_Existence(txtPtId.Text, Format(GetSystemDate, "yyyymmdd"), Format(GetSystemDate, "hhmmss")) = False Then
        If objCollect.SetAccessCheck(txtPtId.Text) = True Then
           '검체가 이미 존재하는 경우
           CollectForBBS = objCollect.SetWardAccess(txtPtId.Text, enBussDiv.BussDiv_OutPatient, Format(GetSystemDate, "yyyymmdd"), _
                                    Format(GetSystemDate, "hhmmss"), ObjSysInfo.EmpId)
                
            Set objCollect = Nothing
            Exit Function
        Else
            GoTo AutoCollect
        End If
    End If
    
AutoCollect:

    Set dicBBS = New clsDictionary
    Set objBAR = New clsDictionary
'    Set objCollect = New clsBBSCollection
    
    With tblOrdSheet
        .Col = 1:                  .COL2 = .MaxCols
        .Row = FRowCnt:            .Row2 = LRowCnt:                         .BlockMode = True
        tmpClipData = .ClipValue:  tmpTotData = Split(tmpClipData, vbCrLf): .BlockMode = False
        
        .Col = 7: strStatFg = IIf(Trim(.value) = "Y", "1", "0")
        
        For i = 0 To UBound(tmpTotData) - 1

            tmpRowData = Split(tmpTotData(i), vbTab)
            If objProgress.Max > objProgress.value Then objProgress.value = objProgress.value + 1
            If tmpRowData(0) = 0 Then GoTo Skip       '선택여부
          
            lngColCnt = lngColCnt + 1
            
            '혈액은행-----------------------------------------------------------------------------
                
                dicBBS.Clear
                dicBBS.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"
                dicBBS.AddNew txtPtId.Text, Join(Array(lblPtNm.Caption, ColDt, ColTm, _
                              ObjSysInfo.EmpId, enBussDiv.BussDiv_OutPatient, ObjSysInfo.BuildingCd, "", strStatFg), COL_DIV)
' gEmpid 를 ObjSysInfo.EmpId로 대체함.
Skip:
       Next
    
    End With
    
    If lngColCnt = 0 Then
        CollectForBBS = True
        Exit Function
    End If
          
    CollectForBBS = objCollect.Set_Collect(dicBBS, , objProgress)
    
    If CollectForBBS Then
        Set objBAR = objCollect.BldDic
        If objBAR.RecordCount > 0 Then
        '바코드 출력
            BarCodePrintForBBS objBAR
        Else
            If objCollect.CheckCol Then
                MsgBox "정상적으로 처리되지 않았습니다.", vbExclamation
            Else
                MsgBox "수혈처방 검체가 이미 존재하므로 바코드가 출력되지 않습니다.", vbInformation + vbOKOnly, "바코드출력"
            End If
        End If
        
        If objCollect.Spc72Chk Then
            MsgBox "해당 환자는 72시간내에 채혈한 검체가 존재합니다.", vbInformation + vbOKOnly, "바코드출력"
        End If
    End If
    
    Set objBAR = Nothing
    Set dicBBS = Nothing
    Set objCollect = Nothing
'    Set objProgress = Nothing
End Function

Private Function CollectForLIS(ByVal FRowCnt As Long, _
                               ByVal LRowCnt As Long, _
                               ByRef objProgress As Object, _
                               Optional pINx As Integer = 0) As Boolean
    Dim tmpData()   As String
    Dim i           As Integer
    Dim SelCount    As Integer
    Dim CollectCnt  As Integer
    
    Dim ColSuccess  As Boolean
    
    CollectCnt = 0
    Call objCollect.InitRtn

    With tblOrdSheet

        ReDim tmpData(0 To 20)
        For i = FRowCnt To LRowCnt
            
            If objProgress.Max > objProgress.value Then objProgress.value = objProgress.value + 1
            
            .Row = i
            
            .Col = enCOLLIST.tcCHECK
            If .value <> 1 Then GoTo Skip

            CollectCnt = CollectCnt + 1
            .Col = enCOLLIST.tcBUILDCD:  tmpData(0) = .value        'Delivery Location
            .Col = enCOLLIST.tcWORKAREA: tmpData(1) = .value        'WorkArea
            .Col = enCOLLIST.tcSPCCD:    tmpData(2) = .value        'SpcCd
            .Col = enCOLLIST.tcSTORECD:  tmpData(3) = .value        'StoreCd
            .Col = enCOLLIST.tcSTATFLAG: tmpData(4) = .value        'StatFg
            .Col = enCOLLIST.tcREQDTTM:  tmpData(5) = .value        'ReqColDate

            .Col = enCOLLIST.tcTESTDIV:  tmpData(6) = .value        'TestDiv
            .Col = enCOLLIST.tcMULTIFG:  tmpData(7) = .value        'MultiFg
            .Col = enCOLLIST.tcSPCGRP:   tmpData(8) = .value        'SpcGrp
            .Col = enCOLLIST.tcORDDATE:  tmpData(9) = .value        'OrdDt
            .Col = enCOLLIST.tcORDNUM:   tmpData(10) = .value       'OrdNo
            .Col = enCOLLIST.tcORDSEQ:   tmpData(11) = .value       'OrdSeq
            .Col = enCOLLIST.tcTESTCD:   tmpData(12) = .value       'OrdCd
            .Col = enCOLLIST.tcDEPTCD:   tmpData(13) = .value       '진료과
            .Col = enCOLLIST.tcORDDOCT:  tmpData(14) = .value       '처방의
            .Col = enCOLLIST.tcMAJDODT:  tmpData(15) = .value       '주치의
            .Col = enCOLLIST.tcABBRNM:   tmpData(16) = .value       '검사 약어명
            .Col = enCOLLIST.tcBARCNT:   tmpData(17) = .value       '라벨출력장수
            .Col = enCOLLIST.tcLABDIV:   tmpData(18) = .value       'LabDiv
            .Col = enCOLLIST.tcSPCABBR:  tmpData(19) = .value       '검체약어명
            .Col = enCOLLIST.tcLABRANGE: tmpData(20) = .value       '미생물접수번호범위
            
            Call objCollect.AddLabCollect(tmpData)

Skip:
        Next
    End With

    If CollectCnt = 0 Then
        CollectForLIS = True
        Exit Function
    End If

    With objCollect

        ReDim tmpData(0 To 16)

        tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)  '검체년도
        tmpData(1) = objPatient.Ptid                            '환자ID
        
' 2008.10.24 전염성 보균자일경우 바코드에 별을 붙이는 기능 추가.

        If Len(lblDiseaseSang.Caption) > 5 Then
            tmpData(2) = "*" & Trim(objPatient.ptnm)
        Else
            tmpData(2) = objPatient.ptnm
        End If
        
        tmpData(3) = objPatient.Sex                             '성별
        If IsDate(Format(objPatient.DOB, CS_DateLongMask)) Then                         '환자일령
            tmpData(4) = DateDiff("y", Format(objPatient.DOB, CS_DateLongMask), GetSystemDate)
        Else
            tmpData(4) = Mid(objPatient.DOB, 1, 4) & "-01-01"
            If IsDate(tmpData(4)) Then
                tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
            Else
                tmpData(4) = 0
            End If
        End If
        tmpData(5) = ""                                          '입원일
        tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)  '입력일
        tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)  '입력시간
        tmpData(8) = ObjSysInfo.EmpId                            '입력자
        tmpData(9) = ""                                          '원접수번호
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat) '채혈일
        .ColTm = Format(GetSystemDate, "HHMMSS")             '채혈일
        tmpData(11) = ObjSysInfo.EmpId                           '채혈자
        tmpData(12) = ""                                         '병동ID
        tmpData(13) = ""                                         '병실ID
        tmpData(14) = ""                                         '침상ID
        tmpData(15) = ""                                         '침상ID
        tmpData(16) = ObjSysInfo.BuildingCd                      '** 채혈이 수행되는 건물코드
        
        

        Call .SetColData(tmpData)
    End With
    ' 채혈 수행
    ColSuccess = objCollect.DoCollection(objProgress)
    
    '** 접수수행:장비PCx(혈당측정) 도입에 따라 채혈-접수 루틴 필요 (외래채혈실에서 바로 결과등록 하기 위함)
    '   추가 By M.G.Choi 2007.04.02
    '---------------------------------------------------------------------------------------------------------------------
    If ColSuccess = True And pINx = 1 Then
        objProgress.Message = "접수 Procedure를 수행하고 있습니다."
        Dim objAccess   As New clsLISAccession
        
        With objCollect
            If .CollectDone Then
                Dim pWorkArea As String
                Dim pAccDt As String
                Dim pAccSeq As Long
                
                For i = 1 To .ColCount
                    objProgress.Message = "접수 Procedure를 수행하고 있습니다. (" & CStr(i) & "/" & CStr(.ColCount) & ")"
                    Call .GetLabNumbers(i, pWorkArea, pAccDt, pAccSeq)
                    ColSuccess = objAccess.DoAccession(pWorkArea, pAccDt, pAccSeq, ObjMyUser.EmpId)
                    If Not ColSuccess Then Exit For
                    If objProgress.value = objProgress.Max Then objProgress.Max = objProgress.Max + 1
                    objProgress.value = objProgress.value + 1
                    DoEvents
                Next
            End If
        End With
        Set objAccess = Nothing
    End If
    '----------------------------------------------------------------------------------------------------------------------
    
    If Not ColSuccess Then
        Set objProgress = Nothing
        MsgBox "채혈처리중 오류가 발생했습니다 !!"
        MouseDefault  '0
        CollectForLIS = False
        Exit Function
    End If
    CollectForLIS = True

End Function

Private Sub BarCodePrintForBBS(objDIC As clsDictionary)
    Dim objSql      As New clsBBSCollection
    Dim objBAR      As New clsBarcode
    Dim strPtid     As String
    Dim strPtnm     As String
    Dim strColDt    As String
    Dim strColTm    As String
    Dim strSpcNo    As String
    Dim strW_Dept   As String
    Dim strAccSeq   As String         'SpcYy-SpcNo 형태의 검체번호
    Dim strHosilid  As String
    Dim strStatFg   As String
    
    Set objSql = New clsBBSCollection
    Set objBAR = New clsBarcode
    
'    Set objBAR.MyDB = dbconn
    Set objBAR.TableInfo = New clsTables
    Set objBAR.FieldInfo = New clsFields
    strW_Dept = strDeptCd
    
    objDIC.MoveFirst
    Do Until objDIC.EOF
        strPtid = medGetP(objDIC.GetString, 1, COL_DIV)
        strPtnm = medGetP(objDIC.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDIC.GetString, 3, COL_DIV)
'        strColDt = Mid(medGetP(objDIC.GetString, 4, COL_DIV), 1, 4)
        strColDt = Format(Mid(medGetP(objDIC.GetString, 4, COL_DIV), 5), "0#/0#")
        strColTm = Mid(medGetP(objDIC.GetString, 5, COL_DIV), 1, 4)
        strStatFg = medGetP(objDIC.GetString, 7, COL_DIV)
        strColTm = Format(strColTm, "0#:##")
        
        '검체번호 출력 : 2001.2.8 추가
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '
        objBAR.Label_PrintOut BBSName, "XM", "", strAccSeq, strSpcNo, strPtid, _
                              strPtnm, "", "", strStatFg, strW_Dept, strColDt, strColTm, _
                              "", 1
        objDIC.MoveNext
    Loop
    Set objBAR = Nothing
    Set objSql = Nothing

End Sub

Private Sub Form_Activate()
    If blnInitFg = False Then
        txtReceptNo.Visible = P_UseReceptForSearch
        lblReceptNo.Visible = P_UseReceptForSearch
    
        blnSelAllFg = False
        blnMsgFg = False
        blnClearFg = True
        optSort(1).value = True
        
        Call ClearRtn
' 08.09.25 양성현 생년월일 대신 집주소로 변겅
'        medInitLvwHead lvwPtList, "환자ID,환자성명,주민등록번호,생년월일,성별/나이,연락처", "100,100,300,300,100,800"
        medInitLvwHead lvwPtList, "환자ID,성 명,주민등록번호,주 소,성별/나이,연락처", "200,100,700,5000,100,700"
        
        Call tblOrdSheet.GetOddEvenRowColor(lngBackOdd, lngForeOdd, lngBackEven, lngForeEven)
    End If
    
    blnInitFg = True
    txtPtId.SetFocus

End Sub

Private Sub Form_Load()
    
    blnInitFg = False

'    cmdSave.Caption = "채 혈(&S)"
    chkPay.Visible = True
    Set objPatient = New clsPatient
    Set objSql = New clsLISSqlCollection
    Set objCollect = New clsLISCollectioin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    mvarPtID = "": mvarOrddt = ""
    Set objMyList = Nothing
    Set objPatient = Nothing
    Set objSql = Nothing
    Set objCollect = Nothing
End Sub

Private Sub lblReset_Click()
    lvwPtList.ListItems.Clear
    txtSearchKey.Text = ""
End Sub

Private Sub lvwPtList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static lngOrder As Long
    With lvwPtList
        lngOrder = (lngOrder + 1) Mod 2
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = Choose(lngOrder + 1, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '환자정보 Display
    If Item = "" Then Exit Sub
    DoEvents
    With Item
        txtPtId.Text = Trim(.Text)                '환자ID
        txtPtId.SetFocus
        Call txtPtId_KeyPress(vbKeyReturn)
    End With
    
End Sub

Private Sub optOption_Click(Index As Integer)
    lvwPtList.ListItems.Clear
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    
    Dim ButtonValue As Variant
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim i           As Integer
    
    If blnSelAllFg Then Exit Sub
    
    blnSelAllFg = True
    
    With tblOrdSheet
      
        If Not blnOrdFg Then Exit Sub
        If Col <> 1 Then Exit Sub
      
        .Row = Row
        .Col = Col:   ButtonValue = .value
        If P_PayDtUsed Then ' And Mid(ObjSysInfo.ProjectId, 1, 1) = LIS_ORDDIV Then
            .Col = enCOLLIST.tcPAYDT
            If .value = "" Then
                .Col = Col
                .value = Val(Val(.value + 1) Mod 2)
                If .value = 0 Then GoTo Skip
                Exit Sub
            End If
        End If
        
        If .value = 0 Then GoTo Skip
        
        .Col = 9:   SvOrdDt = .value
        .Col = 10:  SvOrdNo = .value
        
        For i = 1 To .DataRowCnt
            If i <> Row Then
                .Row = i
                .Col = 9
                If .value = SvOrdDt Then
                    .Col = 10
                    If .value = SvOrdNo Then
                        .Col = 1
                        If .value <> ButtonValue Then .value = ButtonValue
                    End If
                End If
            End If
        Next
    End With
    
Skip:
    blnSelAllFg = False

End Sub

Private Sub LastCollectFg(Optional ByVal blnClick As Boolean = False)
    Dim RS   As Recordset
    Dim SSQL As String
    Dim itmX As ListItem
    
    cmdContent.Visible = False
    
    SSQL = " SELECT workarea,accdt,accseq,coldt,coltm,stscd,testdiv FROM " & T_LAB201 & _
           " WHERE " & _
               DBW("ptid=", txtPtId.Text) & _
           " AND " & DBW("stscd<>", enStsCd.StsCd_LIS_Cancel) & _
           " AND " & DBW("coldt>=", Format(DateAdd("d", -30, GetSystemDate), "YYYYMMDD")) & _
           " AND " & DBW("coldt<=", Format(GetSystemDate, "YYYYMMDD")) & _
           " ORDER BY coldt desc,coltm desc"
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        cmdContent.Visible = True
        If Not blnClick Then
            Set RS = Nothing
            Set itmX = Nothing
            Exit Sub
        End If
        lvwLabNo.ListItems.Clear
        Do Until RS.EOF
            Set itmX = lvwLabNo.ListItems.Add(, , RS.Fields("workarea").value & "" & "-" & _
                                                  RS.Fields("accdt").value & "" & "-" & _
                                                  RS.Fields("accseq").value & "")
            Select Case RS.Fields("stscd").value & ""
                Case enStsCd.StsCd_LIS_Order:       itmX.SubItems(1) = STS_LIS_Order
                Case enStsCd.StsCd_LIS_Collection:  itmX.SubItems(1) = STS_LIS_HaveSpc
                Case enStsCd.StsCd_LIS_Accession:   itmX.SubItems(1) = STS_LIS_Access
                Case enStsCd.StsCd_LIS_InProcess:   itmX.SubItems(1) = STS_LIS_Worksheet
                Case enStsCd.StsCd_LIS_MidRst:
                                                If RS.Fields("testdiv").value & "" <> enTestDiv.TST_MicTest Then
                                                    itmX.SubItems(1) = STS_LIS_Partial
                                                Else
                                                    itmX.SubItems(1) = STS_LIS_MidRst
                                                End If
                Case enStsCd.StsCd_LIS_FinRst:
                                                If RS.Fields("testdiv").value & "" <> enTestDiv.TST_MicTest Then
                                                    itmX.SubItems(1) = STS_LIS_Verify
                                                Else
                                                    itmX.SubItems(1) = STS_LIS_FinRst
                                                End If
                Case enStsCd.StsCd_LIS_Modify:      itmX.SubItems(1) = STS_LIS_Modify
            End Select
            RS.MoveNext
        Loop
                        '
        fraContent.Visible = True
        fraContent.ZOrder
        Set itmX = lvwLabNo.SelectedItem
        
        Call lvwLabNo_ItemClick(itmX)
    End If
    Set RS = Nothing
    Set itmX = Nothing
End Sub
Private Sub lvwLabNo_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sWorkarea   As String
    Dim sAccdt      As String
    Dim sAccSeq     As String
    Dim sTestcd     As String
    Dim strTmp      As String
    Dim SSQL        As String
    Dim RS          As Recordset
    
    DoEvents

    Call medClearTable(tblLabno)
    
    strTmp = Item.Text
    
    sWorkarea = medGetP(strTmp, 1, "-")
    sAccdt = medGetP(strTmp, 2, "-")
    sAccSeq = medGetP(strTmp, 3, "-")
    
    SSQL = " SELECT a.orddt,b.coldt,b.coltm,b.rcvdt,b.rcvtm,b.vfydt,b.vfytm,c.abbrnm10,d.field3 as spcnm" & _
           " FROM " & T_LAB102 & " a," & T_LAB201 & " b," & T_LAB001 & " c," & T_LAB032 & " d" & _
           " WHERE " & _
                     DBW("b.workarea=", sWorkarea) & _
           " AND " & DBW("b.accdt=", sAccdt) & _
           " AND " & DBW("b.accseq=", sAccSeq) & _
           " AND " & DBW("d.cdindex=", LC3_Specimen) & _
           " AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
           " AND c.testcd=a.ordcd" & _
           " AND c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
           "                   WHERE testcd = c.testcd ) " & _
           " AND d.cdval1=a.spccd"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        With tblLabno
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1:   .value = Format(RS.Fields("orddt").value & "", "####-##-##")
                .Col = 2:   .value = RS.Fields("abbrnm10").value & ""
                .Col = 3:   .value = RS.Fields("spcnm").value & ""
                .Col = 4:   .value = Format(RS.Fields("coldt").value & "", "####-##-##")
                            If RS.Fields("coltm").value & "" <> "" Then
                                .value = .value & " " & Format(Mid(RS.Fields("coltm").value & "", 1, 4), "0#:##")
                            End If
                .Col = 5:
                            If RS.Fields("rcvdt").value & "" <> "" Then
                                .value = Format(RS.Fields("rcvdt").value & "", "####-##-##")
                            End If
                            If RS.Fields("rcvtm").value & "" <> "" Then
                                .value = .value & " " & Format(Mid(RS.Fields("rcvtm").value & "", 1, 4), "0#:##")
                            End If
                .Col = 6:
                            If RS.Fields("vfydt").value & "" <> "" Then
                                .value = Format(RS.Fields("vfydt").value & "", "####-##-##")
                            End If
                            If RS.Fields("vfytm").value & "" <> "" Then
                                .value = .value & " " & Format(Mid(RS.Fields("vfytm").value & "", 1, 4), "0#:##")
                            End If
                RS.MoveNext
            Loop
        End With
    End If
    Set RS = Nothing
End Sub

Private Sub txtPtId_LostFocus()
    Dim lngCnt As Long
    Dim lblDiseaseSangT As String
    
    If txtPtId.Text = "" Then Exit Sub
    If Screen.ActiveForm.name <> Me.name Then Exit Sub
   
    Call ClearRtn
    With objPatient
        If IsNumeric(txtPtId.Text) Then
            txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
        End If
        
        Call ICSPatientMark(txtPtId.Text, enICSNum.LIS_ALL)
        
        If .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .ptnm         '성명
            lblSex.Caption = .SEXNM         '성별
            lblAge.Caption = .Age           '연령
            lblAgeDiv.Caption = .AGEDIV     '나이단위
            lblDeptNm.Caption = .DeptNm     '진료과
            blnClearFg = False
            
            If chkPay.value = 1 Then
                lngCnt = objCollect.LoadOrderDate(txtPtId.Text, cboOrdDate)
            Else
                lngCnt = objCollect.LoadOrderDateYesPay(txtPtId.Text, cboOrdDate)
            End If
            
            If lngCnt <= 0 Then
                blnMsgFg = True
                MsgBox objPatient.ptnm & " 님의 처방내역이 없습니다", vbExclamation, "외래채혈"
                txtPtId.Text = ""
                txtPtId.SetFocus
                blnMsgFg = False
                Call txtPtId_GotFocus
                Exit Sub
            Else
                Call LastCollectFg
                Call GetPtRmkVisibleTrueFalse
                
                If tblOrdSheet.DataRowCnt <> 0 Then
                    lblOrdDtCnt.Caption = CStr(lngCnt)
                    cboOrdDate.ListIndex = 0
                End If
            End If
            
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.Ptid = txtPtId.Text
            lblDisease.Caption = objDisease.Disease

'========================================================
' 08.10.24. 양성현 전주 예수병원 요청사항.
'========================================================
            
            lblDiseaseSangT = objDisease.DiseaseSang
            If Len(lblDiseaseSangT) > 5 Then
                lblSang(14).Visible = True
                lblSang(14).Refresh
                lblDiseaseSang.Visible = True
                lblDiseaseSang.Caption = lblDiseaseSangT
                
'            lblDiseaseSang.ForeColor = 1
            End If
' -------------------------------------------------------
            
            Set objDisease = Nothing
        Else
            blnMsgFg = True
            txtPtId.Text = ""
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요..", vbExclamation, "외래채혈"
            txtPtId.SetFocus
            blnMsgFg = False
            Call txtPtId_GotFocus
            Exit Sub
        End If
    End With

End Sub

Private Sub txtReceptNo_Change()
   
    If Not blnClearFg Then
       txtPtId.Text = ""
    End If

End Sub

Private Sub txtReceptNo_GotFocus()
   
    With txtReceptNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtReceptNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtReceptNo_LostFocus()
   
    If txtReceptNo.Text = "" Then Exit Sub
    
    Dim tmpRs As Recordset
    
    Set tmpRs = New Recordset
    tmpRs.Open objSql.SqlLoadOrderDate(txtReceptNo.Text, 1), DBConn
    
    If tmpRs.EOF Then
        MsgBox "해당 영수증번호는 존재하지 않습니다.", vbExclamation, "외래채혈"
        txtReceptNo.SetFocus
        GoTo NoData
    End If
    
    txtPtId.Text = "" & tmpRs.Fields("PtId").value
    
    Call txtPtId_LostFocus
    
    cboOrdDate.Clear
    cboOrdDate.AddItem Format("" & tmpRs.Fields("OrdDt").value, CS_DateMask)
    cboOrdDate.ListIndex = 0
    
    Call cboOrdDate_Click
    
NoData:
    Set tmpRs = Nothing
   
End Sub



'% 환자ID 또는 성명으로 검색 리스트 작성.
Private Sub txtSearchKey_GotFocus()

    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'% 환자ID 또는 성명으로 검색 리스트 작성.
Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Call LoadPatient
       
End Sub

Private Sub LoadPatient()
'검색조건에 해당되는 환자를 조회한다.
'optOption의 값 체크
    Dim objPtInfo1  As clsPatient
'    Dim objPtInfo2  As clsPtInformation
    Dim RS          As Recordset
    Dim itmX        As ListItem
    Dim strSQL      As String

    If Trim(txtSearchKey.Text) = "" Then
        
        Exit Sub
    End If
    If optOption(0).value Then
        If IsNumeric(txtSearchKey.Text) Then
            txtSearchKey.Text = Format(txtSearchKey.Text, P_PatientIdFormat)
        End If
    End If
    
    If optOption(0).value Then  '채혈대상에서 검색
        Set objPtInfo1 = New clsPatient  '  clsHosComSQLStmt
        
        If optSort(0).value Then
            strSQL = objPtInfo1.GetSQLPtNt("3", txtSearchKey.Text)
        Else
            strSQL = objPtInfo1.GetSQLPtNt("4", txtSearchKey.Text)
        End If
'        If optSort(0).Value Then    'ID로 검색
'            strSQL = objPtInfo1.SqlPtntSearch(IIf(optSort(0).Value, 1, 2) + 2, txtSearchKey.Text)
'        Else    '환자명으로 검색
'            strSQL = objPtInfo1.SqlPtntSearch(IIf(optSort(0).Value, 1, 2) + 2, Trim(txtSearchKey.Text))
'        End If
    Else
'        Set objPtInfo2 = New clsPtInformation
        Set objPtInfo1 = New clsPatient
        
        If optSort(0).value Then
            strSQL = objPtInfo1.GetSQLPtNt("1", txtSearchKey.Text)
        Else
            strSQL = objPtInfo1.GetSQLPtNt("2", txtSearchKey.Text)
        End If
'        If optSort(0).Value Then    'ID로 검색
'            strSQL = objPtInfo2.GetPtInfo(Trim(txtSearchKey.Text), True)
'        Else    '환자명으로 검색
'            strSQL = objPtInfo2.GetPtInfo(txtSearchKey.Text, False)
'        End If
    End If

On Error GoTo NoData
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    lvwPtList.ListItems.Clear
    
    If RS.EOF Then MsgBox "검색조건에 맞는 자료가 없습니다.", vbExclamation: GoTo NoData
            
    Do Until RS.EOF
        Set itmX = lvwPtList.ListItems.Add()
        
        itmX.Text = RS.Fields("ptid").value & ""
        itmX.SubItems(1) = RS.Fields("ptnm").value & ""
        itmX.SubItems(2) = RS.Fields("SSN").value & ""

' 08.09.25 양성현 생젼월일 대신 집주소로 변겅
'        itmX.SubItems(3) = Format(Rs.Fields("DOB").Value & "", CS_DateLongMask)
       
        itmX.SubItems(3) = Format(RS.Fields("address").value & "", CS_DateLongMask)
        Dim strYear As String
        
        If IsNull(RS.Fields("ssn").value & "") Then
            itmX.SubItems(4) = "XX"
        Else
'            If Len(Rs.Fields("ssn").Value & "") = 13 Then
'                itmX.SubItems(4) = IIf((Mid(Rs.Fields("ssn").Value & "", 7, 1) Mod 2) = 1, "남", "여")
'                If IsNumeric(Mid(Rs.Fields("ssn").Value & "", 1, 6)) Then
'                    itmX.SubItems(3) = Format(Mid(Rs.Fields("ssn").Value & "", 1, 6), "0#/##/##")
'                End If
'
'                If IsDate(itmX.SubItems(3)) Then
'                    itmX.SubItems(4) = itmX.SubItems(4) & " / " & DateDiff("yyyy", itmX.SubItems(3), GetSystemDate)
'                Else
'                    itmX.SubItems(4) = itmX.SubItems(4) & " / " & "0"
'                End If
'            End If
'
            strYear = ""
        End If
        
        If RS.Fields("hptelno").value & "" <> "" Then
            itmX.SubItems(5) = RS.Fields("hptelno").value & ""
        Else
            itmX.SubItems(5) = RS.Fields("telno").value & ""
        End If
        If lvwPtList.ListItems.Count = 100 Then Exit Do
        RS.MoveNext
    Loop
        
NoData:
    Set RS = Nothing
    Set objPtInfo1 = Nothing
'    Set objPtInfo2 = Nothing
End Sub

'% 정렬 기준 선택
Private Sub optSort_Click(Index As Integer)
    If txtSearchKey.Text <> "" Then
       Call txtSearchKey_KeyPress(vbKeyReturn)
    End If
    If blnInitFg Then txtSearchKey.SetFocus
End Sub

'% 환자ID가 변경되면 화면Clear
Private Sub txtPtId_Change()
    If Not blnClearFg Then Call ClearRtn
End Sub
'% 환자 ID
Private Sub txtPtId_GotFocus()
    With txtPtId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% 환자정보 검색
Private Sub txtPtId_KeyPress(KeyAscii As Integer)

    
    If KeyAscii = vbKeyReturn Then
        
        If txtPtId.Text <> "" Then
            cboOrdDate.SetFocus
        Else
            SendKeys "{TAB}"
        End If
    Else
        If txtReceptNo.Text <> "" Then txtReceptNo.Text = ""
    End If
End Sub

'% 검색한 처방을 테이블에 디스플레이 한다.
Private Sub DisplayOrder()
    Dim RS          As Recordset
    Dim i           As Integer
    Dim SqlStmt     As String
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim strDoctNm   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim strOrdDiv   As String
    Dim strPayDt     As String
   
On Error GoTo NoData
     ' 처방내역 검색
    tmpDate = Format(cboOrdDate.Text, CS_DateDbFormat)
    tmpTime = Format(Now, CS_TimeDbFormat)
    
    txtMesg.Text = ""
    DoEvents

    strOrdDiv = "W"
    strPayDt = IIf(chkPay.value = "1", "", "1")
    
    SqlStmt = objSql.SqlReadWardOrder(txtPtId.Text, tmpDate, tmpTime, , enBussDiv.BussDiv_OutPatient, strPayDt, strOrdDiv)
    
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If RS.EOF Then
        '-- 원본 =============================================
'        MsgBox objPatient.PtNm & " 님의 처방내역이 없습니다"
        '=====================================================
        
        '-- 예수병원 변경 루틴 ===============================
        If lblRIFlag.Caption <> "Y" Then
'            MsgBox objPatient.PtNm & " 님의 처방내역이 없습니다"
        End If
        '=====================================================
        
        MouseDefault
        blnOrdFg = False
        GoTo NoData
    End If
    
    With tblOrdSheet
       
        .ReDraw = False
        .MaxRows = 0
        If RS.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
            .Row = RS.RecordCount + 1
            .Row2 = lngMaxRows
            .Col = 1: .COL2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .Protect = True
            .BlockMode = False
        Else
            .MaxRows = RS.RecordCount   '데이타 건수
        End If

        'Locking Cells
        .Row = -1
        .Col = 2: .COL2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
             
        For i = 1 To RS.RecordCount
            .Row = i
'            strDoctNm = GetDoctName(RS.Fields("orddoct").Value & "")
            strDoctNm = GetDoctNm(RS.Fields("orddoct").value & "") 'Trim(Rs.Fields("DoctNm").Value & "")
            If SvOrdDt <> Trim("" & RS.Fields("OrdDt").value) Then
                .Col = enCOLLIST.tcORDDT:   .Text = Format("" & RS.Fields("OrdDt").value, CS_DateShortMask)    '처방일
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & RS.Fields("OrdNo").value)     '처방번호
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & RS.Fields("SpcNm").value)     '검체
                .Col = enCOLLIST.tcDOCTNM:  .Text = strDoctNm                               '처방의
                SvOrdDt = Trim("" & RS.Fields("OrdDt").value)
                SvOrdNo = Trim("" & RS.Fields("OrdNo").value)    '처방번호
                SvSpcNm = Trim("" & RS.Fields("SpcNm").value)    '검체
                SvOrdDoct = strDoctNm                           '처방의
            End If
            If SvOrdNo <> Trim("" & RS.Fields("OrdNo").value) Then
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & RS.Fields("OrdNo").value)     '처방번호
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & RS.Fields("SpcNm").value)     '검체
                .Col = enCOLLIST.tcDOCTNM:  .Text = strDoctNm                               '처방의
                SvOrdNo = Trim("" & RS.Fields("OrdNo").value)    '처방번호
                SvSpcNm = Trim("" & RS.Fields("SpcNm").value)    '검체
                SvOrdDoct = strDoctNm                            '처방의
            End If
            
            If SvSpcNm <> Trim("" & RS.Fields("SpcNm").value) Then
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & RS.Fields("SpcNm").value)     '검체
                SvSpcNm = Trim("" & RS.Fields("SpcNm").value)

            End If
            If SvOrdDoct <> strDoctNm Then
                .Col = enCOLLIST.tcDOCTNM: .Text = strDoctNm  '처방의
                SvOrdDoct = Trim(.Text)
            End If

'            Dim objBld As clsBasisData
            Dim strBld As String
            
            Select Case RS.Fields("orddiv").value & ""

                Case BBS_ORDDIV:
                    .Col = enCOLLIST.tcSTATFG: .value = Trim("" & RS.Fields("StatFg").value)     '응급여부  --> 위에서 처리...
                    .Col = enCOLLIST.tcBUILDCD: .value = ObjSysInfo.BuildingCd
                    
'                    Set objBld = Nothing
'                    Set objBld = New clsBasisData
                    strBld = GetBuildNm(.value)
'                    Set objBld = Nothing
                    
'                    If ObjLISComCode.Building.Exists(.Value) Then
'                        ObjLISComCode.Building.KeyChange (.Value)
'                    End If
                    .Col = enCOLLIST.tcBUILDNM: .value = strBld 'ObjLISComCode.Building.Fields("buildnm")
                
                Case LIS_ORDDIV:
                    .Col = enCOLLIST.tcBUILDCD:  .Text = ObjSysInfo.BuildingCd
                    .Col = enCOLLIST.tcBUILDNM:  .Text = ObjSysInfo.BuildingNm
                    .Col = enCOLLIST.tcSTATFLAG: .Text = Trim(RS.Fields("StatFg").value & "")
            End Select
          
DataSet:
            .Col = enCOLLIST.tcTESTNM:  .Text = Trim("" & RS.Fields("TestNm").value)     '처방명
            Select Case RS.Fields("orddiv").value & ""
                Case BBS_ORDDIV: .ForeColor = &H496835     '&H6C6181     '&H81815A     '약간녹색   &H00845584&보라색
                Case LIS_ORDDIV: .ForeColor = &H553755
            End Select
            .Col = enCOLLIST.tcSTATFG:
                .Text = IIf("" & RS.Fields("StatFg").value = "1", "Y", "") '응급여부
                .ForeColor = DCM_Red                                '빨간색
            .Col = enCOLLIST.tcREQDTTM: .Text = Format("" & RS.Fields("ReqDt").value, CS_DateMask) & " " & _
                                         Format(IIf(Len("" & RS.Fields("ReqTm").value) = 4, "" & RS.Fields("ReqTm").value & "00", "" & RS.Fields("ReqTm").value), CS_TimeLongMask) '희망채취일시
            
            '처방일과 희망채혈일시가 다른경우 표시
            If Trim("" & RS.Fields("OrdDt").value <> Trim("" & RS.Fields("ReqDt").value)) Then
                .ForeColor = DCM_LightRed
            End If
            
            .Col = enCOLLIST.tcORDDATE: .Text = Trim("" & RS.Fields("OrdDt").value)      '처방일
            .Col = enCOLLIST.tcORDNUM:  .Text = Trim("" & RS.Fields("OrdNo").value)      '처방번호
            .Col = enCOLLIST.tcORDSEQ:  .Text = Trim("" & RS.Fields("OrdSeq").value)     '처방Seq
            .Col = enCOLLIST.tcTESTCD:  .Text = Trim("" & RS.Fields("OrdCd").value)      '검사코드
            
            Dim strLabDiv As String
            strLabDiv = GetLabDiv(.Text)
            
'            Call ObjLISComCode.LisItem.KeyChange(.Text)
            .Col = enCOLLIST.tcLABDIV:  .Text = strLabDiv 'ObjLISComCode.LisItem.Fields("labdiv")      'LabDiv

            .Col = enCOLLIST.tcSPCCD:   .Text = Trim("" & RS.Fields("SpcCd").value)      '검체코드
            
            Dim strLabRng As String
            Dim strSpcAbbr As String
            
            Call GetSpcInfo(.Text, strSpcAbbr, strLabRng)
            
'            Call ObjLISComCode.LisSpc.KeyChange(.Text)
            .Col = enCOLLIST.tcSPCABBR:  .Text = Trim("" & RS.Fields("spcnm5").value)         '검체약어명
            .Col = enCOLLIST.tcLABRANGE: .Text = strLabRng 'ObjLISComCode.LisSpc.Fields("labrange")    '미생물접수번호범위

            .Col = enCOLLIST.tcWORKAREA: .Text = Trim("" & RS.Fields("WorkArea").value)  'WorkArea
            
                
            .Col = enCOLLIST.tcSTORECD:  .Text = Trim("" & RS.Fields("StoreCd").value)   '보관코드
            .Col = enCOLLIST.tcTESTDIV:  .Text = Trim("" & RS.Fields("TestDiv").value)   '검사구분
            .Col = enCOLLIST.tcMULTIFG:  .Text = Trim("" & RS.Fields("MultiFg").value)   '복수검체여부
            .Col = enCOLLIST.tcSPCGRP:   .Text = Trim("" & RS.Fields("SpcGrp").value)    '검체군
            .Col = enCOLLIST.tcORDDOCT:  .Text = Trim("" & RS.Fields("OrdDoct").value)   '처방의
            .Col = enCOLLIST.tcMAJDODT:  .Text = Trim("" & RS.Fields("MajDoct").value)   '주치의
            .Col = enCOLLIST.tcDEPTCD:   .Text = Trim("" & RS.Fields("DeptCd").value)    '진료과
                                         '진료과명
                                         'If .Text <> "" And lblDeptNm.Caption = "" Then
                                         If .Text <> "" Then
                                            strDeptCd = .Text
                                            
'                                            Set objBld = Nothing
'                                            Set objBld = New clsBasisData
                                            strBld = GetDeptNm(.Text)
'                                            Set objBld = Nothing
                                            
                                            lblDeptNm.Caption = strBld
'                                            If ObjLISComCode.DeptCd.Exists(.Text) Then
'                                                ObjLISComCode.DeptCd.KeyChange (.Text)
'                                                lblDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'                                            End If
                                         End If
            .Col = enCOLLIST.tcABBRNM:  .Text = Trim("" & RS.Fields("AbbrNm5").value)    '약어명
            .Col = enCOLLIST.tcBARCNT:  .Text = Trim("" & RS.Fields("LabelCnt").value)   '라벨출력장수
            
            .Col = enCOLLIST.tcPAYDT:   .Text = Trim("" & RS.Fields("PayDt").value)   '영수증번호
                                        .ForeColor = vbRed

            .Col = enCOLLIST.tcWARDID:  .Text = Trim("" & RS.Fields("WardId").value)     '병동
            .Col = enCOLLIST.tcROOMID:  .Text = Trim("" & RS.Fields("hosilid").value)     '병실
            .Col = enCOLLIST.tcBEDID:   .Text = Trim("" & RS.Fields("roomid").value)      '병상

            .Col = enCOLLIST.tcFRZFG:   .Text = Trim("" & RS.Fields("FzFg").value)       '동결절편
            .Col = enCOLLIST.tcORDDIV:  .Text = Trim("" & RS.Fields("OrdDiv").value)     '처방구분
            
            '진료부서 Remark
            If Trim("" & RS.Fields("Mesg").value) <> "" Then
                txtMesg.Text = txtMesg.Text & "# " & Format(Trim("" & RS.Fields("OrdNo").value), "##") & " - "
                txtMesg.Text = txtMesg.Text & Trim("" & RS.Fields("TestNm").value) & vbCrLf
                txtMesg.Text = txtMesg.Text & Trim("" & RS.Fields("Mesg").value) & vbCrLf
            End If

            RS.MoveNext
        Next

        .RowHeight(-1) = lngRowHeight
        .ReDraw = True
       
    End With
    blnOrdFg = True
    fraOrder.Enabled = True
    
NoData:
    Call MouseDefault
    Set RS = Nothing
End Sub

Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b"
    strSQL = strSQL & " where " & DBW("b.cdindex=", LC3_WorkArea)
    strSQL = strSQL & " and a.workarea=b.cdval1"
    strSQL = strSQL & " and " & DBW("a.testcd=", vTestCd)
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    If RS.EOF = False Then
        GetLabDiv = RS.Fields("field2").value & ""
    End If
    Set RS = Nothing
End Function

Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbr As String, _
                            ByRef vLabRng As String)
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = " select  a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm,  " & _
            " a.field1 multifg, a.field2 spcgrp, b.field2 labrange " & _
            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
            " where " & DBW("a.cdindex =", LC3_Specimen) & _
            " and " & DBW("a.cdval1=", vSpcCd) & _
            " and    " & DBJ("b.cdindex ='C217'") & _
            " and    " & DBJ("b.cdval1  =* a.field2")

    Set RS = New Recordset
    RS.Open strSQL, DBConn
    If RS.EOF = False Then
        vSpcAbbr = RS.Fields("spcabbr").value & ""
        vLabRng = RS.Fields("labrange").value & ""
    End If
    Set RS = Nothing
End Sub

Private Sub ClearRtn()
    lblDiseaseSang.Visible = False
    lblSang(14).Visible = False
    lblDisease.Caption = ""
    lblDiseaseSang.Caption = ""
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDeptNm.Caption = ""
    lblOrdDtCnt.Caption = ""
    lblColID.Caption = ""
    txtMesg.Text = ""
    chkSelAll.value = 0
    cboOrdDate.Clear
    fraOrder.Enabled = False
    With tblOrdSheet
        .Row = -1
        .Col = -1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    cmdSave(0).Enabled = False
    cmdSave(1).Enabled = False
    blnOrdFg = False
    blnMsgFg = False
    blnClearFg = True
    cmdContent.Visible = False
    fraContent.Visible = False
    fraPtRmk.Visible = False
    cmdRRmk.Visible = False
    lblRIFlag.Caption = ""
    lblRIMsg.Caption = ""
    
    strWrkDiv = ""
    
    Call PtRmkClear
    Set objPatient = Nothing
    Set objPatient = New clsPatient
'    Set objPatient.objDB = dbconn
    Set objCollect = Nothing
    Set objCollect = New clsLISCollectioin
    
End Sub
Public Sub Call_cmdClear_Click()
    Call cmdClear_Click
End Sub

Private Function CollectionTargetChk() As Boolean
    Dim ii      As Integer
    Dim tmpDate As String
    Dim strTmp  As String
    Dim strMsg  As String
    
    tmpDate = Format(GetSystemDate, CS_DateDbFormat) & " " & Format(GetSystemDate, CS_TimeDbFormat)
    tmpDate = Replace(Replace(Replace(tmpDate, "-", ""), ":", ""), " ", "")
    
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = enCOLLIST.tcCHECK
            If .value = 1 Then
                CollectionTargetChk = True
                Exit For
            End If
        Next
        
        If CollectionTargetChk = False Then
            MsgBox "채혈할 항목을 선택하세요..", vbInformation, "항목선택"
            Exit Function
        End If
        
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = enCOLLIST.tcORDDIV
            If .value = LIS_ORDDIV Then
                .Col = enCOLLIST.tcREQDTTM: strTmp = .value
                strTmp = Replace(Replace(Replace(strTmp, "-", ""), ":", ""), " ", "")
                If tmpDate < strTmp Then
                    strMsg = MsgBox("희망채혈시간과 관련되어 채혈할수 없습니다." & vbCrLf & _
                            "채혈처리 하시겠습니까?", vbYesNo + vbInformation, "Info")
                    If strMsg = vbNo Then CollectionTargetChk = False
                    
                    Exit For
                End If
            End If
        Next
        
    End With
    
End Function

Private Sub tblordersheet()
    With tblOrdSheet
        .SortBy = SortByRow
        .SortKey(1) = enCOLLIST.tcORDDIV
        .SortKeyOrder(1) = SortKeyOrderAscending
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Action = ActionSort
    End With
End Sub

Public Function AccListDisplayer()
    If mvarPtID = "" Then Exit Function
    blnClearFg = False
    txtPtId.SetFocus
    txtPtId.Text = mvarPtID
    cboOrdDate.AddItem Format(mvarOrddt, CS_DateLongFormat)
    
    cboOrdDate.ListIndex = medComboFind(cboOrdDate, Format(mvarOrddt, CS_DateLongFormat))

End Function

'환자별 특이사항등록 조회에 관련된 부분입니다.
'Table: s2ptmesg
'20040203

Private Sub cmdRSave_Click()
    Call InsertPtRmk
    Call QueryPtRmk
End Sub
Private Sub cmdRClose_Click()
    fraPtRmk.Visible = False
End Sub

Private Sub cmdRRmk_Click()
    Call RmkScreenSetting
    fraPtRmk.Visible = True: cmdRSave.Visible = False
    fraPtRmk.ZOrder
    Call QueryPtRmk
End Sub

Private Sub cmdRRmkVisible_Click()
    If txtPtId.Text = "" Then Exit Sub
    Call QueryPtRmk
    
    fraPtRmk.Visible = True:  cmdRSave.Visible = True
    fraPtRmk.ZOrder
    Call RmkScreenSetting(False)
    Call PtRmkClear
End Sub

Private Sub RmkScreenSetting(Optional ByVal blnSave As Boolean = True)
    LisLabel4(9).Visible = True:  lblRPtid.Visible = True: lblRptNm.Visible = True
    LisLabel4(0).Visible = True:  dtpEntDt.Visible = True: lblREntNm.Visible = True
    
    LisLabel4(1).Visible = True:  txtRDept.Visible = True: cmdHelpList(0).Visible = True: lblRDeptNm.Visible = True
    LisLabel4(2).Visible = True:  txtRColid.Visible = True: cmdHelpList(1).Visible = True: lblRColNm.Visible = True
    
    LisLabel4(3).Visible = True:  Frame2.Visible = True:
    
    LisLabel4(6).Visible = True:  cboRmk.Visible = True
    LisLabel4(4).Visible = True:  txtRTitle.Visible = True
    LisLabel4(10).Visible = True: txtRMesg.Visible = True

    LisLabel4(1).Top = 1185:  txtRDept.Top = 1185:  cmdHelpList(0).Top = 1185: lblRDeptNm.Top = 1185
    LisLabel4(1).Enabled = True:  txtRDept.Enabled = True:  cmdHelpList(0).Enabled = True: lblRDeptNm.Enabled = True
    
    LisLabel4(2).Top = 1515:  txtRColid.Top = 1515: cmdHelpList(1).Top = 1515: lblRColNm.Top = 1515
    LisLabel4(2).Enabled = True:  txtRColid.Enabled = True: cmdHelpList(1).Enabled = True: lblRColNm.Enabled = True
    
    LisLabel4(6).Top = 2175:  cboRmk.Top = 2175
    LisLabel4(4).Top = 2520:  txtRTitle.Top = 2520
    LisLabel4(10).Top = 2855: txtRMesg.Top = 2855
    LisLabel4(6).Enabled = True:  cboRmk.Enabled = True
    LisLabel4(4).Enabled = True:  txtRTitle.Enabled = True
    LisLabel4(10).Enabled = True: txtRMesg.Enabled = True
    LisLabel4(10).Height = 1800: txtRMesg.Height = 1800
    If blnSave = True Then
        LisLabel4(9).Visible = False:  lblRPtid.Visible = False: lblRptNm.Visible = False
        LisLabel4(0).Visible = False:  dtpEntDt.Visible = False: lblREntNm.Visible = False
        LisLabel4(3).Visible = False:  Frame2.Visible = False:
    
    
        LisLabel4(1).Enabled = False:  txtRDept.Enabled = False:  cmdHelpList(0).Enabled = False: lblRDeptNm.Enabled = False
        LisLabel4(2).Enabled = False:  txtRColid.Enabled = False: cmdHelpList(1).Enabled = False: lblRColNm.Enabled = False
        LisLabel4(6).Enabled = False:  cboRmk.Enabled = False
        LisLabel4(4).Enabled = False:  txtRTitle.Enabled = False
        LisLabel4(10).Enabled = False: txtRMesg.Enabled = False
        LisLabel4(1).Top = 525:  txtRDept.Top = 525:  cmdHelpList(0).Top = 525: lblRDeptNm.Top = 525
        LisLabel4(2).Top = 855:  txtRColid.Top = 855: cmdHelpList(1).Top = 855: lblRColNm.Top = 855
        
        LisLabel4(6).Top = 1185:  cboRmk.Top = 1185
        LisLabel4(4).Top = 1515:  txtRTitle.Top = 1515
        LisLabel4(10).Top = 1845: txtRMesg.Top = 1845
        LisLabel4(10).Height = 2805: txtRMesg.Height = 2805
    End If

End Sub


Private Sub PtRmkClear()
    lblRPtid.Caption = "":  lblRptNm.Caption = "":  lblRDeptNm.Caption = "":    lblREntNm.Caption = ""
    lblRColNm.Caption = "": lblRSeq.Caption = ""
    txtRMesg.Text = "":     txtRTitle.Text = "":    txtRColid.Text = "":        txtRDept.Text = ""
    optExp(1).value = True
    dtpEntDt.value = GetSystemDate
    lblRPtid.Caption = txtPtId.Text: lblRptNm.Caption = lblPtNm.Caption
    lblREntNm.Caption = ObjSysInfo.EmpNm
    cmdRSave.Caption = "저장"
    If cboRmk.ListCount > 0 Then cboRmk.ListIndex = 0
    If fraPtRmk.Visible = True And txtRDept.Enabled Then txtRDept.SetFocus

End Sub
Private Sub dtpEntDt_Click()
    txtRDept.SetFocus
End Sub

Private Sub txtRColid_KeyPress(KeyAscii As Integer)
'    Dim objEmp As clsBasisData
    
    If KeyAscii = vbKeyReturn Then
'        Set objEmp = New clsBasisData
        
        lblRColNm.Caption = GetEmpNm(UCase(txtRColid.Text)) 'GetEmpname(UCase(txtRColid.Text))
        If lblRColNm.Caption <> "" Then
            txtRTitle.SetFocus
        Else
            lblRColNm.Caption = ""
            txtRColid.Text = "": txtRColid.SetFocus
        End If
        
'        Set objEmp = Nothing
    End If
End Sub

Private Sub txtRDept_KeyPress(KeyAscii As Integer)
'    Dim objDept As clsBasisData
    Dim strDept As String
    
    If KeyAscii = vbKeyReturn Then
'        Set objDept = New clsBasisData
        strDept = GetDeptNm(txtRDept.Text)
'        Set objDept = Nothing
        
        If strDept <> "" Then
            lblRDeptNm.Caption = strDept
            txtRColid.SetFocus
        Else
            lblRDeptNm.Caption = "": txtRDept.Text = "": txtRDept.SetFocus
        End If
'        If ObjLISComCode.DeptCd.Exists(UCase(txtRDept.Text)) Then
'            ObjLISComCode.DeptCd.KeyChange UCase(txtRDept.Text)
'            lblRDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'            txtRColid.SetFocus
'        Else
'            lblRDeptNm.Caption = "": txtRDept.Text = "": txtRDept.SetFocus
'        End If
    End If
End Sub

Private Sub txtRTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtRMesg.SetFocus
End Sub

Private Sub cmdHelpList_Click(Index As Integer)
'    Dim objEmp As clsBasisData
    
    Set objMyList = New clsPopUpList
'    Set objEmp = New clsBasisData
    
    With objMyList
        .Connection = DBConn
        .HideToolTipText = True
        
        Select Case Index
            Case 1
                 .Connection = DBConn
                 .FormCaption = "사용자조회"
                 .tag = "ColID"
                 .ColumnHeaderText = "사번;성명"
                 Call .LoadPopUp(GetSQLEmpList) ', fraPtRmk.Top + cmdHelpList(Index).Top, fraPtRmk.Left + cmdHelpList(Index).Left)
            Case 0
                 .FormCaption = "진료과 조회"
                 .tag = "DeptCd"
                 .ColumnHeaderText = "진료과코드;진료과명"
                 Call .LoadPopUp(GetSQLDeptList) ', fraPtRmk.Top + cmdHelpList(Index).Top, fraPtRmk.Left + cmdHelpList(Index).Left)
        End Select
    End With
'    Set objEmp = Nothing
    Set objMyList = Nothing
End Sub
Private Sub objMyList_SendCode(ByVal SelString As String)
    Select Case objMyList.tag
        Case "DeptCd"
            txtRDept.Text = Trim(medGetP(SelString, 1, ";"))
            lblRDeptNm.Caption = Trim(medGetP(SelString, 2, ";"))
        Case "ColID"
            txtRColid.Text = Trim(medGetP(SelString, 1, ";"))
            lblRColNm.Caption = Trim(medGetP(SelString, 2, ";"))
    End Select
End Sub
Private Sub tblRData_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ii  As Integer
    
    cmdRSave.Caption = "저장"
    If Row < 1 Then Exit Sub
    Call PtRmkClear
    With tblRData
        .Row = Row
        .Col = 1: If .value = "" Then Exit Sub
        .Col = 1: dtpEntDt.value = CDate(.value)
        .Col = 2: txtRTitle.Text = .value
        .Col = 3: lblREntNm.Caption = medGetP(.value, 1, COL_DIV)
        .Col = 4: txtRColid.Text = medGetP(.value, 1, COL_DIV): lblRColNm.Caption = medGetP(.value, 2, COL_DIV)
        .Col = 5: txtRDept.Text = medGetP(.value, 1, COL_DIV):  lblRDeptNm.Caption = medGetP(.value, 2, COL_DIV)
        .Col = 6: optExp(0).value = IIf(.value = "1", False, True)
        .Col = 7: txtRMesg.Text = .value
        .Col = 8: lblRSeq.Caption = .value
        .Col = 9:
            For ii = 0 To cboRmk.ListCount
                If .value = Trim(medGetP(cboRmk.List(ii), 2, vbTab)) Then
                    cboRmk.ListIndex = ii
                    Exit For
                End If
            Next
    End With
    cmdRSave.Caption = "수정"
    
End Sub
Private Sub QueryPtRmk()
    Dim RS   As Recordset
    Dim SSQL As String
    Dim ii   As Integer
    
    Call GetcboRmkCd
    Call medClearTable(tblRData)
    On Error GoTo Error_Jump
    Set RS = New Recordset
    
    SSQL = GetPtRmkSQL
    
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tblRData
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .value = Format(RS.Fields("entdt").value & "", "####-##-##")
                .Col = 2: .value = RS.Fields("title").value & ""
                .Col = 3: .value = RS.Fields("empid").value & ""
                .Col = 4: .value = RS.Fields("colid").value & ""
                .Col = 5: .value = RS.Fields("deptcd").value & ""
                .Col = 6: .value = RS.Fields("expfg").value & ""
                .Col = 7: .value = RS.Fields("mesg").value & ""
                .Col = 8: .value = RS.Fields("seq").value & ""
                .Col = 9: .value = RS.Fields("rmkcd").value & ""
                RS.MoveNext
            Loop
        End With
    
    
    End If
    Call tblRData_Click(1, 1)
    
Error_Jump:
    Set RS = Nothing
End Sub
Private Sub InsertPtRmk()
    Dim SSQL        As String
    Dim strMaxSEQ   As String
    Dim strExpFg    As String
    Dim strExpDt    As String
    Dim strRmkCD    As String
    
    
    strMaxSEQ = GetPtRmkMaxSeq
    If cboRmk.ListCount = 0 Then
        strRmkCD = ""
    Else
        strRmkCD = Trim(medGetP(cboRmk.Text, 2, vbTab))
    End If
    If strMaxSEQ = "" Then Exit Sub
    strExpFg = "0"
    
    If optExp(1).value Then
        strExpFg = "1"
        strExpDt = Format(GetSystemDate, "YYYYMMDD")
    End If
    
    On Error GoTo Error_Jump
    DBConn.BeginTrans
    
    If lblRSeq.Caption = "" Then
        SSQL = " insert into " & T_LAB902 & " ( ptid,seq,entdt,empid,colid,deptcd,expfg,expdt,rmkcd,title,mesg) values(" & _
                DBV("ptid", lblRPtid.Caption, 1) & DBV("seq", strMaxSEQ, 1) & DBV("entdt", Format(dtpEntDt.value, "YYYYMMDD"), 1) & _
                DBV("empid", lblREntNm.Caption, 1) & DBV("colid", txtRColid.Text & COL_DIV & lblRColNm.Caption, 1) & DBV("deptcd", txtRDept.Text & COL_DIV & lblRDeptNm.Caption, 1) & _
                DBV("expfg", strExpFg, 1) & DBV("expdt", strExpDt, 1) & DBV("rmkcd", strRmkCD, 1) & DBV("title", txtRTitle.Text, 1) & DBV("mesg", txtRMesg) & ")"
    Else
        SSQL = " update  " & T_LAB902 & " set " & _
                DBW("entdt", Format(dtpEntDt.value, "YYYYMMDD"), 3) & DBW("empid", ObjSysInfo.EmpId & COL_DIV & lblREntNm.Caption, 3) & _
                DBW("colid", txtRColid.Text & COL_DIV & lblRColNm.Caption, 3) & DBW("deptcd", txtRDept.Text & COL_DIV & lblRDeptNm.Caption, 3) & _
                DBW("expfg", strExpFg, 3) & DBW("expdt", strExpDt, 3) & DBW("rmkcd", strRmkCD, 3) & DBW("title", txtRTitle.Text, 3) & DBW("mesg", txtRMesg, 2) & _
               " WHERE " & _
                DBW("ptid=", lblRPtid.Caption) & " AND " & DBW("seq=", lblRSeq.Caption)
    End If
    
    DBConn.Execute SSQL
    
    DBConn.CommitTrans
    Exit Sub
    
Error_Jump:
    DBConn.RollbackTrans
    MsgBox Err.Description
End Sub
Private Function GetPtRmkMaxSeq() As String
    Dim RS      As Recordset
    Dim SSQL    As String
    
    On Error GoTo Error_Jump
    
    SSQL = " SELECT max(seq) as maxseq FROM  " & T_LAB902 & "" & _
           " WHERE " & DBW("ptid", lblRPtid.Caption, 2)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        GetPtRmkMaxSeq = Val(RS.Fields("maxseq").value & "") + 1
    Else
        GetPtRmkMaxSeq = 1
    End If
    
Error_Jump:
    Set RS = Nothing
    
End Function
Private Sub GetPtRmkVisibleTrueFalse()
    Dim SSQL As String
    Dim RS   As Recordset
'    Dim objEmp As clsBasisData
    
    lblColID.Caption = ""
    
    cmdRRmk.Visible = False
    On Error GoTo Error_Jump
    SSQL = GetPtRmkSQL
    
'    Set objEmp = New clsBasisData
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        lblColID.Caption = GetEmpNm(medGetP(RS.Fields("colid").value & "", 1, COL_DIV)) 'GetEmpname(medGetP(Rs.Fields("colid").Value & "", 1, COL_DIV))

        cmdRRmk.Visible = True
    End If
Error_Jump:
    Set RS = Nothing
'    Set objEmp = Nothing

End Sub
Private Function GetPtRmkSQL() As String
    Dim SSQL As String
    
    SSQL = " SELECT * FROM  " & T_LAB902 & " " & _
           " WHERE " & DBW("ptid=", txtPtId.Text)
           
    GetPtRmkSQL = SSQL & " ORDER BY entdt desc"
End Function

Private Sub GetcboRmkCd()
    Dim SSQL    As String
    Dim RS      As Recordset
    
    SSQL = " SELECT cdval1,text1 FROM " & T_LAB034 & " WHERE " & DBW("cdindex=", LC4_AccessComment)
    
    cboRmk.Clear
    cboRmk.AddItem "[ Remark Templete ]"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            cboRmk.AddItem RS.Fields("text1").value & "" & vbTab & RS.Fields("cdval1").value & ""
            RS.MoveNext
        Loop
        cboRmk.ListIndex = 1
    End If
    Set RS = Nothing
End Sub
