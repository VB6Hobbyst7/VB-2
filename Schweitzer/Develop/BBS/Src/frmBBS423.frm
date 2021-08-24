VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS423 
   BackColor       =   &H00DBE6E6&
   Caption         =   "헌혈 혈액등록"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "등록취소"
      Height          =   495
      Left            =   8145
      Style           =   1  '그래픽
      TabIndex        =   53
      Tag             =   "15101"
      Top             =   7995
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   10695
      Style           =   1  '그래픽
      TabIndex        =   55
      Tag             =   "128"
      Top             =   7995
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   9420
      Style           =   1  '그래픽
      TabIndex        =   54
      Tag             =   "124"
      Top             =   7995
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   495
      Left            =   6870
      Style           =   1  '그래픽
      TabIndex        =   52
      Tag             =   "15101"
      Top             =   7995
      Width           =   1215
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2295
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4260
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  헌 혈 내 역"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   315
      Left            =   7275
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3030
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   " 지정 환자"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   0
      Left            =   2295
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3030
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  헌혈 종류"
      Appearance      =   0
   End
   Begin VB.Frame fraAcc1 
      BackColor       =   &H00DBE6E6&
      Enabled         =   0   'False
      Height          =   870
      Left            =   2295
      TabIndex        =   23
      Top             =   3255
      Width           =   4950
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Pheresis"
         Height          =   435
         Index           =   1
         Left            =   2625
         Style           =   1  '그래픽
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   1245
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "지정 헌혈"
         Height          =   435
         Index           =   0
         Left            =   960
         Style           =   1  '그래픽
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1245
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   2295
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1980
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  헌혈자 등록 일자"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   690
      Left            =   2295
      TabIndex        =   18
      Top             =   2205
      Width           =   9945
      Begin VB.ComboBox cboDonoraccdt 
         Height          =   300
         Left            =   1950
         Sorted          =   -1  'True
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   3255
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   6
         Left            =   510
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   225
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "등록일자"
         Appearance      =   0
      End
      Begin VB.Label lblAccCnt 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         ForeColor       =   &H000040C0&
         Height          =   180
         Left            =   5235
         TabIndex        =   21
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.Frame fraBld 
      BackColor       =   &H00DBE6E6&
      Height          =   3300
      Left            =   2295
      TabIndex        =   31
      Top             =   4485
      Width           =   9930
      Begin VB.Frame fraVol 
         BackColor       =   &H00DBE6E6&
         Height          =   615
         Left            =   1560
         TabIndex        =   38
         Top             =   1095
         Width           =   4830
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "320cc"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   240
            Width           =   795
         End
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "기타"
            Height          =   225
            Index           =   3
            Left            =   2715
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtVolume 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            Height          =   285
            Left            =   3570
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   210
            Width           =   870
         End
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "400cc"
            Height          =   225
            Index           =   1
            Left            =   1005
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optVol 
            BackColor       =   &H00DBE6E6&
            Caption         =   "250cc"
            Height          =   225
            Index           =   2
            Left            =   1860
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "cc"
            Height          =   180
            Left            =   4485
            TabIndex        =   44
            Top             =   315
            Width           =   210
         End
      End
      Begin VB.TextBox txtBldNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   35
         Top             =   765
         Width           =   2160
      End
      Begin VB.CheckBox chkBar 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드로 입력(&B)"
         Height          =   315
         Left            =   3945
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   765
         Value           =   1  '확인
         Width           =   1755
      End
      Begin VB.ComboBox cboCompo 
         Height          =   300
         ItemData        =   "frmBBS423.frx":0000
         Left            =   1560
         List            =   "frmBBS423.frx":000D
         Style           =   2  '드롭다운 목록
         TabIndex        =   33
         Top             =   330
         Width           =   4095
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   16
         Left            =   360
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "혈액량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   360
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   765
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "혈액번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "헌혈일자"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpColdt 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   46
         Top             =   1830
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   60227587
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2295
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "폐기일자"
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
         Height          =   345
         Left            =   1560
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2295
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   60227587
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   360
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2730
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "보관일수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   14
         Left            =   360
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   330
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "혈액제제"
         Appearance      =   0
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "일"
         Height          =   180
         Index           =   0
         Left            =   2565
         TabIndex        =   51
         Top             =   2805
         Width           =   180
      End
      Begin VB.Label lblAvailable 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Left            =   1560
         TabIndex        =   50
         Top             =   2745
         Width           =   915
      End
   End
   Begin VB.Frame fraAcc2 
      BackColor       =   &H00DBE6E6&
      Height          =   870
      Left            =   7275
      TabIndex        =   27
      Top             =   3255
      Width           =   4950
      Begin VB.TextBox txtReservedID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   945
         MaxLength       =   10
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "10293023"
         Top             =   315
         Width           =   1305
      End
      Begin MedControls1.LisLabel lblReservedNm 
         Height          =   315
         Left            =   2280
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   315
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         BackColor       =   12632256
         ForeColor       =   0
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
         Caption         =   "홍길동"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2295
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   645
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  헌혈자 기본정보"
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2295
      TabIndex        =   1
      Top             =   870
      Width           =   9945
      Begin VB.TextBox txtDonorNm 
         Appearance      =   0  '평면
         Height          =   330
         Left            =   1050
         TabIndex        =   3
         Top             =   165
         Width           =   1515
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   5655
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   5655
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   525
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "총 헌혈량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDOB 
         Height          =   330
         Left            =   4290
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
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
         Caption         =   "2001-01-01"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSex 
         Height          =   330
         Left            =   6645
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   165
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   "M/100"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   330
         Left            =   8955
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   165
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCnt 
         Height          =   330
         Left            =   4290
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   525
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
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
         Alignment       =   2
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblTotVol 
         Height          =   330
         Left            =   6645
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   525
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
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
         Alignment       =   2
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDonorID 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSSN 
         Height          =   315
         Left            =   8415
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "성   명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   3300
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "생년월일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   3300
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   525
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "헌혈횟수"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   7965
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "cc"
         Height          =   180
         Left            =   7605
         TabIndex        =   15
         Top             =   660
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmBBS423"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CANEDIT_STATUS& = 5   'STATUS 가 5인 경우까지 수정가능(즉, 헌혈등록이나 페리시스 등록이 안된 경우)

Private CurrentStatus As Long

Private Function GetBldNo() As String
    '입력된 혈액번호를 ##-##-#양식으로 반환한다.
    If chkBar.value = 1 Then
        GetBldNo = Mid(txtBldNo.Text, 1, 2) & "-" & Mid(txtBldNo.Text, 3, 2) & "-" & Mid(txtBldNo.Text, 5, 6)
    Else
        GetBldNo = txtBldNo.Text
    End If
End Function

Private Function CheckExist(ByVal vBldNo As String) As Boolean
    Dim RS As Recordset
    Dim strSQL As String
    Dim strSrc As String
    Dim strYY As String
    Dim strNo As String
    Dim strCompocd As String
    
    strSrc = Mid(vBldNo, 1, 2)
    strYY = Mid(vBldNo, 3, 2)
    strNo = Mid(vBldNo, 5, 6)
    strCompocd = medGetP(cboCompo.Text, 1, COL_DIV)
    
    strSQL = " SELECT bldsrc, bldyy, bldno, compocd FROM " & T_BBS401
    strSQL = strSQL & " WHERE " & DBW("bldsrc=", strSrc)
    strSQL = strSQL & " AND " & DBW("bldyy=", strYY)
    strSQL = strSQL & " AND " & DBW("bldno=", strNo)
    strSQL = strSQL & " AND " & DBW("compocd=", strCompocd)
    
    Set RS = New Recordset
    
    RS.Open strSQL, DBConn
    
    If RS.EOF Then
        CheckExist = False
    Else
        CheckExist = True
    End If
    
    Set RS = Nothing
End Function

Private Sub cboCompo_Click()
    lblAvailable.Caption = medGetP(Get_CompNm(medGetP(cboCompo.Text, 1, COL_DIV)), 2, COL_DIV)
    dtpExpDt.value = DateAdd("d", Val(lblAvailable.Caption) - 1, dtpColDt.value)

    On Error Resume Next
    txtBldNo.SetFocus
End Sub

Private Sub cboDonoraccdt_Click()
    Call InitInfo
    Call LoadAccInfo
    Call LoadCompo
    Call LoadBldInfo

    On Error Resume Next
    txtReservedID.SetFocus
End Sub

Private Sub LoadAccInfo()
'등록되어 있는 기본정보 조회
    Dim RS As Recordset
    Dim strSQL As String
    
    'donorcd 0:지정헌혈,3:페리시스
    
    strSQL = " SELECT a.donorcd,a.reservedid, a.weight, a.height, a.pulse, a.bodytemp,a.bldpres1, a.bldpres2, "
    strSQL = strSQL & " b.stscd, b.rmk1, b.rmk2, b.rmk3, a.bldsrc, a.bldyy, a.bldno, a.compocd "
    strSQL = strSQL & " FROM " & T_BBS602 & " a, " & T_BBS603 & " b "
    strSQL = strSQL & " WHERE a.donorid=b.donorid"
    strSQL = strSQL & " AND " & DBW("a.donorid=", lblDonorID.Caption)
    strSQL = strSQL & " AND " & DBW("a.donoraccdt=", Format(cboDonoraccdt.Text, PRESENTDATE_FORMAT))
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    With RS
        If .EOF = False Then
            If .Fields("donorcd").value & "" = "1" Then
                optDonorCd(0).value = True
            ElseIf .Fields("donorcd").value & "" = "3" Then
                optDonorCd(1).value = True
            End If
            
            txtReservedID.Text = .Fields("reservedid").value & ""
            lblReservedNm.Caption = GetPtNm(txtReservedID.Text)
            cboCompo.tag = RS.Fields("compocd").value & ""
            If chkBar.value = 1 Then
                txtBldNo.Text = RS.Fields("bldsrc").value & "" & RS.Fields("bldyy").value & "" & Format(RS.Fields("bldno").value & "", "0#####")
            Else
                txtBldNo.Text = RS.Fields("bldsrc").value & "" & "-" & RS.Fields("bldyy").value & "" & "-" & Format(RS.Fields("bldno").value & "", "0#####")
            End If
            
            CurrentStatus = .Fields("stscd").value & ""
            
            If .Fields("stscd").value & "" > 5 Then
                fraBld.Enabled = False
                MsgBox "이미 헌혈자의 혈액이 등록되어있습니다.", vbExclamation
            End If
        End If
    End With
    
    Set RS = Nothing
End Sub

Private Sub LoadBldInfo()
'등록된 혈액은 정보 조회
    Dim strSQL As String
    Dim RS As Recordset
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim strBldno As String
    Dim strCompocd As String
    
    If chkBar.value = 1 Then
        strBldSrc = Mid(txtBldNo.Text, 1, 2)
        strBldYY = Mid(txtBldNo.Text, 3, 2)
        strBldno = Format(Mid(txtBldNo.Text, 5, 6), "00000#")
    Else
        strBldSrc = medGetP(txtBldNo.Text, 1, "-")
        strBldYY = medGetP(txtBldNo.Text, 2, "-")
        strBldno = Format(medGetP(txtBldNo.Text, 3, "-"), "######")
    End If
    strCompocd = cboCompo.tag
    
    strSQL = " select * from " & T_BBS401 & _
             " where " & DBW("bldsrc=", strBldSrc) & _
             " and " & DBW("bldyy=", strBldYY) & _
             " and " & DBW("bldno=", strBldno) & _
             " and " & DBW("compocd=", strCompocd)
             
    'compocd, bldno, volume, coldt, expdt, available
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    If RS.EOF = False Then
        cboCompo.ListIndex = medComboFind(cboCompo, RS.Fields("compocd").value & "")
        
        Select Case RS.Fields("volumn").value & ""
            Case "320"
                optVol(0).value = True
            Case "400"
                optVol(1).value = True
            Case "250"
                optVol(2).value = True
            Case Else
                optVol(3).value = True
                txtVolume.Text = RS.Fields("volumn").value & ""
        End Select
        dtpColDt.value = Format(RS.Fields("coldt").value & "", "####-##-##")
        dtpExpDt.value = Format(RS.Fields("expdt").value & "", "####-##-##")
        lblAvailable.Caption = RS.Fields("available").value & ""
    End If
    
    Set RS = Nothing
End Sub

Private Sub LoadCompo()
'제제 조회
    Dim strSQL As String
    Dim RS As Recordset
    
    If optDonorCd(0).value Then
        strSQL = " select * from " & T_BBS006 & " where compocd>' ' and pherefg='0' and (expdt is null or expdt ='') "
    ElseIf optDonorCd(1).value Then
        strSQL = " select * from " & T_BBS006 & " where compocd>' ' and pherefg='1' and (expdt is null or expdt ='') "
    End If
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    cboCompo.Clear
    Do Until RS.EOF
        cboCompo.AddItem RS.Fields("compocd").value & "" & COL_DIV & RS.Fields("componm").value & ""
        
        RS.MoveNext
    Loop
    
    cboCompo.ListIndex = -1
    
    Set RS = Nothing
End Sub

Private Sub chkBar_Click()
    On Error Resume Next
    txtBldNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Dim strDonorid As String
    Dim strDonoraccdt As String
    
    If txtDonorNm.Text = "" Then Exit Sub
    
    If CurrentStatus = 0 Then
        MsgBox "헌혈자가 등록되지 않았습니다.", vbExclamation
        Exit Sub
    End If

    If CurrentStatus < CANEDIT_STATUS Then
        MsgBox "등록된 혈액이 없습니다. 등록된 혈액만 등록취소 할 수 있습니다.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("혈액 등록을 취소하시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    strDonorid = lblDonorID.Caption
    strDonoraccdt = Format(cboDonoraccdt.Text, PRESENTDATE_FORMAT)

On Error GoTo ErrTrap
    DBConn.BeginTrans
    
    If optDonorCd(0).value Then
        If CancelBlood(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
    ElseIf optDonorCd(1).value Then
        If CancelBlood(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
'        If CancelPheresis(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
    End If
    
    DBConn.CommitTrans
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    Call cmdClear_Click
    
    Exit Sub
ErrTrap:
    DBConn.RollbackTrans
    MsgBox "정상적으로 처리되지 않았습니다.", vbExclamation
End Sub

Private Function CancelBlood(ByVal vDonorid As String, ByVal vDonoraccdt As String) As Boolean
    Dim objSql As clsBBSSQLStatement
    Dim RS As Recordset
    Dim strSql1 As String
    Dim strSql2 As String
    Dim arySql(3) As String
    Dim i As Long
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim strBldno As String
    Dim strCompocd As String
    
    If chkBar.value = 1 Then
        strBldSrc = Mid(txtBldNo.Text, 1, 2)
        strBldYY = Mid(txtBldNo.Text, 3, 2)
        strBldno = Val(Format(Mid(txtBldNo.Text, 5, 6), "00000#"))
    Else
        strBldSrc = medGetP(txtBldNo.Text, 1, "-")
        strBldYY = medGetP(txtBldNo.Text, 2, "-")
        strBldno = Val(Format(medGetP(txtBldNo.Text, 3, "-"), "######"))
    End If
    strCompocd = medGetP(cboCompo.Text, 1, COL_DIV)
    
    Set objSql = New clsBBSSQLStatement
    strSql1 = objSql.GetStorageHistory(strBldSrc, strBldYY, strBldno, strCompocd)
    
    Set RS = New Recordset
    RS.Open strSql1, DBConn
    If Not RS.EOF Then
        If RS.Fields("stscd").value & "" Then
            Select Case RS.Fields("stscd").value & ""
                Case "1": MsgBox "반환처리되었던 혈액입니다. 취소할 수 없습니다.", vbExclamation
                Case "2": MsgBox "결과등록이 된 혈액입니다. 취소할 수 없습니다.", vbExclamation
                Case "3": MsgBox "출고처리된 혈액입니다. 취소할 수 없습니다.", vbExclamation
                Case "4": MsgBox "폐기처리된 혈액입니다. 취소할 수 없습니다.", vbExclamation
            End Select
            
            Set RS = Nothing
            Set objSql = Nothing
            Exit Function
        End If
    Else
        MsgBox "혈액의 입고내역이 없습니다. 취소할 수 없습니다.", vbExclamation
        Set objSql = Nothing
        Set RS = Nothing
        Exit Function
    End If
    Set RS = Nothing
    Set objSql = Nothing
    
    On Error GoTo ErrTrap
    
    strSql1 = " delete " & T_BBS401 & _
             " WHERE" & DBW("bldsrc", strBldSrc, 2) & _
             " AND  " & DBW("bldyy", strBldYY, 2) & _
             " AND  " & DBW("bldno", strBldno, 2) & _
             " AND  " & DBW("compocd", strCompocd, 2)
    
    DBConn.Execute strSql1
    
    strSql2 = " SELECT * FROM " & T_BBS401 & _
             " WHERE " & DBW("donorid", vDonorid, 2) & _
             " AND   " & DBW("donoraccdt", vDonoraccdt, 2)
    
    Set RS = New Recordset
    RS.Open strSql2, DBConn
    
    If RS.EOF Then
        arySql(0) = " update " & T_BBS602 & _
                    " set" & DBW("donationdt", "", 3) & _
                    DBW("bldsrc", "", 3) & _
                    DBW("bldyy", "", 3) & _
                    DBW("bldno", 0, 3) & _
                    DBW("compocd", "", 3) & _
                    DBW("volumn", 0, 3) & _
                    DBW("entfg", "0", 3) & _
                    DBW("cancelfg", "1", 2) & _
                    " WHERE" & DBW("donorid", vDonorid, 2) & _
                    " AND  " & DBW("donoraccdt", vDonoraccdt, 2)
        
    End If
    Set RS = Nothing
    
    arySql(1) = " update " & T_BBS603 & _
                " set    " & DBW("stscd", "5", 2) & _
                " WHERE " & DBW("donorid", vDonorid, 2) & _
                " AND   " & DBW("donoraccdt", vDonoraccdt, 2)
                
    arySql(2) = " update " & T_BBS411 & _
                " set  " & DBW("usedt", "", 3) & _
                           DBW("useid", "", 2) & _
                " WHERE" & "   " & DBW("bldsrc", strBldSrc, 2) & _
                " AND " & DBW("bldyy", strBldYY, 2) & _
                " AND " & DBW("bldno", strBldno, 2)
    
    For i = LBound(arySql) To UBound(arySql)
'        Debug.Print arySql(i)
        If arySql(i) <> "" Then DBConn.Execute arySql(i)
    Next
    CancelBlood = True
    Exit Function
    
ErrTrap:
    CancelBlood = False
End Function

Private Function CancelPheresis(ByVal vDonorid As String, ByVal vDonoraccdt As String) As Boolean
    Dim arySql(3) As String
    Dim i As Long
    
    arySql(0) = " update " & T_BBS602 & _
                " set" & DBW("donationdt", "", 3) & _
                DBW("bldsrc", "", 3) & _
                DBW("bldyy", "", 3) & _
                DBW("bldno", 0, 3) & _
                DBW("compocd", "", 3) & _
                DBW("volumn", 0, 3) & _
                DBW("entfg", "0", 3) & _
                DBW("cancelfg", "1", 2) & _
                " WHERE" & DBW("donorid", vDonorid, 2) & _
                " AND  " & DBW("donoraccdt", vDonoraccdt, 2)

    arySql(1) = " update " & T_BBS603 & " " & _
                " set    " & DBW("stscd", "5", 2) & _
                " WHERE " & DBW("donorid", vDonorid, 2) & _
                " AND   " & DBW("donoraccdt", vDonoraccdt, 2)
           
'    arySql(2) = " update " & T_BBS411 & _
'                " set  " & DBW("usedt", "", 3) & _
'                           DBW("useid", "", 2) & _
'                " WHERE" & "   " & DBW("bldsrc", strBldSrc, 2) & _
'                " AND " & DBW("bldyy", strBldYY, 2) & _
'                " AND " & DBW("bldno", strBldNo, 2)
           
    On Error GoTo ErrTrap
    For i = LBound(arySql) To UBound(arySql)
'        Debug.Print arySql(i)
        If arySql(i) <> "" Then DBConn.Execute arySql(i)
    Next
    CancelPheresis = True
    Exit Function
    
ErrTrap:
    CancelPheresis = False
End Function
Private Sub cmdClear_Click()
    txtDonorNm.Text = ""
    Call InitDonor
    Call InitInfo
    On Error Resume Next
    txtDonorNm.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS423 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim strDonorid As String
    Dim strDonoraccdt As String
    
    If CheckValidation = False Then Exit Sub
    
    strDonorid = lblDonorID.Caption
    strDonoraccdt = Format(cboDonoraccdt.Text, PRESENTDATE_FORMAT)
    
On Error GoTo ErrTrap
    DBConn.BeginTrans
    
    If optDonorCd(0).value Then
        If SaveBlood(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
    ElseIf optDonorCd(1).value Then
        If SavePheresis(strDonorid, strDonoraccdt) = False Then GoTo ErrTrap
    End If
    
    DBConn.CommitTrans
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    '혈액Tag 바코드 출력
    Call TagPrint(GetBldNo, medGetP(cboCompo.Text, 1, COL_DIV))
    Call cmdClear_Click
    
    Exit Sub
ErrTrap:
    DBConn.RollbackTrans
    MsgBox "정상적으로 처리되지 않았습니다.", vbExclamation
End Sub

Private Function CheckValidation() As Boolean
    CheckValidation = False
    
    If txtDonorNm.Text = "" Then
        MsgBox "헌혈자 성명을 입력하십시오.", vbExclamation
        txtDonorNm.SetFocus
        Exit Function
    End If
    
    If CurrentStatus = 0 Then
        MsgBox "헌혈자가 등록되지 않았습니다. 등록된 헌혈자만 혈액등록을 할 수 있습니다.", vbExclamation
        Exit Function
    End If

    If CurrentStatus > CANEDIT_STATUS Then
        MsgBox "헌혈자의 혈액이 이미 등록되어 있습니다.", vbExclamation
        Exit Function
    End If
    
    If cboDonoraccdt.ListIndex < 0 Then
        MsgBox "등록일자를 선택하십시오.", vbExclamation
        cboDonoraccdt.SetFocus
        Exit Function
    End If
    
    If cboCompo.ListIndex < 0 Then
        MsgBox "제제를 선택하십시오.", vbExclamation
        cboCompo.SetFocus
        Exit Function
    End If
    
    If txtBldNo.Text = "" Then
        MsgBox "혈액번호를 입력하십시오.", vbExclamation
        txtBldNo.SetFocus
        Exit Function
    End If
    
    If txtVolume.Text = "" Then
        MsgBox "용량을 선택하거나 입력하십시오.", vbExclamation
        txtVolume.SetFocus
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Function SaveBlood(ByVal vDonorid As String, ByVal vDonoraccdt As String) As Boolean
'insert into s2bbs401( bldsrc,bldyy,bldno,compocd,volumn,abo,rh,ptid,reserved,autofg,pherefg,coldt,coltm,colid,available, expdt,exptm,entdt,enttm,entid,centercd,localcd,stscd,hosfg,splitoutfg,splitinfg, realexpdt,realexptm,expid,exprcvid,expbilldiv,exprsnrmk,donorid,donoraccdt)values ('20','04',123456,'01',320,'AB','+','00404245','1','0','0','20041230','134937','9999',35,'20050203','134937','20041230','134937','9999','10','','0','1','','0','','',0,'0','','','1','20041223')
'update s2bbs602 set     donationdt  ='20041230', bldsrc  ='20', bldyy  ='04', bldno   = 123456, compocd  ='01', volumn   = 320, entfg  ='1', cancelfg  ='0' WHERE   donorid  ='1' AND  donoraccdt  ='20041223'
'update s2bbs603 set     stscd  ='6' WHERE   donorid  ='1' AND  donoraccdt  ='20041223'
'update s2bbs411 set   usedt  ='20041230', useid  ='9999' WHERE    bldsrc  ='20' AND  bldyy  ='04' AND  bldno   = 123456
'update s2bbs602 set  cancelfg  ='0' WHERE  donorid = '1' AND  donoraccdt = '20041223'
'update s2bbs603  set     stscd  ='6' WHERE  donorid  ='1' AND    donoraccdt  ='20041223'
    Dim arySql(6) As String
    Dim i As Long
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim strBldno As String
    Dim strCompocd As String
    Dim strVolume As String
    Dim strABO As String
    Dim strRh As String
    Dim strPtid As String
    Dim strRFg As String 'ReservedFg
    Dim strAFg As String 'AutoFg
    Dim strPFg As String 'PhereFg
    Dim strDate As String
    Dim strTime As String
    Dim strId As String
    Dim strAvailable As String
    Dim strExpDt As String
    
    If chkBar.value = 1 Then
        strBldSrc = Mid(txtBldNo.Text, 1, 2)
        strBldYY = Mid(txtBldNo.Text, 3, 2)
        strBldno = Format(Mid(txtBldNo.Text, 5, 6), "00000#")
    Else
        strBldSrc = medGetP(txtBldNo.Text, 1, "-")
        strBldYY = medGetP(txtBldNo.Text, 2, "-")
        strBldno = Format(medGetP(txtBldNo.Text, 3, "-"), "######")
    End If
    strCompocd = medGetP(cboCompo.Text, 1, COL_DIV)
    strVolume = txtVolume.Text
    If Len(lblABO.Caption) > 2 Then
        strABO = Mid(lblABO.Caption, 1, 2)
        strRh = Mid(lblABO.Caption, 3)
    Else
        strABO = Mid(lblABO.Caption, 1, 1)
        strRh = Mid(lblABO.Caption, 2, 1)
    End If
    strPtid = txtReservedID.Text
    strId = ObjMyUser.EmpId
    strDate = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strTime = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strAvailable = lblAvailable.Caption
    strExpDt = Format(dtpExpDt.value, PRESENTDATE_FORMAT)
    
    arySql(0) = " insert into " & T_BBS401 & "(" & _
                    " bldsrc,bldyy,bldno,compocd,volumn,abo,rh,ptid,reserved,autofg,pherefg,coldt,coltm,colid,available," & _
                    " expdt,exptm,entdt,enttm,entid,centercd,localcd,stscd,hosfg,splitoutfg,splitinfg," & _
                    " realexpdt,realexptm,expid,exprcvid,expbilldiv,exprsnrmk,donorid,donoraccdt)" & _
                    "values (" & _
                    DBV("bldsrc", strBldSrc, 1) & DBV("bldyy", strBldYY, 1) & DBV("bldno", strBldno, 1) & DBV("compocd", strCompocd, 1) & DBV("volumn", strVolume, 1) & _
                    DBV("abo", strABO, 1) & DBV("rh", strRh, 1) & DBV("ptid", strPtid, 1) & DBV("reserved", "1", 1) & DBV("autofg", "0", 1) & DBV("pherefg", "0", 1) & _
                    DBV("coldt", strDate, 1) & DBV("coltm", strTime, 1) & DBV("colid", strId, 1) & _
                    DBV("available", strAvailable, 1) & DBV("expdt", strExpDt, 1) & DBV("exptm", "", 1) & DBV("entdt", strDate, 1) & _
                    DBV("enttm", strTime, 1) & DBV("entid", strId, 1) & DBV("centercd", "10", 1) & DBV("localcd", "", 1) & _
                    DBV("stscd", "0", 1) & DBV("hosfg", "1", 1) & DBV("splitoutfg", "", 1) & DBV("splitinfg", "0", 1) & DBV("realexpdt", "", 1) & DBV("realexptm", "", 1) & _
                    DBV("expid", 0, 1) & DBV("exprcvid", 0, 1) & DBV("expbilldiv", "", 1) & DBV("exprsnrmk", "", 1) & DBV("donorid", vDonorid, 1) & DBV("donoraccdt", vDonoraccdt) & ")"

    arySql(1) = " update " & T_BBS602 & _
                " set    " & DBW("donationdt", strDate, 3) & _
                             DBW("bldsrc", strBldSrc, 3) & _
                             DBW("bldyy", strBldYY, 3) & _
                             DBW("bldno", strBldno, 3) & _
                             DBW("compocd", strCompocd, 3) & _
                             DBW("volumn", strVolume, 3) & _
                             DBW("entfg", "1", 3) & _
                             DBW("cancelfg", "0", 2) & _
                " WHERE  " & DBW("donorid", vDonorid, 2) & _
                " AND " & DBW("donoraccdt", vDonoraccdt, 2)

    arySql(2) = " update " & T_BBS603 & _
                " set    " & DBW("stscd", "6", 2) & _
                " WHERE  " & DBW("donorid", vDonorid, 2) & _
                " AND " & DBW("donoraccdt", vDonoraccdt, 2)

    arySql(3) = " update " & T_BBS411 & _
                " set  " & DBW("usedt", strDate, 3) & _
                           DBW("useid", strId, 2) & _
                " WHERE" & "   " & DBW("bldsrc", strBldSrc, 2) & _
                " AND " & DBW("bldyy", strBldYY, 2) & _
                " AND " & DBW("bldno", strBldno, 2)

    arySql(4) = " update " & T_BBS603 & " " & _
                " set    " & DBW("stscd", "6", 2) & _
                " WHERE " & DBW("donorid", vDonorid, 2) & _
                " AND   " & DBW("donoraccdt", vDonoraccdt, 2)

    arySql(5) = "update " & T_BBS602 & " set " & DBW("cancelfg", "0", 2) & _
                " WHERE " & DBW("donorid=", vDonorid) & _
                " AND " & DBW("donoraccdt=", vDonoraccdt)
    
    On Error GoTo ErrTrap
    For i = LBound(arySql) To UBound(arySql)
        If arySql(i) <> "" Then DBConn.Execute arySql(i)
    Next
    SaveBlood = True
    Exit Function
    
ErrTrap:
    SaveBlood = False
End Function

Private Function SavePheresis(ByVal vDonorid As String, ByVal vDonoraccdt As String) As Boolean
' insert into s2bbs401( bldsrc,bldyy,bldno,compocd,volumn,abo,rh,ptid,reserved,autofg,pherefg,coldt,coltm,colid,available, expdt,exptm,entdt,enttm,entid,centercd,localcd,stscd,hosfg,splitoutfg,splitinfg, realexpdt,realexptm,expid,exprcvid,expbilldiv,exprsnrmk)values ('01','04',123456,'22',400,'A','B','00481772','0','0','1','20041230','115405','9999',0,'20041230','115405','20041230','115405','9999','10','','0','1','','','','',0,'0','','')
' update s2bbs602 set     donationdt  ='20041230', bldsrc  ='01', bldyy  ='04', bldno   = 123456, compocd  ='22', volumn   = 400, entfg  ='1', cancelfg  ='0' WHERE   donorid  ='2' AND  donoraccdt  ='20041230'
' update s2bbs603 set stscd  ='6' WHERE   donorid  ='2' AND  donoraccdt  ='20041230'
' update s2bbs411 set usedt  ='20041230', useid  ='9999' WHERE bldsrc  ='01' AND  bldyy  ='04' AND  bldno   = 123456
    Dim arySql(4) As String
    Dim i As Long
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim strBldno As String
    Dim strCompocd As String
    Dim strVolume As String
    Dim strABO As String
    Dim strRh As String
    Dim strPtid As String
    Dim strRFg As String 'ReservedFg
    Dim strAFg As String 'AutoFg
    Dim strPFg As String 'PhereFg
    Dim strDate As String
    Dim strTime As String
    Dim strId As String
    Dim strAvailable As String
    Dim strExpDt As String
    
    If chkBar.value = 1 Then
        strBldSrc = Mid(txtBldNo.Text, 1, 2)
        strBldYY = Mid(txtBldNo.Text, 3, 2)
        strBldno = Format(Mid(txtBldNo.Text, 5, 6), "00000#")
    Else
        strBldSrc = medGetP(txtBldNo.Text, 1, "-")
        strBldYY = medGetP(txtBldNo.Text, 2, "-")
        strBldno = Format(medGetP(txtBldNo.Text, 3, "-"), "######")
    End If
    strCompocd = medGetP(cboCompo.Text, 1, COL_DIV)
    strVolume = txtVolume.Text
    If Len(lblABO.Caption) > 2 Then
        strABO = Mid(lblABO.Caption, 1, 2)
        strRh = Mid(lblABO.Caption, 3)
    Else
        strABO = Mid(lblABO.Caption, 1, 1)
        strRh = Mid(lblABO.Caption, 2, 1)
    End If
    strPtid = txtReservedID.Text
    strId = ObjMyUser.EmpId
    strDate = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strTime = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strAvailable = lblAvailable.Caption
    strExpDt = Format(dtpExpDt.value, PRESENTDATE_FORMAT)
    
    arySql(0) = " insert into " & T_BBS401 & "(" & _
                    " bldsrc,bldyy,bldno,compocd,volumn,abo,rh,ptid,reserved,autofg,pherefg,coldt,coltm,colid,available," & _
                    " expdt,exptm,entdt,enttm,entid,centercd,localcd,stscd,hosfg,splitoutfg,splitinfg," & _
                    " realexpdt,realexptm,expid,exprcvid,expbilldiv,exprsnrmk,donorid,donoraccdt) " & _
                    "values (" & _
                    DBV("bldsrc", strBldSrc, 1) & DBV("bldyy", strBldYY, 1) & DBV("bldno", strBldno, 1) & DBV("compocd", strCompocd, 1) & _
                    DBV("volumn", strVolume, 1) & DBV("abo", strABO, 1) & DBV("rh", strRh, 1) & DBV("ptid", strPtid, 1) & _
                    DBV("reserved", "0", 1) & DBV("autofg", "0", 1) & DBV("pherefg", "1", 1) & _
                    DBV("coldt", strDate, 1) & DBV("coltm", strTime, 1) & DBV("colid", strId, 1) & _
                    DBV("available", strAvailable, 1) & DBV("expdt", strExpDt, 1) & DBV("exptm", "", 1) & _
                    DBV("entdt", strDate, 1) & DBV("enttm ", strTime, 1) & DBV("entid", strId, 1) & _
                    DBV("centercd", "10", 1) & DBV("localcd", "", 1) & DBV("stscd", "0", 1) & _
                    DBV("hosfg", "1", 1) & DBV("splitoutfg", "", 1) & DBV("splitinfg", "", 1) & _
                    DBV("realexpdt", "", 1) & DBV("realexptm", "", 1) & DBV("expid", 0, 1) & _
                    DBV("exprcvid", 0, 1) & DBV("expbilldiv", "", 1) & DBV("exprsnrmk", "", 1) & DBV("donorid", vDonorid, 1) & DBV("donoraccdt", vDonoraccdt) & ")"
    
    arySql(1) = " update " & T_BBS411 & _
                " set  " & DBW("usedt", strDate, 3) & _
                           DBW("useid", strId, 2) & _
                " WHERE" & DBW("bldsrc", strBldSrc, 2) & _
                " AND " & DBW("bldyy", strBldYY, 2) & _
                " AND " & DBW("bldno", strBldno, 2)

    arySql(2) = " update " & T_BBS603 & _
                " set    " & DBW("stscd", "6", 2) & _
                " WHERE  " & DBW("donorid   ", vDonorid, 2) & _
                " AND " & DBW("donoraccdt", vDonoraccdt, 2)

    arySql(3) = " update " & T_BBS602 & _
                " set    " & DBW("donationdt", strDate, 3) & _
                             DBW("bldsrc    ", strBldSrc, 3) & _
                             DBW("bldyy     ", strBldYY, 3) & _
                             DBW("bldno     ", strBldno, 3) & _
                             DBW("compocd   ", strCompocd, 3) & _
                             DBW("volumn    ", strVolume, 3) & _
                             DBW("entfg     ", "1", 3) & _
                             DBW("cancelfg  ", "0", 2) & _
                " WHERE  " & DBW("donorid", vDonorid, 2) & _
                " AND " & DBW("donoraccdt", vDonoraccdt, 2)
    On Error GoTo ErrTrap
    For i = LBound(arySql) To UBound(arySql)
        If arySql(i) <> "" Then DBConn.Execute arySql(i)
    Next
    SavePheresis = True
    Exit Function
    
ErrTrap:
    SavePheresis = False
End Function

Private Sub TagPrint(ByVal vBldNo As String, ByVal vCompo As String)
'헌혈자 Tag 출력
'혈액번호, 제제로 쿼리해서 출력한다.
    Dim strSQL As String
    Dim RS As Recordset
    Dim strBldSrc As String
    Dim strBldYY As String
    Dim strBldno As String
    Dim strCompo As String
    Dim aryData(1 To 11) As Variant
    
    strBldSrc = medGetP(vBldNo, 1, "-")
    strBldYY = medGetP(vBldNo, 2, "-")
    strBldno = medGetP(vBldNo, 3, "-")
    strCompo = vCompo
    
    strSQL = " select a.bldsrc,a.bldyy,a.bldno,a.compocd,b.abbrnm,a.volumn,a.abo||a.rh bldabo,a.ptid,d." & F_PTNM & " ptnm, a.coldt, a.expdt,a.donorid,c.donornm,c.abo||c.rh donorabo   "
    strSQL = strSQL & " from " & T_BBS401 & " a, " & T_BBS006 & " b, " & T_BBS601 & " c, " & T_HIS001 & " d "
    strSQL = strSQL & " where " & DBW("a.bldsrc=", strBldSrc)
    strSQL = strSQL & " and " & DBW("a.bldyy=", strBldYY)
    strSQL = strSQL & " and " & DBW("a.bldno=", strBldno)
    strSQL = strSQL & " and " & DBW("a.compocd=", strCompo)
    strSQL = strSQL & " and (a.reserved='1' or a.pherefg='1')"
    strSQL = strSQL & " and a.compocd=b.compocd"
    strSQL = strSQL & " and a.donorid=c.donorid"
    strSQL = strSQL & " and a.ptid=d.patno"
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.EOF Then
        MsgBox "출력할 내역이 없습니다.", vbExclamation
        Set RS = Nothing
        Exit Sub
    End If

'aryData(1):혈액번호, aryData(2):혈액제제, aryData(3):용량
'aryData(4):혈액형, aryData(5):지정환자ID, aryData(6):환자명
'aryData(7):헌혈일, aryData(8):유효일, aryData(9):헌혈자
'aryData(10):헌혈자혈액형, aryData(11):바코드용 혈액번호
    aryData(1) = vBldNo
    aryData(2) = RS.Fields("abbrnm").value & ""
    aryData(3) = RS.Fields("volumn").value & ""
    aryData(4) = RS.Fields("bldabo").value & ""
    aryData(5) = RS.Fields("ptid").value & ""
    aryData(6) = RS.Fields("ptnm").value & ""
    aryData(7) = Mid(Format(RS.Fields("coldt").value & "", "####/##/##"), 3)
    aryData(8) = Mid(Format(RS.Fields("expdt").value & "", "####/##/##"), 3)
    aryData(9) = RS.Fields("donornm").value & ""
    aryData(10) = RS.Fields("donorabo").value & ""
    aryData(11) = RS.Fields("bldsrc").value & "" & RS.Fields("bldyy").value & "" & Format(RS.Fields("bldno").value & "", "000000")
    
    PrintDonorLabel aryData()
    
    Set RS = Nothing
End Sub

Private Sub dtpColdt_Change()
    dtpExpDt.value = DateAdd("d", Val(lblAvailable.Caption) - 1, dtpColDt.value)
End Sub

Private Sub Form_Load()
    txtDonorNm.Text = ""
    Call InitDonor
    Call InitInfo
End Sub

Private Sub InitDonor()
    lblDonorID.Caption = ""
    lblDOB.Caption = ""
    lblSex.Caption = ""
    lblABO.Caption = ""
    lblCnt.Caption = ""
    lblTotVol.Caption = ""
    cboDonoraccdt.Clear
    lblAccCnt.Caption = ""
End Sub

Private Sub InitInfo()
    Dim i As Long
    
    CurrentStatus = 0
    
    optDonorCd(0).value = False
    optDonorCd(1).value = False
    txtReservedID.Text = ""
    lblReservedNm.Caption = ""
    fraBld.Enabled = True
    cboCompo.Clear
    cboCompo.tag = ""
    txtBldNo.Text = ""
    chkBar.value = 1
    For i = optVol.LBound To optVol.UBound
        optVol(i).value = False
    Next
    txtVolume.Text = ""
    dtpColDt.value = GetSystemDate
    dtpExpDt.value = GetSystemDate
    lblAvailable.Caption = ""
End Sub

Private Sub optVol_Click(Index As Integer)
    Select Case Index
        Case 0: txtVolume.Text = "320": txtVolume.Locked = True
        Case 1: txtVolume.Text = "400": txtVolume.Locked = True
        Case 2: txtVolume.Text = "250": txtVolume.Locked = True
        Case 3: txtVolume.Text = "": txtVolume.Locked = False
    End Select
End Sub

Private Sub txtBldNo_Change()
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
    Dim strBldno As String
    Dim strBNum As String
    
    If txtBldNo.Text = "" Then Exit Sub
    If cboCompo.ListIndex = -1 Then Exit Sub
    
    strBldno = GetBldNo
    
    strBNum = Replace(strBldno, "-", "")
    
    If CheckExist(strBNum) Then
        Cancel = True
        MsgBox "이미 입고된 혈액입니다.", vbExclamation
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtDonorNm_Change()
    If lblDonorID.Caption <> "" Then
        Call InitDonor
        Call InitInfo
    End If
End Sub

Private Sub txtDonorNm_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDonorNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtDonorNm.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDonorNm_Validate(Cancel As Boolean)
    If txtDonorNm.Text = "" Then Exit Sub
    If lblDonorID.Caption <> "" Then Exit Sub
    
    If DonorFind = False Then
        Cancel = True
        MsgBox "등록된 헌혈자가 아닙니다. 먼저 헌혈자를 등록하십시오.", vbExclamation
    Else
        Call ShowAccList
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Function DonorFind() As Boolean
    Dim objDonor As clsBBSBldDonationBusi
        
    Set objDonor = New clsBBSBldDonationBusi
    With objDonor
        DonorFind = .DonorFind(txtDonorNm.Text)
            
        lblDonorID.Caption = .mDonorID
'        txtDonorNm = .mDonorNm
        lblDOB.Caption = .mDOB
        lblSex.Caption = .mSEX
        lblABO.Caption = .mABO
        lblCnt.Caption = .Mcnt
        lblTotVol.Caption = .mTotVol
        lblSSN.Caption = .mSSN
    End With
    Set objDonor = Nothing
End Function

Private Sub ShowAccList()
    Dim strAccDt    As String
    Dim RS          As Recordset
    Dim objMySQL    As clsBBSSQLStatement
    '헌혈자에 대해서 접수된 정보가 있을 경우에 접수 내역을 보여준다.

    Set objMySQL = New clsBBSSQLStatement
    Set RS = New Recordset

    Set RS = objMySQL.GetDonorAccHistory(lblDonorID.Caption)
    
    cboDonoraccdt.Clear
    Do Until RS.EOF
        strAccDt = Format(RS.Fields("donoraccdt").value & "", "####-##-##")
        cboDonoraccdt.AddItem strAccDt
        RS.MoveNext
    Loop
    If cboDonoraccdt.ListCount > 0 Then
        lblAccCnt.Caption = cboDonoraccdt.ListCount
        cboDonoraccdt.ListIndex = 0
    Else
        MsgBox "등록된 헌혈자가 아닙니다. 먼저 헌혈자를 등록하십시오.", vbExclamation
    End If
    
    Set RS = Nothing
    Set objMySQL = Nothing
End Sub

