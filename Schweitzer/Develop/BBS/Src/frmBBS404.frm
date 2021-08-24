VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS404 
   BackColor       =   &H00DBE6E6&
   Caption         =   "헌혈 등록"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   Icon            =   "frmBBS404.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   14700
   WindowState     =   2  '최대화
   Begin MSComctlLib.TabStrip tabAccDt 
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Top             =   2025
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2000-01-01"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2280
      TabIndex        =   17
      Top             =   2310
      Width           =   9930
      Begin VB.ComboBox cboDonorCd 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "frmBBS404.frx":076A
         Left            =   1050
         List            =   "frmBBS404.frx":077A
         Locked          =   -1  'True
         Style           =   1  '단순 콤보
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Width           =   2055
      End
      Begin VB.TextBox txtReservedID 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00CFDCDE&
         Height          =   330
         Left            =   4260
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   225
         Width           =   1125
      End
      Begin MedControls1.LisLabel lblReservedNm 
         Height          =   330
         Left            =   5385
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   225
         Width           =   2640
         _ExtentX        =   4657
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
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   10
         Left            =   60
         TabIndex        =   39
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
         Caption         =   "헌혈종류"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   11
         Left            =   3270
         TabIndex        =   40
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
         Caption         =   "지정환자"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "헌혈취소"
      Height          =   510
      Left            =   6915
      Style           =   1  '그래픽
      TabIndex        =   9
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   8235
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   9555
      Style           =   1  '그래픽
      TabIndex        =   7
      Tag             =   "124"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   10875
      Style           =   1  '그래픽
      TabIndex        =   8
      Tag             =   "128"
      Top             =   7575
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      Top             =   1695
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
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2280
      TabIndex        =   30
      Top             =   480
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
      Caption         =   "  기 본 정 보"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   2280
      TabIndex        =   31
      Top             =   2910
      Width           =   9930
      Begin MedControls1.LisLabel lblStsNm 
         Height          =   315
         Left            =   1065
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblStsCd 
         Height          =   315
         Left            =   2310
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
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
      End
      Begin MedControls1.LisLabel lblOkDiv1Nm 
         Height          =   315
         Left            =   3600
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv1Cd 
         Height          =   315
         Left            =   4545
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
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
      End
      Begin MedControls1.LisLabel lblOkDiv2Nm 
         Height          =   315
         Left            =   5865
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv2Cd 
         Height          =   315
         Left            =   6810
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
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
      End
      Begin MedControls1.LisLabel lblOkDiv3Nm 
         Height          =   315
         Left            =   8130
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv3Cd 
         Height          =   315
         Left            =   9090
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
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
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   6
         Left            =   60
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   180
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
         Caption         =   "현재상태"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   7
         Left            =   2610
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   180
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
         Caption         =   "접수결과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   8
         Left            =   4860
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   180
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
         Caption         =   "문진결과"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   9
         Left            =   7125
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   180
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
         Caption         =   "검사결과"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraDonation 
      BackColor       =   &H00DBE6E6&
      Height          =   990
      Left            =   2280
      TabIndex        =   11
      Top             =   3450
      Width           =   9930
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   14
         Left            =   3615
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
         _ExtentX        =   1879
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
         Caption         =   "혈액량"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   15
         Left            =   3615
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   585
         Width           =   1065
         _ExtentX        =   1879
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
         Caption         =   "혈액번호"
         Appearance      =   0
      End
      Begin VB.CheckBox chkBar 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드"
         ForeColor       =   &H004A5580&
         Height          =   195
         Left            =   2805
         TabIndex        =   32
         Top             =   675
         Width           =   840
      End
      Begin VB.ComboBox cboBuilding 
         Height          =   300
         Left            =   7095
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   585
         Width           =   2520
      End
      Begin VB.OptionButton optVo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Etc"
         Height          =   270
         Index           =   2
         Left            =   6315
         TabIndex        =   16
         Top             =   270
         Width           =   675
      End
      Begin VB.OptionButton optVo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "400cc"
         Height          =   270
         Index           =   1
         Left            =   5505
         TabIndex        =   15
         Top             =   270
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optVo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "320cc"
         Height          =   270
         Index           =   0
         Left            =   4695
         TabIndex        =   14
         Top             =   270
         Width           =   795
      End
      Begin VB.TextBox txtBldNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4710
         MaxLength       =   13
         TabIndex        =   4
         Top             =   585
         Width           =   2100
      End
      Begin VB.ComboBox cboCompo 
         Height          =   300
         ItemData        =   "frmBBS404.frx":07A8
         Left            =   1140
         List            =   "frmBBS404.frx":07B5
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   225
         Width           =   2010
      End
      Begin VB.TextBox txtVolumn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7095
         MaxLength       =   10
         TabIndex        =   3
         Top             =   225
         Width           =   825
      End
      Begin MSComCtl2.DTPicker dtpDonationDt 
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   585
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   59834371
         CurrentDate     =   36797
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   12
         Left            =   60
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   225
         Width           =   1065
         _ExtentX        =   1879
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
         Caption         =   "혈액제제"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   13
         Left            =   60
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   585
         Width           =   1065
         _ExtentX        =   1879
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
         Caption         =   "헌혈일"
         Appearance      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "cc"
         Height          =   180
         Left            =   7950
         TabIndex        =   12
         Top             =   345
         Width           =   210
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2280
      TabIndex        =   21
      Top             =   720
      Width           =   9930
      Begin VB.TextBox txtDonorNm 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Top             =   165
         Width           =   1515
      End
      Begin MedControls1.LisLabel lblDOB 
         Height          =   330
         Left            =   4275
         TabIndex        =   22
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
         Left            =   6630
         TabIndex        =   23
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
         Left            =   8940
         TabIndex        =   24
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
         Left            =   4275
         TabIndex        =   25
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
         Left            =   6630
         TabIndex        =   26
         Top             =   510
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
         Left            =   1050
         TabIndex        =   27
         Top             =   540
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
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
         Left            =   1800
         TabIndex        =   28
         Top             =   540
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
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
         TabIndex        =   33
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
         Left            =   3285
         TabIndex        =   34
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
         Left            =   3285
         TabIndex        =   35
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
         Index           =   3
         Left            =   5640
         TabIndex        =   36
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
         Caption         =   "성/나이"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   5640
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   510
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   7950
         TabIndex        =   38
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
         Left            =   7575
         TabIndex        =   29
         Top             =   660
         Width           =   210
      End
   End
   Begin VB.Frame fraTest 
      BackColor       =   &H00DBE6E6&
      Height          =   3090
      Left            =   2280
      TabIndex        =   57
      Top             =   4380
      Width           =   9930
      Begin FPSpread.vaSpread tblSave 
         Height          =   2565
         Left            =   6855
         TabIndex        =   58
         Top             =   450
         Width           =   2985
         _Version        =   196608
         _ExtentX        =   5265
         _ExtentY        =   4524
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
         GridShowVert    =   0   'False
         MaxCols         =   4
         MaxRows         =   10
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS404.frx":07C9
         TextTip         =   4
      End
      Begin MedControls1.LisLabel lblTestChk 
         Height          =   315
         Left            =   30
         TabIndex        =   60
         Top             =   450
         Visible         =   0   'False
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   556
         BackColor       =   12632256
         ForeColor       =   16576
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
         Caption         =   "이미 검사의뢰된 헌혈자입니다."
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   30
         TabIndex        =   61
         Top             =   120
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   556
         BackColor       =   8388608
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
         BorderStyle     =   0
         Caption         =   "   검 사 항 목"
         Appearance      =   0
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   2190
         Left            =   15
         TabIndex        =   62
         Tag             =   "10114"
         Top             =   825
         Width           =   6780
         _Version        =   196608
         _ExtentX        =   11959
         _ExtentY        =   3863
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
         GridShowVert    =   0   'False
         MaxCols         =   20
         MaxRows         =   7
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS404.frx":0C38
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   7
      End
      Begin MedControls1.LisLabel lblTmpPtId 
         Height          =   330
         Left            =   5430
         TabIndex        =   63
         Top             =   450
         Width           =   1365
         _ExtentX        =   2408
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
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Left            =   6855
         TabIndex        =   64
         Top             =   120
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         BackColor       =   8388608
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
         BorderStyle     =   0
         Caption         =   "  헌혈혈액정보"
         Appearance      =   0
      End
      Begin MSComctlLib.TabStrip tabGroup 
         Height          =   345
         Left            =   15
         TabIndex        =   65
         Top             =   465
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   16
         Left            =   4725
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   450
         Width           =   690
         _ExtentX        =   1217
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
         Caption         =   "임시ID"
         Appearance      =   0
      End
      Begin VB.ListBox lstBldNo 
         BackColor       =   &H00F3F2E9&
         Height          =   1860
         Left            =   7200
         TabIndex        =   59
         Top             =   540
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmBBS404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Enum TblColumn
    tcSEL = 1
    tcName
    tcCODE
    tcQTY
End Enum
Private objMySQL As New clsBBSSQLStatement
Private objMyOrder As New clsDonorBusiOrder
Private objMyCollection As New clsDonorTestCollection
'2001-11-27추가
Private strSaveDonorId As String
Private strSaveDonorNm As String


Private Sub cboCompo_Click()
    If cboDonorCd.ListIndex <> 3 Then Exit Sub
    Dim objSql     As clsBBSSQLStatement
    Dim aryOrdCd() As String
    Dim today      As Date
    Dim Volumn     As String
    Dim CompoCd    As String
    Dim Cnt        As Long
    Dim i          As Long
    
    today = GetSystemDate
    Volumn = "0"
    Set objSql = New clsBBSSQLStatement
'    objSql.setDbConn DBConn
    CompoCd = medGetP(cboCompo.Text, 1, " ")
    Cnt = objSql.GetOrdCd(Volumn, CompoCd, Format(today, PRESENTDATE_FORMAT), aryOrdCd)
    Set objSql = Nothing
    
'    cboNewTest.Clear
'    If cnt > 0 Then
'        For i = 1 To cnt
'            cboNewTest.AddItem aryOrdCd(i - 1)
'        Next i
'        cboNewTest.ListIndex = 0
'    End If
End Sub

Private Sub cmdCancel_Click()
'헌혈취소(602에 cancelfg="1",401 레코드 삭제,lab102(Dcfg='1')

    Dim objSql      As clsBBSSQLStatement
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim BldSrc      As String
    Dim BldYY       As String
    Dim BldNo       As String
    Dim CompoCd     As String
    Dim donorid     As String
    Dim donoraccdt  As String
    Dim tmpptid     As String
    Dim strTmp      As String
    
    If cboDonorCd.ListIndex = 3 Then
        MsgBox "Pheresis 헌혈은 취소 하실수 없습니다.", vbInformation + vbOKOnly, "헌혈취소"
        Exit Sub
    End If
    strTmp = MsgBox(txtBldNo.Text & "  " & medGetP(cboCompo.Text, 2, " ") & vbCrLf & " 혈액을 취소하시겠습니까?", vbYesNo + vbInformation, "Info")
    If strTmp = vbNo Then Exit Sub
    
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    donorid = lblDonorID.Caption
    CompoCd = medGetP(cboCompo.Text, 1, " ")
    BldSrc = medGetP(txtBldNo, 1, "-")
    BldYY = medGetP(txtBldNo, 2, "-")
    BldNo = medGetP(txtBldNo, 3, "-")
    tmpptid = lblTmpPtId.Caption
    
    Set objSql = New clsBBSSQLStatement
    SSQL = objSql.GetStorageHistory(BldSrc, BldYY, CLng(BldNo), CompoCd)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        If RS.Fields("stscd").value & "" Then
            Select Case RS.Fields("stscd").value & ""
                Case "1": MsgBox "반환처리되었던 혈액입니다.취소하실수 없습니다.", vbInformation + vbOKOnly, "Info"
                Case "2": MsgBox "결과등록이 된 혈액입니다.취소하실수 없습니다.", vbInformation + vbOKOnly, "Info"
                Case "3": MsgBox "출고처리된 혈액입니다.취소하실수 없습니다.", vbInformation + vbOKOnly, "Info"
                Case "4": MsgBox "폐기처리된 혈액입니다.취소하실수 없습니다.", vbInformation + vbOKOnly, "Info"
            End Select
            Set objSql = Nothing
            Set RS = Nothing
            Exit Sub
        End If
    Else
        MsgBox "관련혈액정보가 없습니다.취소하실수 없습니다.", vbInformation + vbOKOnly, "Info"
        Set RS = Nothing
        Set objSql = Nothing
        Exit Sub
    End If
    
    Set objSql = New clsBBSSQLStatement
    
    If objSql.SetBldCancel(donorid, donoraccdt, tmpptid, BldSrc, BldYY, BldNo, CompoCd) Then
        MsgBox "헌혈등록이 취소되었습니다.", vbInformation + vbOKOnly, "헌혈취소"
        FrameInitialize
    End If
    
    Set objSql = Nothing
End Sub

Private Sub cmdClear_Click()
    FrameInitialize
End Sub

Private Sub cmdExit_Click()
    Set objMySQL = Nothing
    Set objMyOrder = Nothing
    Set objMyCollection = Nothing
    Unload Me
End Sub


Private Sub cmdSave_Click()
'입력한 혈액번호가 입고내역에 존재한다면, 저장 하면 않된다.
    Dim Resp As VbMsgBoxResult
    If Bld_Check(txtBldNo) = False Then Exit Sub
    
    If SaveAll = True Then
        Resp = MsgBox("혈액입고를 계속하시겠습니까?", vbQuestion + vbYesNo, "헌혈자등록")
        If Resp = vbNo Then
            Call FrameInitialize
        Else
            Call tabAccDt_Click
        End If
    End If
End Sub
Private Function Bld_Check(ByVal BldNum As String) As Boolean
    Dim objSql As clsBBSSQLStatement
    Dim BldSrc As String
    Dim BldYY  As String
    Dim BldNo  As String
    Dim CompoCd As String
    
            
'    If lblOkDiv3Cd.Caption = "" Then
'        MsgBox "검사결과가 없으므로 헌혈등록을 하실수 없습니다.", vbInformation + vbOKOnly, "검사결과 미판정"
'        Exit Function
'    ElseIf lblOkDiv3Cd.Caption <> "1" Then
'        MsgBox "검사결과가 부적격 판정이므로 헌혈등록을 하실수 없습니다.", vbInformation + vbOKOnly, "검사결과 부적격"
'        Exit Sub
'    End If
'
'    If txtBldNo = "" Or cboCompo.ListIndex < 0 Then
'        MsgBox "혈액정보를 입력후 작업을 진행하십시요", vbInformation + vbOKOnly, "정보누락"
'        Exit Function
'    End If
    
     


    If chkBar.value <> 1 Then
        BldSrc = medGetP(BldNum, 1, "-")
        BldYY = medGetP(BldNum, 2, "-")
        BldNo = medGetP(BldNum, 3, "-")
    Else
        BldSrc = Mid(BldNum, 1, 2)
        BldYY = Mid(BldNum, 3, 2)
        BldNo = Mid(BldNum, 5, 6)
    End If
    CompoCd = medGetP(cboCompo.Text, 1, " ")
    
    If BldSrc = "" Or BldYY = "" Or BldNo = "" Then
        MsgBox "혈액번호 입력오류 입니다. 확인후 등록하세요.", vbInformation + vbOKOnly, "Info"
        txtBldNo.Text = ""
        txtBldNo.SetFocus
        Exit Function
    Else
        If Len(BldSrc) <> 2 Or Len(BldYY) <> 2 Then
            MsgBox "혈액번호 입력오류 입니다. 확인후 등록하세요.", vbInformation + vbOKOnly, "Info"
            txtBldNo.Text = ""
            txtBldNo.SetFocus
            Exit Function
        End If
    
    End If
        
    
    Set objSql = New clsBBSSQLStatement
    If objSql.GetBloodCheck(BldSrc, BldYY, BldNo, CompoCd) = True Then
        Bld_Check = True
    Else
        MsgBox "이미 입고된 혈액번호입니다. 확인후 등록하세요", vbInformation + vbOKOnly, "헌혈등록"
    End If
    Set objSql = Nothing
End Function

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
'    lblTestChk.Visible = False
End Sub

Private Sub Form_Load()
    Dim objDonorTest As clsDonorTest
    Dim strGroup()   As String
    Dim iCnt         As Long
    Dim i            As Long
    
    
    Set objDonorTest = New clsDonorTest
    iCnt = objDonorTest.GetGroup(strGroup)
    
    tabGroup.Tabs.Clear
    For i = 1 To iCnt
        tabGroup.Tabs.Add , strGroup(i - 1), strGroup(i - 1)
    Next i
    
    Set objDonorTest = Nothing


'    fraKIT.Left = 0
'    fraKIT.Top = 0
    dtpDonationDt = GetSystemDate
    
    'Call SetCboCompo
    Call SetMaterial
    Call FrameInitialize
    Call ClassInitialize
    '2001-12-07 추가
    '건물정보를 사용할 경우 건물리스트 로드
    If ObjSysInfo.UseBuildingInfo Then
        cboBuilding.Visible = True
        Call LoadBuilding
    Else
        cboBuilding.Visible = False
    End If
    
End Sub

'2001-12-07 추가
Private Sub LoadBuilding()
    
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim i As Long
    Dim iTmx As ListItem
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_CENTER)
    Set objcom003 = Nothing
    
    cboBuilding.Clear
    If Not DrRS.EOF Then
        With DrRS
            For i = 1 To .RecordCount
                cboBuilding.AddItem .Fields("cdval1").value & "" & " " & .Fields("field1").value & ""
                .MoveNext
            Next i
        End With
    End If
    Set DrRS = Nothing
    If cboBuilding.ListCount > 1 Then
        cboBuilding.ListIndex = medComboFind(cboBuilding, ObjSysInfo.BuildingCd)
    Else
        cboBuilding.ListIndex = 0
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objMySQL = Nothing
    Set objMyOrder = Nothing
    Set objMyCollection = Nothing
End Sub

Private Sub optVo_Click(Index As Integer)
    If Index = 2 Then
        txtVolumn.Enabled = True
    Else
        txtVolumn.Enabled = False
    End If

End Sub

Private Sub tabAccDt_Click()
    
    Dim donorid As String
    Dim canEdit As Boolean
    
    donorid = lblDonorID.Caption
    Call tabAccdtClickCode(donorid, Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT))
    Call SetDonorStatus(donorid, Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT))
    'Call SetDonorMaterial
    
    '---------------------------------------------
    '적격/부적격 판정이 났어도 헌혈등록이 가능하게
    '불가능이면 풀어주면 된다(2001/08/17)
    '---------------------------------------------
'    canEdit = GetCanEdit
'    fraDonation.Enabled = canEdit

End Sub

Private Function GetCanEdit() As Boolean
    '수정이 가능한지를 판단한다.
    If tabAccDt.SelectedItem.Index > 1 Then
        '최종 접수일자가 아니다. 수정할 수 없다.
        GetCanEdit = False
    Else
        Select Case lblStsCd.Caption
            Case DonorStatus.stsAccessSave
                GetCanEdit = False
            Case DonorStatus.stsAccessVerify
                GetCanEdit = False
            Case DonorStatus.stsAskSave
                GetCanEdit = False
            Case DonorStatus.stsAskVerify
                GetCanEdit = (lblOkDiv2Cd.Caption = "1")
            Case DonorStatus.stsDonation
                GetCanEdit = True
            Case DonorStatus.stsFinish
                GetCanEdit = False
            Case DonorStatus.stsPrint
                GetCanEdit = False
            Case Else
                GetCanEdit = False
        End Select
    End If
End Function

Private Sub SetDonorStatus(ByVal donorid As String, ByVal accdt As String)
    Dim objDonor As clsBBSSQLStatement
    Dim strStatus As String
    Dim IsPhere As Boolean
    
    
    Set objDonor = New clsBBSSQLStatement
    strStatus = objDonor.GetDonorStatus(donorid, accdt, IsPhere)
    Set objDonor = Nothing
    
    lblStsNm.Caption = medGetP(strStatus, 1, vbTab)
    lblStsCd.Caption = medGetP(strStatus, 2, vbTab)
    lblOkDiv1Nm.Caption = medGetP(strStatus, 3, vbTab)
    lblOkDiv1Cd.Caption = medGetP(strStatus, 4, vbTab)
    lblOkDiv2Nm.Caption = medGetP(strStatus, 5, vbTab)
    lblOkDiv2Cd.Caption = medGetP(strStatus, 6, vbTab)
    lblOkDiv3Nm.Caption = medGetP(strStatus, 7, vbTab)
    lblOkDiv3Cd.Caption = medGetP(strStatus, 8, vbTab)
    
    If lblOkDiv1Nm.Caption = "부적격" Then
        lblOkDiv1Nm.ForeColor = vbRed
        lblOkDiv1Cd.ForeColor = vbRed
    Else
        lblOkDiv1Nm.ForeColor = vbBlack
        lblOkDiv1Cd.ForeColor = vbBlack
    End If
    
    If lblOkDiv2Nm.Caption = "부적격" Then
        lblOkDiv2Nm.ForeColor = vbRed
        lblOkDiv2Cd.ForeColor = vbRed
    Else
        lblOkDiv2Nm.ForeColor = vbBlack
        lblOkDiv2Cd.ForeColor = vbBlack
    End If
    
    If lblOkDiv3Nm.Caption = "부적격" Then
        lblOkDiv3Nm.ForeColor = vbRed
        lblOkDiv3Cd.ForeColor = vbRed
    Else
        lblOkDiv3Nm.ForeColor = vbBlack
        lblOkDiv3Cd.ForeColor = vbBlack
    End If
    
End Sub

Private Sub tabGroup_Click()
    Dim NewTest       As Recordset
    Dim strGroup      As String
    
    '검사의뢰가 되지 않은 환자에 대해서는 검사항목 마스터에등록된 검사항목을 보여준다.
    If tabAccDt.Tabs.Count = 0 Then
        Exit Sub
    End If
    
    strGroup = tabGroup.SelectedItem.Key
    
    
    Set NewTest = objMySQL.GetTestSpc2(strGroup)
    If Not NewTest.EOF Then
        Dim ObjDic As New clsDictionary
        Dim lngseq As Long

        ObjDic.Clear
        ObjDic.FieldInialize "seq", "testcd,spccd"
        Do Until NewTest.EOF
            lngseq = lngseq + 1
            ObjDic.AddNew lngseq, Join(Array(NewTest.Fields("cdval2").value & "", NewTest.Fields("field1").value & ""), COL_DIV)
            NewTest.MoveNext
        Loop
        lblTestChk.Visible = False
        Call Default_Test(ObjDic)
        Set NewTest = Nothing
        Set ObjDic = Nothing
        cmdSave.Enabled = True
        cmdCancel.Enabled = False
    End If
End Sub




Private Sub txtDonorNm_GotFocus()
    txtDonorNm.tag = txtDonorNm
End Sub

Private Sub txtDonorNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call DonorFind
        txtDonorNm.tag = txtDonorNm
    End If
End Sub

Private Sub txtDonorNm_LostFocus()
    If txtDonorNm.tag <> txtDonorNm Then
        Call DonorFind
    End If
End Sub

Private Sub DonorFind()
    Dim objDonor As clsBBSBldDonationBusi
    
    If txtDonorNm = "" Then Call FrameInitialize: Exit Sub
    
    Set objDonor = New clsBBSBldDonationBusi
    With objDonor

        If .DonorFind(txtDonorNm) = True Then
            Call FrameInitialize
            
            lblDonorID.Caption = .mDonorID
            txtDonorNm = .mDonorNm
            '2001-11-27 추가
            strSaveDonorId = lblDonorID.Caption
            strSaveDonorNm = txtDonorNm.Text
            '
            lblDOB.Caption = .mDOB
            lblSex.Caption = .mSEX
            lblABO.Caption = .mABO
            lblCnt.Caption = .Mcnt
            lblTotVol.Caption = .mTotVol
        
            Call ShowAccList
        End If
    End With
    Set objDonor = Nothing
End Sub

Private Sub ShowAccList()
    Dim strAccDt As String
    Dim RS As Recordset
    Dim objMySQL As clsBBSSQLStatement
    '헌혈자에 대해서 접수된 정보가 있을 경우에 접수 내역을 보여준다.

    Set objMySQL = New clsBBSSQLStatement

'    objMySQL.setDbConn DBConn
    'Set Rs = objMySQL.GetDonorAccHistory(Trim(lblDonorID.Caption))
    
    '성분헌혈을 제외한 헌혈에 대해서만 조회할수 있게 추가.
    '2001/10/04 울산동강병원
    Set RS = objMySQL.GetDonorAccdtHistoryDivPheresis(Trim(lblDonorID.Caption))
    
    
    If RS.EOF Then
        tabAccDt.Tabs.Clear
        tabAccDt.Visible = False
        MsgBox "헌혈등록 대상이 없습니다.", vbInformation + vbOKOnly, "헌혈등록"
        
    Else
        tabAccDt.Tabs.Clear
        tabAccDt.Visible = True
        
        Do Until RS.EOF
            strAccDt = Format(RS.Fields("donoraccdt").value & "", "####-##-##")
            tabAccDt.Tabs.Add , , strAccDt
            RS.MoveNext
        Loop
        
        cmdSave.Enabled = True
        Call tabAccDt_Click
    End If

End Sub

Private Sub FrameInitialize()
    dtpDonationDt = GetSystemDate
    tabAccDt.Tabs.Clear
    tabAccDt.Visible = False
    
    cboDonorCd.ListIndex = -1
    txtReservedID.Text = ""
    lblReservedNm.Caption = ""
    
    lblStsNm.Caption = ""
    lblStsCd.Caption = ""
    lblOkDiv1Nm.Caption = ""
    lblOkDiv1Cd.Caption = ""
    lblOkDiv2Nm.Caption = ""
    lblOkDiv2Cd.Caption = ""
    lblOkDiv3Nm.Caption = ""
    lblOkDiv3Cd.Caption = ""
    
'    fraKIT.Visible = False
    
    medClearTable tblResult
    lblTmpPtId.Caption = ""
    txtDonorNm = ""
    lblDonorID.Caption = ""
    lblSex.Caption = ""
    lblABO.Caption = ""
    lblCnt.Caption = ""
    lblTotVol.Caption = ""
    lblDOB.Caption = ""
    lblTestChk.Visible = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    txtVolumn.Enabled = False
    medClearTable tblSave
    optVo(0).value = True
    
    Clear
    lstBldNo.Clear
End Sub

Private Sub Clear()
    Dim r As Long
    
    cboCompo.ListIndex = -1
    txtVolumn = ""
    txtBldNo = ""
End Sub

Private Sub SetCboCompo(ByVal TF As Boolean)
    'tf:t이면pheresis 혈액제제
    Dim objCompo  As New clsBBSSQLStatement
    Dim RS        As New Recordset
    Dim i         As Integer
    

    Set RS = objCompo.Compolist(TF)
    
    If Not RS.EOF Then
        cboCompo.Clear
        For i = 1 To RS.RecordCount
            Do Until RS.EOF
                cboCompo.AddItem RS.Fields("compocd").value & "" & " " & RS.Fields("abbrnm").value & "" & Space(20) & COL_DIV & RS.Fields("keepday").value & ""
                RS.MoveNext
            Loop
        Next i
    End If
    
    Set RS = Nothing
    Set objCompo = Nothing
End Sub

Private Sub SetMaterial()
'    Dim objcom003 As clsCom003
'    Dim DrRS As RECORDSET
'    Dim i As Long
'
'    tblMaterial.MaxRows = 0
'
'    Set objcom003 = New clsCom003
'
'    Set DrRS = objcom003.OpenRecordSet( BC2_MATERIAL)
'    If DrRS Is Nothing Then Exit Sub
'
'    With tblMaterial
'        .MaxRows = DrRS.RecordCount
'        For i = 1 To DrRS.RecordCount
'            .Row = i
'            .Col = TblColumn.tcSEL:  .value = 0
'            .Col = TblColumn.tcCODE: .value = DrRS.Fields("cdval1")
'            .Col = TblColumn.TcName: .value = DrRS.Fields("field1")
'
'            DrRS.MoveNext
'        Next i
'    End With
'
'    Set DrRS = Nothing
'    Set objcom003 = Nothing
End Sub

Private Sub SetDonorMaterial()
    Dim donorid As String
    Dim donoraccdt As String
    Dim objDonorMaterial As clsDonorMaterial
    Dim DrRS As Recordset
    Dim i As Long
    Dim r As Long
    
    Dim RsTestReq As Recordset
    Clear
   
    donorid = lblDonorID.Caption
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    
    
    Set objMySQL = New clsBBSSQLStatement
    With objMySQL
'        .setDbConn DBConn
        Set RsTestReq = .GetDonorAccHistory(donorid, donoraccdt)
    End With
    
    If Not RsTestReq.EOF Then
        '임시환자id
'        lblTmpPtId = RsTestReq.Fields("tmpid")
        
        '헌혈종류
        Select Case RsTestReq.Fields("donorcd").value & ""
            Case "0": cboDonorCd.ListIndex = 0
            Case "1": cboDonorCd.ListIndex = 1
            Case "2": cboDonorCd.ListIndex = 2
            Case "3": cboDonorCd.ListIndex = 3
        End Select
        txtReservedID = RsTestReq.Fields("reservedid").value & ""
    End If
    Set RsTestReq = Nothing
    Set objMySQL = Nothing
    
    
    '--------------------------------------------------
    '이 접수일에 등록된 헌혈내역과 재료내역을 불러온다.
    '--------------------------------------------------
    
    
    Set objDonorMaterial = New clsDonorMaterial
    '------------------------------------------------------------------------
    '헌 혈 내 역
    '------------------------------------------------------------------------
    Set DrRS = objDonorMaterial.GetDonorDonation(donorid, donoraccdt)
    If DrRS Is Nothing Then Exit Sub
    With DrRS
        If .RecordCount > 0 Then
            cboCompo.ListIndex = medComboFind(cboCompo, .Fields("compocd").value & "" & " " & .Fields("componm").value & "")
            txtVolumn = .Fields("volumn").value & "" & ""
            If Trim(.Fields("donationdt").value & "") <> "" Then
                dtpDonationDt = Format(.Fields("donationdt").value & "", "####-##-##")
            End If
            If Trim(.Fields("bldsrc").value & "") <> "" Then
                txtBldNo = .Fields("bldsrc").value & "" & "-" & .Fields("bldyy").value & "" & "-" & .Fields("bldno").value & ""
            End If
        End If
    End With
    Set DrRS = Nothing
    
    '------------------------------------------------------------------------
    '재 료 내 역
    '------------------------------------------------------------------------
'    Set DrRS = objDonorMaterial.GetDonorMaterial(Donorid, DonorAccdt)
'    If DrRS Is Nothing Then Exit Sub
'
'    With tblMaterial
'        For i = 1 To DrRS.RecordCount
'            For r = 1 To .MaxRows
'                .Row = r
'                .Col = TblColumn.tcCODE
'                If Trim(.value) = Trim(DrRS.Fields("ordcd")) Then
'                    .Col = TblColumn.tcSEL: .value = 1
'                    .Col = TblColumn.tcQTY: .value = DrRS.Fields("qty")
'
'                    Exit For
'                End If
'            Next r
'            DrRS.MoveNext
'        Next i
'    End With
'
'    Set DrRS = Nothing
    Set objDonorMaterial = Nothing
End Sub
'Private Sub Used_Material(ByVal donorid, ByVal donoraccdt As String)
'   '------------------------------------------------------------------------
'   ' 재 료 내 역
'   ' ------------------------------------------------------------------------
'   Dim DrRS             As New RECORDSET
'   Dim objDonorMaterial As New clsDonorMaterial
'   Dim i                As Integer
'   Dim r                As Integer
'
'    Set DrRS = objDonorMaterial.GetDonorMaterial(donorid, donoraccdt)
'    If DrRS Is Nothing Then Exit Sub
'
'    With tblMaterial
'        For i = 1 To DrRS.RecordCount
'            For r = 1 To .MaxRows
'                .Row = r
'                .Col = TblColumn.tcCODE
'                If Trim(.value) = Trim(DrRS.Fields("ordcd")) Then
'                    .Col = TblColumn.tcSEL: .value = 1
'                    .Col = TblColumn.tcQTY: .value = DrRS.Fields("qty")
'                End If
'            Next r
'            DrRS.MoveNext
'        Next i
'    End With
'
'    Set DrRS = Nothing
'    Set objDonorMaterial = Nothing
'End Sub
Private Function Save() As Boolean
    
    Dim objSql         As clsBBSSQLStatement
    Dim RS             As Recordset
    Dim BldSrc         As String
    Dim BldYY          As String
    Dim BldNo          As String
    Dim CompoCd        As String
    Dim Volumn         As String
    Dim ABO            As String
    Dim Rh             As String
    Dim PtId           As String
    Dim RFg            As String
    Dim AFg            As String
    Dim PFg             As String
    Dim ExpDt          As String
    Dim Dt             As String
    Dim Tm             As String
    Dim id             As String
    Dim CenterCd       As String
    Dim stscd          As String
    
    Dim donorid        As String
    Dim DonationDt     As String
    Dim donoraccdt     As String
    Dim Available      As Long
    '수가내역작성시....
    Dim ObjDic         As clsDictionary
    Dim DeliveryHold   As String
    Dim strTmp         As String
    Dim Orderptid      As String
    Dim orddt          As String
    Dim ordno          As String
    Dim Ordseq         As String
    Dim FilterFg       As String
    Dim IrradFg        As String
    Dim ordcd          As String
    Dim Ostscd         As String
    Dim MaterialCd     As String
    Dim MateriaQty     As String
    
    Dim Bordcd         As String
    Dim accdt          As String
    Dim accseq         As String
    Dim Method         As String
    
    Dim ii             As Integer
    
    If optVo(2).value = True Then
        If Trim(txtVolumn) = "" Then
            MsgBox "Volumn을 입력하십시요.", vbCritical, Me.Caption
            Save = False
            Exit Function
        End If
    End If
    If cboCompo.ListIndex < 0 Then
        MsgBox "혈액제제를 선택하십시요.", vbCritical, Me.Caption
        Save = False
        Exit Function
    End If
    If cboDonorCd.ListIndex = 3 Then
'        If cboNewTest.ListIndex < 0 Then
'            MsgBox "처방코드를 선택하세요.", vbCritical + vbOKOnly, "헌혈등록"
'            Exit Function
'        End If
    End If
    
    Set ObjDic = New clsDictionary
    Set objSql = New clsBBSSQLStatement
    Set RS = New Recordset
'    objSql.setDbConn DBConn
    
    ObjDic.Clear
    ObjDic.FieldInialize "ptid,orddt,ordno,ordseq,ordcd,div", "unitqty"
    
    If chkBar.value <> 1 Then
        BldSrc = medGetP(txtBldNo, 1, "-")
        BldYY = medGetP(txtBldNo, 2, "-")
        BldNo = medGetP(txtBldNo, 3, "-")
    Else
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
        BldNo = Mid(txtBldNo, 5, 6)
    End If
    
    If optVo(0).value = True Then
        Volumn = "320"
    ElseIf optVo(1).value = True Then
        Volumn = "400"
    Else
        Volumn = txtVolumn
    End If
    
    Select Case cboDonorCd.ListIndex
        Case "0":   cboDonorCd.ListIndex = 0: PtId = "":            RFg = "0": AFg = "0": PFg = "0"
        Case "1":   cboDonorCd.ListIndex = 1: PtId = txtReservedID: RFg = "1": AFg = "0": PFg = "0"
        Case "2":   cboDonorCd.ListIndex = 2: PtId = txtReservedID: RFg = "0": AFg = "1": PFg = "0"
        Case "3":   cboDonorCd.ListIndex = 3: PtId = txtReservedID: RFg = "0": AFg = "0": PFg = "1": 'Method = cboMethod.ListIndex
    End Select
    
    If cboDonorCd.ListIndex = 3 Then
'        'pheresis 헌혈인경우....추가재료도 등록한다....(출고까지....
'        Set Rs = objSql.Get_OrdInformation(lblTmpPtId.Caption)
'        If Not Rs.EOF Then
'            Orderptid = Rs.Fields("ptid")
'            orddt = Rs.Fields("orddt")
'            ordno = Rs.Fields("ordno")
'            Ordseq = Rs.Fields("ordseq")
'            ordcd = Rs.Fields("ordcd")
'            IrradFg = Rs.Fields("irradfg")
'            FilterFg = Rs.Fields("filterfg")
'            accdt = medGetP(lblTmpPtId.Caption, 1, "-")
'            accseq = medGetP(lblTmpPtId.Caption, 2, "-")
'        End If
'        '추가재료내역을 담는다.....
'        objdic.AddNew Join(Array(Orderptid, orddt, ordno, Ordseq, ordcd, 1), COL_DIV), 1
'        With tblMaterial
'            For ii = 1 To .MaxRows
'                .Row = ii
'                .Col = TblColumn.tcSEL
'                If .value = 1 Then
'                    .Col = TblColumn.tcCODE: MaterialCd = .value
'                    .Col = TblColumn.tcQTY:  MateriaQty = .value
'                    objdic.AddNew Join(Array(Orderptid, orddt, ordno, Ordseq, MaterialCd, 2), COL_DIV), MateriaQty
'                End If
'            Next ii
'        End With
'        Volumn = "0"
'        Ostscd = BBSOrderStatus.stsEnd
'        '출고등록,결과등록....
'        '출고등록체크시(출고까지 진행)
'        '결과등록체크시(Assign까지 진행)
'        '결과등록,출고등록 동시체크시(출고까지 진행)
'        '결과등록,출고등록 둘다 체크 않할때(입고까지만 진행)
'
'        If chkDelivery.value = 1 And chkResult.value = 1 Then
'            stscd = BBSBloodStatus.stsDELIVERY
'            DeliveryHold = "1"
'        ElseIf chkDelivery.value = 1 And chkResult.value = 0 Then
'            stscd = BBSBloodStatus.stsDELIVERY
'            DeliveryHold = "1"
'        ElseIf chkDelivery.value = 0 And chkResult.value = 1 Then
'            stscd = BBSBloodStatus.stsASSIGN
'            DeliveryHold = "0"
'        ElseIf chkDelivery.value = 0 And chkResult.value = 0 Then
'            stscd = BBSBloodStatus.stsENTER
'            DeliveryHold = "0"
'        End If
'
'        '수가 계산내역의 처방코드
'        Bordcd = medGetP(cboNewTest.Text, 1, " ")
'        Set Rs = Nothing
    Else
    '입고까지만 진행
        FilterFg = ""
        IrradFg = ""
        Bordcd = ""
        Ostscd = ""
        accdt = ""
        accseq = ""
        stscd = BBSBloodStatus.stsENTER
    End If
    
    donorid = lblDonorID.Caption                                    '헌혈자Id
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)  '헌혈접수일
    DonationDt = Format(dtpDonationDt.value, PRESENTDATE_FORMAT)            '헌혈일
    CompoCd = medGetP(cboCompo.Text, 1, " ")                        '혈액제제코드
    Available = Val(medGetP(cboCompo.Text, 2, COL_DIV))             '유효일
    If Len(lblABO.Caption) > 2 Then
        ABO = Mid(lblABO.Caption, 1, 2)
        Rh = Mid(lblABO.Caption, 3)
    Else
        ABO = Mid(lblABO.Caption, 1, 1)                                 '혈액형
        Rh = Mid(lblABO.Caption, 2, 1)                                  'rh
    End If
    id = ObjMyUser.EmpId                                            '담당자ID
    Dt = Format(GetSystemDate, PRESENTDATE_FORMAT)                      '모든등록일
    Tm = Format(GetSystemDate, PRESENTTIME_FORMAT)                        '모든등록시간
    
    '헌혈인경우 입고 시 CENTER지정
    If ObjSysInfo.UseBuildingInfo = "1" Then
        CenterCd = medGetP(cboBuilding.Text, 1, " ")    ' ObjSysInfo.BuildingCd                                '센터코드
    Else
        CenterCd = ObjSysInfo.BuildingCd                                '센터코드
    End If
    
    ExpDt = DateAdd("d", Available, dtpDonationDt.value)            '
    ExpDt = Format(ExpDt, PRESENTDATE_FORMAT)                               '폐기일(유효일과 계산)
    
    
    Save = objSql.Set_BldEnter(BldSrc, BldYY, BldNo, CompoCd, Volumn, ABO, Rh, PtId, _
                                RFg, AFg, PFg, Dt, Tm, id, Available, ExpDt, Tm, Dt, Tm, id, _
                                CenterCd, stscd, donorid, DonationDt, donoraccdt, Ostscd, _
                                IrradFg, FilterFg, Bordcd, accdt, accseq, DeliveryHold, Method, ObjDic)
    
    Set ObjDic = Nothing
    Set objSql = Nothing
End Function

'20010212 검사의뢰내역에서 옮겨온 코드를
Private Sub txtBldNo_Change()
    Dim lngLen As Long
    
    If chkBar.value = 1 Then Exit Sub
    
    
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

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If chkBar.value = 1 Then Exit Sub
    
    If Len(txtBldNo) <> 3 Or Len(txtBldNo) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
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

Private Sub tabAccdtClickCode(ByVal donorid As String, ByVal donoraccdt As String)
    Dim RsDonorTest   As Recordset
    Dim RsTestReq     As Recordset
    Dim QueryTest     As Recordset
    Dim NewTest       As Recordset
    Dim ii            As Integer
    
    With tblResult
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    '헌혈자에 대해서 임상병리에 검사의뢰를 한경우는 제외된다.
    If tabAccDt.SelectedItem.Selected Then
        Set objMySQL = New clsBBSSQLStatement
'        objMySQL.setDbConn DBConn
        
        '헌혈자 접수내역을 읽는다.--------------------------------------
        Set RsTestReq = objMySQL.GetDonorAccHistory(donorid, donoraccdt)
        If RsTestReq.EOF Then
'            'dbconn.DisplayErrors
            Set objMySQL = Nothing
            Exit Sub
        End If
        
        If RsTestReq.RecordCount < 1 Then
            MsgBox "접수내역을 찾을 수 없습니다.", vbCritical, "오류"
            Set RsTestReq = Nothing
            Set objMySQL = Nothing
            Exit Sub
        End If
        
        '접수정보 Display-----------------------------------------------
        
        
        Select Case RsTestReq.Fields("donorcd").value & ""
            Case "0":   cboDonorCd.ListIndex = 0
                        
            Case "1":   cboDonorCd.ListIndex = 1
            Case "2":   cboDonorCd.ListIndex = 2
            Case "3":   cboDonorCd.ListIndex = 3
            Case Else:  cboDonorCd.ListIndex = -1
        End Select
        
        'Pheresis인경우, 이 화면을 사용할 수 없다.
        If cboDonorCd.ListIndex = 3 Then
'            fraKIT.Visible = True
            cmdSave.Enabled = False
            cmdCancel.Enabled = False
            Exit Sub
        End If
        
        'Pheresis가 아닌 경우만 가능
'        fraKIT.Visible = False
        lblTmpPtId.Caption = RsTestReq.Fields("tmpid").value & ""
        SetCboCompo False
        txtReservedID = RsTestReq.Fields("reservedid").value & ""
        lblReservedNm.Caption = objMySQL.GetPtntNm(txtReservedID)
        
        
        '검사의뢰내역을 읽는다-----------------------------------------
        Set RsDonorTest = objMySQL.Get_TestHistory(donorid, donoraccdt)
        If RsDonorTest.EOF Then
'            'dbconn.DisplayErrors
            Set objMySQL = Nothing
            Exit Sub
        End If
        
        
        If RsDonorTest.RecordCount < 1 Then
            '검사의뢰가 되지 않은 환자에 대해서는 검사항목 마스터에등록된 검사항목을
            '무조건 보여준다.
            Set NewTest = objMySQL.GetTestSpc
            If Not NewTest.EOF Then
                Dim ObjDic As New clsDictionary
                Dim lngseq As Long

                ObjDic.Clear
                ObjDic.FieldInialize "seq", "testcd,spccd"
                Do Until NewTest.EOF
                    lngseq = lngseq + 1
                    ObjDic.AddNew lngseq, Join(Array(NewTest.Fields("cdval1"), NewTest.Fields("field1")), COL_DIV)
                    NewTest.MoveNext
                Loop
                lblTestChk.Visible = False
                Call Default_Test(ObjDic)
                Set NewTest = Nothing
                Set ObjDic = Nothing
                txtBldNo = "": txtVolumn = ""
                cmdSave.Enabled = True
                cmdCancel.Enabled = False
            End If
        Else
            '검사의뢰내역을 조회하여 보여준다.
            '이미 검사의뢰가 진행된 상태의 헌혈자정보
            
            If RsTestReq.Fields("donationdt").value & "" <> "" Then
                dtpDonationDt.value = Format(RsTestReq.Fields("donationdt").value & "", "####-##-##")
                For ii = 0 To cboCompo.ListCount
                    If medGetP(cboCompo.List(ii), 1, " ") = RsTestReq.Fields("compocd").value & "" Then
                        cboCompo.ListIndex = ii
                        Exit For
                    End If
                Next
                Select Case RsTestReq.Fields("volumn").value & ""
                    Case "320": optVo(0).value = True
                    Case "400": optVo(1).value = True
                    Case Else:  optVo(2).value = True: txtVolumn = RsTestReq.Fields("volumn").value & ""
                End Select
                txtBldNo = RsTestReq.Fields("bldsrc").value & "" & "-" & RsTestReq.Fields("bldyy").value & "" & "-" & RsTestReq.Fields("bldno").value & ""
                
                '2001-11-27수정 : 검사의뢰 되었어도 혈액입고는 하나 이상 할수 있게...
                'cmdSave.Enabled = False
                cmdSave.Enabled = True

                cmdCancel.Enabled = True
            Else
                txtBldNo = "": txtVolumn = "" '--> **추가**
                cmdSave.Enabled = True
                cmdCancel.Enabled = False
            End If
                        
            Set QueryTest = objMySQL.GetDonorTestDt(donorid, donoraccdt)
            Dim strTmpID As String
            lblTmpPtId.Caption = RsTestReq.Fields("tmpid").value & ""
            strTmpID = QueryTest.Fields("tmpid").value & ""
            'h7lab102에서 검사의뢰 정보를 불러온다.
            lblTestChk.Visible = True
            
            Call QueryInformation(strTmpID)
            Set QueryTest = Nothing
        End If
        
        Set RsDonorTest = Nothing
        Set RsTestReq = Nothing
        Set objMySQL = Nothing
    End If
    
    '2001-11-27 추가
    '해당접수일에 해당헌혈자에게 부여된 혈액번호를 리스트업한다.
    Call DonorBloodList(donorid, donoraccdt)
        
End Sub
Private Sub DonorBloodList(ByVal donorid As String, ByVal donoraccdt As String)
    Dim objSql  As clsBBSSQLStatement
    Dim RS      As Recordset
    Dim SSQL    As String
    
    Set objSql = New clsBBSSQLStatement
    Call medClearTable(tblSave)
    cmdCancel.Enabled = False
    SSQL = objSql.GetDonorBldList(donorid, donoraccdt)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tblSave
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .value = RS.Fields("bldsrc").value & "" & "-" & RS.Fields("bldyy").value & "" & "-" & Format(RS.Fields("bldno").value & "", "000000")
                .Col = 2: .value = RS.Fields("componm").value & ""
                .Col = 3: .value = RS.Fields("volumn").value & ""
                .Col = 4: .value = RS.Fields("compocd").value & ""
                RS.MoveNext
            Loop
        End With
        Call tblSave_Click(1, 1)
        cmdCancel.Enabled = True
    End If
    Set RS = Nothing
    Set objSql = Nothing
End Sub

Private Sub tblSave_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    With tblSave
        .Row = Row
        .Col = 1:
        If .value = "" Then Exit Sub
        txtBldNo.Text = .value
        .Col = 4:
'        Debug.Print cboCompo.ListIndex
        cboCompo.ListIndex = medComboFind(cboCompo, .value)
'        Debug.Print cboCompo.ListIndex
    End With
End Sub
'Private Sub XM_Method(ByVal strTmp As String)
''접수번호를 가지고 해당 처방에 대한 검사 방법을 가지고온다......
'    Dim objSql As New clsBBSSQLStatement
'    Dim accdt  As String
'    Dim accseq As String
'
'    accdt = medGetP(strTmp, 1, "-")
'    accseq = medGetP(strTmp, 2, "-")
'
''    objSql.setDbConn DBConn
''    cboMethod.ListIndex = objSql.Get_XMethod(accdt, accseq)
'
'    Set objSql = Nothing
'End Sub
Private Sub QueryInformation(tmpid As String)
'임시환자id와 처방번호를 가지고 검사정보를 조회한다.
'테스트 컬럼 정보
'1:처방번호,2:검사명,3:검사코드,4:검체,5:급여,6:응급여부,7:희망채열일시
'8:응급여부:9:WorkArea,10:storecd,11:rndfg,12;testdiv,13:multifg,14:spcgrp,15:ordseq
'16:약어명,17:바코드출력장수,18:검사가능여부,19:Location,20:검사장소
    Dim objQueryTest As New clsBBSSQLStatement
    Dim objDicT As New clsDictionary
    Dim objDicD As New clsDictionary
    Dim RsDonorTest As Recordset
    Dim RsDisplay As Recordset
    Dim strTmp As String
    
    Set objMySQL = New clsBBSSQLStatement
'    objMySQL.setDbConn DBConn
    
    objDicT.Clear
    objDicT.FieldInialize "ptid,orddt,ordno,ordseq", "ordcd,spccd,reqdate"
    
    
    Set RsDonorTest = objMySQL.GetDonnorTest(tmpid)
    
    If Not RsDonorTest.EOF Then
        Do Until RsDonorTest.EOF
            objDicT.AddNew Join(Array(RsDonorTest.Fields("ptid").value & "", RsDonorTest.Fields("orddt").value & "", RsDonorTest.Fields("ordno").value & "", _
                                RsDonorTest.Fields("ordseq").value & ""), COL_DIV), Join(Array(RsDonorTest.Fields("ordcd").value & "", _
                                RsDonorTest.Fields("spccd").value & "", RsDonorTest.Fields("reqdt").value & "" & RsDonorTest.Fields("reqtm").value & ""), COL_DIV)
            RsDonorTest.MoveNext
        Loop
    End If
    
    
    If objDicT.RecordCount > 0 Then
        objDicD.Clear
        objDicD.FieldInialize "ptid,orddt,ordno,ordseq", "ordno1,testnm,testcd,spccd,gubyu,stat,reqdt,statfg,workarea," & _
                              "storecd,rndfg,testdiv,multifg,spcgrp,ordseq1,abbrnm5,labelcnt,statflag,location,testlocation"
        objDicT.MoveFirst
        Do Until objDicT.EOF
            strTmp = objDicT.Fields("ordcd") & vbTab & objDicT.Fields("spccd")
            Set RsDisplay = objMySQL.GetTestFindList(strTmp)
            
            '2001/10/28: 마스터에서 잘못설정되어있을경우의 처리위해서
            If Not RsDisplay.EOF Then
                With RsDisplay
                    objDicD.AddNew Join(Array(objDicT.Fields("ptid"), objDicT.Fields("orddt"), objDicT.Fields("ordno"), objDicT.Fields("ordseq")), COL_DIV), _
                                   Join(Array(objDicT.Fields("ordno"), .Fields("testnm").value & "", .Fields("testcd").value & "", .Fields("spccd").value & "", _
                                              "1", "", objDicT.Fields("reqdate"), .Fields("statfg").value & "", .Fields("workarea").value & "", _
                                              .Fields("storecd").value & "", .Fields("rndfg").value & "", .Fields("testdiv").value & "", .Fields("multifg").value & "", _
                                              .Fields("spcgrp").value & "", objDicT.Fields("ordseq"), .Fields("abbrnm5").value & "", _
                                              .Fields("labelcnt").value & "", .Fields("statflags").value & "", "location", "중앙"), COL_DIV)
                End With
            End If
            
            objDicT.MoveNext
        Loop
    End If
    '화면에 보여주자......
    Call TblResult_Display(objDicD)
    '''
    
    Set objDicD = Nothing
End Sub
Private Sub TblResult_Display(ObjDic As clsDictionary)
'테스트 컬럼 정보
'1:처방번호,2:검사명,3:검사코드,4:검체,5:급여,6:응급여부,7:희망채열일시
'8:응급여부:9:WorkArea,10:storecd,11:rndfg,12;testdiv,13:multifg,14:spcgrp,15:ordseq
'16:약어명,17:바코드출력장수,18:검사가능여부,19:Location,20:검사장소
    Dim ii As Integer
    Dim tmpStatFg As String
    Dim tmpTestFg As String
    
    
    If ObjDic.RecordCount < 1 Then Exit Sub
    With tblResult
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .Action = ActionClear
        .BlockMode = False
        
        ObjDic.MoveFirst
        Do Until ObjDic.EOF
            If .DataRowCnt = .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            Else
                .Row = .DataRowCnt + 1
            End If
            .Col = 1: .value = ObjDic.Fields("ordno1")
            .Col = 2: .value = ObjDic.Fields("testnm")
            .Col = 3: .value = ObjDic.Fields("testcd")
            .Col = 4: .value = ObjDic.Fields("spccd")
            .Col = 5: .value = ObjDic.Fields("gubyu")
            If ObjDic.Fields("statfg") = "1" Then
                .Col = 6: .CellType = CellTypeCheckBox
                   .TypeCheckCenter = True
            Else
                .Col = 6: .CellType = CellTypeStaticText
            End If
            If Len(ObjDic.Fields("reqdt")) = 14 Then
                .Col = 7: .value = Format(Mid(ObjDic.Fields("reqdt"), 1, 12), "####-##-## ##:##")
            Else
                .Col = 7: .value = ObjDic.Fields("reqdt")
            End If
            .Col = 8: .value = ObjDic.Fields("statfg")
            .Col = 9: .value = ObjDic.Fields("workarea")
            .Col = 10: .value = ObjDic.Fields("storecd")
            .Col = 11: .value = ObjDic.Fields("rndfg")
            .Col = 12: .value = ObjDic.Fields("testdiv")
            .Col = 13: .value = ObjDic.Fields("multifg")
            .Col = 14: .value = ObjDic.Fields("spcgrp")
            .Col = 15: .value = ObjDic.Fields("ordseq1")
            .Col = 16: .value = ObjDic.Fields("abbrnm5")
            .Col = 17: .value = ObjDic.Fields("labelcnt")
            
            tmpStatFg = medGetP(ObjDic.Fields("statflag"), 1, ";")  '건물별 응급가능 여부
            tmpTestFg = medGetP(ObjDic.Fields("statflag"), 2, ";")  '건물별 검사가능 여부
            If ObjDic.Fields("statfg") = "1" Then
                .Col = 18: .value = Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1)
                If .value = "1" Then
                    If ObjSysInfo.BuildingCd = "10" Or ObjSysInfo.BuildingCd = "40" Then
                        .Col = 19: .value = "50"
                        .Col = 20: .value = "응급"
                    Else
                        .Col = 19: .value = ObjSysInfo.BuildingCd
                        .Col = 20: .value = ObjSysInfo.BuildingNm
                    End If
                Else
                    If ObjSysInfo.BuildingCd = "20" Or ObjSysInfo.BuildingCd = "30" Then
                        If Mid(tmpStatFg, 5, 1) = "1" Then
                            .Col = 19: .value = "50"
                            .Col = 20: .value = "응급"
                        Else
                        End If
                    Else
                        .Col = 18: .value = Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1)
                        If .value = "1" Then
                            .Col = 19: .value = ObjSysInfo.BuildingCd
                            .Col = 20: .value = ObjSysInfo.BuildingNm
                        Else
                            .Col = 19: .value = "10"
                            .Col = 20: .value = "중앙"
                        End If
                        .Col = 8: .value = "0"
                    End If
                End If
            Else
                .Col = 18: .value = Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1)
                If .value = "1" Then
                    .Col = 19: .value = ObjSysInfo.BuildingCd
                    .Col = 20: .value = ObjSysInfo.BuildingNm
                Else
                    .Col = 19: .value = "10"
                    .Col = 20: .value = "중앙"
                End If
            End If
            
            ObjDic.MoveNext
        Loop
    End With
    
            
End Sub
Private Sub Default_Test(objDefault As clsDictionary)
    Dim objQueryTest As New clsBBSSQLStatement
    Dim objGDic As New clsDictionary
    Dim DefaultTest As Recordset
    Dim strTmp As String
    Dim lngseq As Long
    
'    objQueryTest.setDbConn DBConn
    objGDic.Clear
    objGDic.FieldInialize "seq", "ordno1,testnm,testcd,spccd,gubyu,stat,reqdt,statfg,workarea," & _
                          "storecd,rndfg,testdiv,multifg,spcgrp,ordseq1,abbrnm5,labelcnt,statflag,location,testlocation"
    objDefault.MoveFirst
    
    Do Until objDefault.EOF
        
        strTmp = objDefault.Fields("testcd") & vbTab & objDefault.Fields("spccd")
        Set DefaultTest = Nothing
        Set DefaultTest = New Recordset
        DefaultTest.Open objQueryTest.GetDefaultTestList(strTmp), DBConn
        With DefaultTest
            If Not DefaultTest.EOF Then
                lngseq = lngseq + 1
                objGDic.AddNew lngseq, _
                               Join(Array("", .Fields("testnm").value & "", .Fields("testcd").value & "", .Fields("spccd").value & "", _
                                          "1", "", Format(GetSystemDate, "yyyy-MM-dd" & " " & "hh:MM"), .Fields("statfg").value & "", .Fields("workarea").value & "", _
                                          .Fields("storecd").value & "", .Fields("rndfg").value & "", .Fields("testdiv").value & "", .Fields("multifg").value & "", _
                                          .Fields("spcgrp").value & "", "", .Fields("abbrnm5").value & "", _
                                          .Fields("labelcnt").value & "", .Fields("statflags").value & "", "location", "중앙"), COL_DIV)
            End If
        End With
        objDefault.MoveNext
    Loop
    
    '화면에 보여주자......
    Call TblResult_Display(objGDic)
    Set objGDic = Nothing
    Set objQueryTest = Nothing
End Sub

Private Function SaveAll() As Boolean
    SaveAll = SaveDonation
End Function

'Private Function SavePheresis() As Boolean
'    Dim strOrdDt As String
'    Dim strWorkArea As String
'    Dim strAccdt As String
'    Dim lngAccSeq As Long
'    Dim blnSuccess As Boolean
'    Dim objSql As clsBBSSQLStatement
'
'    Dim donorid As String
'    Dim accdt As String
'
'    donorid = lblDonorID.Caption
'    accdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
'
'On Error GoTo ErrSave
'
'    '----- Begin Transaction -----
'    DbConn.BeginTrans
'
'   ' 혈액입고내역 생성
'    If Save = False Then GoTo ErrSave
''----- Commit Transaction -----
'
''    Set objSql = New clsBBSSQLStatement
''    Call objSql.SetDonorStatus(donorid, accdt, DonorStatus.stsDonation, False)
''    Set objSql = Nothing
'
'    DbConn.CommitTrans
'    SavePheresis = True
'    MsgBox "정상적으로 처리되었습니다.", vbInformation, "정보확인"
'
'    Call ClassInitialize
'    'Call FormInitialize
'
'    Exit Function
'
'ErrSave:
''----- Rollback Transaction -----
'    DbConn.RollbackTrans
'    'dbconn.DisplayErrors
'
'    Call ClassInitialize
'
'    SavePheresis = False
'End Function

Private Function SaveDonation() As Boolean
    Dim strOrdDt    As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim lngAccSeq   As Long
    Dim blnSuccess  As Boolean
    Dim objSql      As clsBBSSQLStatement
    
    Dim donorid     As String
    Dim accdt       As String
    Dim ii          As Integer
    
    donorid = lblDonorID.Caption
    accdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    
On Error GoTo ErrOther
    DBConn.BeginTrans
   ' 혈액입고내역 생성
    If Save = False Then GoTo ErrSave
    If tblResult.DataRowCnt = 0 Then
        'MsgBox "검사의뢰할 항목이 없습니다.", vbInformation, "정보확인"
        SaveDonation = False
        GoTo Skip
    End If
            
    Call TblSort
    
    '이미 검사의뢰된 항목에 대해서는 추가로 검사의뢰내역을 만들지 않는다.
    If lblTestChk.Visible = False Then
        '처방 루틴
        If SaveOrder = False Then GoTo ErrOther
        '채혈 루틴
        Call ReadyToCollect              '채혈준비
        If objMyCollection.DoCollection = False Then GoTo ErrOther    '채혈수행
    End If
    
 '----- Begin Transaction -----
   
On Error GoTo ErrSave
    '이미 검사의뢰된 항목에 대해서는 추가로 검사의뢰내역을 만들지 않는다.
    If lblTestChk.Visible = False Then
        '처방내역 생성
        blnSuccess = objMyOrder.ExecuteSqlStmt
        If blnSuccess = False Then GoTo ErrSave
        
        '채혈내역 생성
        blnSuccess = objMyCollection.ExecuteSqlStmt
        If blnSuccess = False Then GoTo ErrSave
        
        For ii = 1 To objMyCollection.ColCount
            objMyCollection.GetBarcodeLabel (ii)
        Next
    '바코드 라벨 폼피드 추가....
        Dim objBar As New clsBarcode
        
'        Set objBar.MyDB = dbconn
        Set objBar.TableInfo = New clsTables
        
        objBar.Get_PortNo
        objBar.Label_FormFeed
    End If


Skip:
    
    
'----- Commit Transaction -----

    Set objSql = New clsBBSSQLStatement
    If objSql.SetDonorStatus(donorid, accdt, DonorStatus.stsDonation, False) = False Then GoTo ErrSave
    
    '헌혈자 접수내역 canclefg 업데이트
    '취소했다가 업데이트 하는 경우가 있기에.
    '2001/09/20. 추가
    
    Dim SSQL As String
    SSQL = objSql.SetDonorAcc(donorid, accdt)
    DBConn.Execute SSQL
    
    Set objSql = Nothing
    
    DBConn.CommitTrans
    SaveDonation = True
    MsgBox "정상적으로 처리되었습니다.", vbInformation, "정보확인"
    
    
    
    Call ClassInitialize
    'Call FormInitialize
    
    Exit Function
    
ErrSave:
'----- Rollback Transaction -----
    DBConn.RollbackTrans
     MsgBox "정상적으로 처리되지 않았습니다.", vbInformation, "정보확인"
    Call ClassInitialize
    
    SaveDonation = False
    Set objSql = Nothing
    Exit Function
    
ErrOther:
    
    MsgBox "정상적으로 처리되지 않았습니다.", vbInformation, "정보확인"
    
    Call ClassInitialize

    SaveDonation = False
End Function

Private Sub TblSort()
    With tblResult
        .SortBy = SortByRow
        .SortKey(1) = 19  'DeliveryLocation
        .SortKey(2) = 7   '희망채취시간
        .SortKey(3) = 9   'WorkArea
        .SortKey(4) = 4   '검체코드
        .SortKey(5) = 10  '보관구분
        .SortKey(6) = 6   '응급여부
        .SortKey(7) = 3   '검사코드
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending
        .SortKeyOrder(3) = SortKeyOrderAscending
        .SortKeyOrder(4) = SortKeyOrderAscending
        .SortKeyOrder(5) = SortKeyOrderAscending
        .SortKeyOrder(6) = SortKeyOrderAscending
        .SortKeyOrder(7) = SortKeyOrderAscending
        .Col = 1: .COL2 = .MaxCols
        .Row = 0: .Row2 = .MaxRows
        .Action = ActionSort
    End With
End Sub

Private Function SaveOrder() As Boolean
'처방 루틴

    Dim i As Long
    Dim lngStartOrdNo As Long
    Dim strTmpPtID As String
    Dim strDonoraccdt As String
    Dim datDateTime As Date
    
    datDateTime = GetSystemDate
    'strTmpPtID = GetPtID
    '헌혈자 id에 대한 임시환자id를 넘겨준다.
    '20010206
    'strDonorAccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    strTmpPtID = lblTmpPtId.Caption ' GetPtID(strDonorAccdt, lblDonorID.Caption)
    
    If strTmpPtID = "" Or strTmpPtID = "0" Then SaveOrder = False: Exit Function
    
    With objMyOrder
        'Order Class 기본 데이타 Store
        .PtId = strTmpPtID   '번호 부여 정보에서 생성
        .orddt = Format(datDateTime, CS_DateDbFormat)
        .Bussdiv = "1"  '외래 1, 병동 2, 응급 3 단체 검진 4
        .bedindt = ""
        .DeptCd = BLOOD_DEPTCD
        .MajDoct = ""
        .wardid = ""
        .HosilID = ""
        .ROOMID = ""
        .OrdDoct = ObjMyUser.EmpId
        .ReceptNo = ""
        .EntID = ObjMyUser.EmpId
        .EntDt = Format(datDateTime, CS_DateDbFormat)
        .EntTm = Format(datDateTime, CS_TimeDbFormat)
        .donefg = "0" '처방 '0'
        .OrgAccNo = ""
        .orddiv = "L"
        Call .MoveData(tblResult)                   '클래스로 데이타 Move
        If .CreateSqlStmt(lngStartOrdNo) = False Then MsgBox "Createsqlstmt 에러": Exit Function  'Database로 저장
        
    End With
          
    '처방번호 Display
    With tblResult
        .Col = 1
        For i = 1 To .DataRowCnt
            .Row = i
            .value = Val(.value) + lngStartOrdNo
        Next
    End With
    
    SaveOrder = True
End Function
'% 발생한 처방데이타를 기준으로 채혈접수내역을 생성하기 위해
'% 모든 데이타를 Array로 Assign한다.
Private Sub ReadyToCollect()
    
    Dim i As Integer
    Dim tmpData() As String
    Dim datDateTime As Date
    
    datDateTime = GetSystemDate
    
    With objMyCollection
    
        .spcyy = Mid(Format(datDateTime, "YYYY"), 4)         '검체년도
        .PtId = objMyOrder.PtId                                    '환자ID
        .ptnm = txtDonorNm
        
        'DonorID, DonorAccDt를 넘겨준다.
        .donorid = lblDonorID.Caption
        .donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
        
        .Sex = Mid(lblSex.Caption, 1, 1)                            '성별
        
        .AgeDay = DateDiff("y", medGetP(lblSex.Caption, 2, "/"), datDateTime) '환자일령
        .bedindt = ""                                               '입원일
        .EntDt = Format(datDateTime, CS_DateDbFormat)         '입력일
        .EntTm = Format(datDateTime, CS_TimeDbFormat)         '입력시간
        .EntID = ObjMyUser.EmpId                                    '입력자
        .OrgAccNo = ""                                              '원접수번호
        .wardid = ""                                                '병동ID
        .HosilID = ""                                               '병실ID
        .ROOMID = ""                                                '병실ID
        .BedID = ""                                                 '침상ID
        .coldt = Format(datDateTime, CS_DateDbFormat)         '채혈일
        .colid = ObjMyUser.EmpId                                    '채혈자
        .orgbuildcd = ObjSysInfo.BuildingCd                         '** 채혈이 수행되는 건물코드
    End With
        
    With tblResult
        ReDim tmpData(0 To 17)
        
        For i = 1 To .DataRowCnt
           .Row = i
           .Col = 19:  tmpData(0) = .value                          'Delivery Location
           .Col = 12:  tmpData(1) = .value                          'TestDiv
           .Col = 9:   tmpData(2) = .value                          'WorkArea
           .Col = 4:   tmpData(3) = .value                          'SpcCd
           .Col = 10:  tmpData(4) = .value                          'StoreCd
           .Col = 6:   tmpData(5) = CStr(Val(.value))               'StatFg
           .Col = 7:   tmpData(6) = .value                          'ReqColDate
           
           .Col = 13:  tmpData(7) = .value                          'MultiFg
           .Col = 14:  tmpData(8) = .value                          'SpcGrp
           tmpData(9) = Format(datDateTime, CS_DateDbFormat)        '처방일을 희망채혈일로.. 2000/04/03 by 정미경
           .Col = 1:   tmpData(10) = .value                         'OrdNo
           .Col = 15:  tmpData(11) = .value                         'OrdSeq
           .Col = 3:   tmpData(12) = .value                         'OrdCd
           tmpData(13) = ObjMyUser.DeptCd                           '진료과
           tmpData(14) = ObjMyUser.EmpId                            '처방의
           tmpData(15) = ""                                         '주치의
           .Col = 16:  tmpData(16) = .value                         '약어명
           .Col = 17:  tmpData(17) = .value                         '라벨출력장수
           Call objMyCollection.AddLabCollect(tmpData(0), tmpData(1), tmpData(2), tmpData(3), tmpData(4), _
                                      tmpData(5), tmpData(6), tmpData(7), tmpData(8), tmpData(9), tmpData(10), _
                                      tmpData(11), tmpData(12), tmpData(13), tmpData(14), tmpData(15), tmpData(16), tmpData(17))
        Next
    End With

End Sub

Private Sub ClassInitialize()
    Dim datDateTime  As Date
    
    datDateTime = GetSystemDate
    
    Set objMySQL = Nothing
    Set objMySQL = New clsBBSSQLStatement
'    objMySQL.setDbConn DBConn
    
    Set objMyOrder = Nothing
    Set objMyOrder = New clsDonorBusiOrder
    With objMyOrder
        .DateTime = datDateTime
        .BuildingNo = ObjSysInfo.BuildingNo
'        .setDbConn DBConn
    End With
    
    Set objMyCollection = Nothing
    Set objMyCollection = New clsDonorTestCollection
    
    With objMyCollection
        .DateTime = datDateTime
'        .setDbConn DBConn
        Set .SortList = frmControls.lstList
        Call .InitRtn
    End With
End Sub
'2001-11-27추가
Public Sub CallDonorNmLostFocus()
    Call txtDonorNm_LostFocus
End Sub



