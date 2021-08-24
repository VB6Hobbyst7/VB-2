VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS411 
   BackColor       =   &H00DBE6E6&
   Caption         =   "검사의뢰"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13920
   Icon            =   "frmBBS411.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   13920
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdCallBlood 
      BackColor       =   &H00F4F0F2&
      Caption         =   "헌혈등록(&N)"
      Height          =   510
      Left            =   2250
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdPhersis 
      BackColor       =   &H00C8CEDF&
      Caption         =   "Phersis등록(&N)"
      Height          =   510
      Left            =   3585
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   10875
      Style           =   1  '그래픽
      TabIndex        =   6
      Tag             =   "128"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   9555
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "124"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   8235
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "검사취소"
      Height          =   510
      Left            =   6915
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "15101"
      Top             =   7575
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSComctlLib.TabStrip tabAccDt 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
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
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
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
      TabIndex        =   10
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
      TabIndex        =   8
      Top             =   2310
      Width           =   9930
      Begin VB.TextBox txtReservedID 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00CFDCDE&
         Height          =   330
         Left            =   4245
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   225
         Width           =   1125
      End
      Begin VB.ComboBox cboDonorCd 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "frmBBS411.frx":076A
         Left            =   1035
         List            =   "frmBBS411.frx":077A
         Locked          =   -1  'True
         Style           =   1  '단순 콤보
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   225
         Width           =   2055
      End
      Begin MedControls1.LisLabel lblReservedNm 
         Height          =   330
         Left            =   5370
         TabIndex        =   35
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
         Left            =   45
         TabIndex        =   36
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
         Left            =   3255
         TabIndex        =   37
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   2280
      TabIndex        =   11
      Top             =   2910
      Width           =   9930
      Begin MedControls1.LisLabel lblStsNm 
         Height          =   315
         Left            =   1050
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   2295
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   3585
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   4530
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   5850
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   6795
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   8115
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   9075
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   45
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   2595
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   4845
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   7110
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   195
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2280
      TabIndex        =   16
      Top             =   720
      Width           =   9930
      Begin VB.TextBox txtDonorNm 
         Appearance      =   0  '평면
         Height          =   330
         Left            =   1035
         TabIndex        =   0
         Top             =   180
         Width           =   1515
      End
      Begin MedControls1.LisLabel lblDOB 
         Height          =   315
         Left            =   4260
         TabIndex        =   17
         Top             =   180
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "2001-01-01"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSex 
         Height          =   330
         Left            =   6615
         TabIndex        =   18
         Top             =   180
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
         Left            =   8925
         TabIndex        =   19
         Top             =   180
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
         Left            =   4260
         TabIndex        =   20
         Top             =   540
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
         Left            =   6615
         TabIndex        =   21
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
         Left            =   1035
         TabIndex        =   22
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
         Left            =   1815
         TabIndex        =   23
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   0
         Left            =   45
         TabIndex        =   27
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
         Caption         =   "성   명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   3270
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   3270
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   540
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
         Left            =   5625
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   5625
         TabIndex        =   31
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   7935
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   7530
         TabIndex        =   24
         Top             =   690
         Width           =   210
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   2280
      TabIndex        =   25
      Top             =   3525
      Width           =   9915
      _ExtentX        =   17489
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
   Begin VB.Frame fraTest 
      BackColor       =   &H00DBE6E6&
      Height          =   3720
      Left            =   2280
      TabIndex        =   12
      Top             =   3765
      Width           =   9930
      Begin MedControls1.LisLabel lblTestChk 
         Height          =   345
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   609
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
      Begin FPSpread.vaSpread tblResult 
         Height          =   3150
         Left            =   30
         TabIndex        =   13
         Tag             =   "10114"
         Top             =   495
         Width           =   9765
         _Version        =   196608
         _ExtentX        =   17224
         _ExtentY        =   5556
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
         MaxCols         =   24
         MaxRows         =   11
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS411.frx":07A8
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   11
      End
      Begin MedControls1.LisLabel lblTmpPtId 
         Height          =   315
         Left            =   8190
         TabIndex        =   14
         Top             =   150
         Width           =   1605
         _ExtentX        =   2831
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MSComctlLib.TabStrip tabGroup 
         Height          =   345
         Left            =   30
         TabIndex        =   15
         Top             =   135
         Width           =   7125
         _ExtentX        =   12568
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
         Height          =   315
         Index           =   12
         Left            =   7200
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   150
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
         Caption         =   "임시 ID"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmBBS411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Enum TblColumn
    tcSEL = 1
    TcName
    tcCODE
    tcQTY
End Enum
Private objMySQL As New clsBBSSQLStatement
Private objMyOrder As New clsDonorBusiOrder
Private objMyCollection As New clsDonorTestCollection
Private objCollect As New clsLISCollectioin

'2001-11-27추가
Private strSaveDonorId As String
Private strSaveDonorNm As String


Private Sub FrameInitialize()
    tabAccDt.Tabs.Clear
    tabAccDt.Visible = False
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

End Sub
'2001-11-27추가
Private Sub cmdCallBlood_Click()
    frmBBS404.Show
    frmBBS404.txtDonorNm.Text = strSaveDonorNm
    Call frmBBS404.CallDonorNmLostFocus

End Sub

Private Sub cmdCancel_Click()
    Dim donorid As String
    Dim donoraccdt As String
    Dim tmpptid As String
    
    If tabAccDt.SelectedItem Is Nothing Then Exit Sub
    
    donorid = lblDonorID.Caption
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    tmpptid = lblTmpPtId.Caption
    
    If objMySQL.SetDonorScreenCancel(donorid, donoraccdt, tmpptid) = True Then
        Call FrameInitialize
    End If
End Sub

Private Sub cmdClear_Click()
    Call FrameInitialize
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
'2001-11-27추가
Private Sub cmdPhersis_Click()
    
    frmBBS412.Show
    frmBBS412.txtDonorNm.Text = strSaveDonorNm
    Call frmBBS412.CallDonorNmLostFocus

End Sub
Private Sub cmdSave_Click()
    
    If Not TEST_FOR_PHERSIS And cboDonorCd.ListIndex = 3 Then
        MsgBox "성분헌혈은 검사의뢰하실수 없습니다." & vbCrLf & "Pheresis등록화면을 사용하십시오", vbInformation + vbOKOnly
        Exit Sub
    End If
        
    If tblResult.DataRowCnt = 0 Then
        MsgBox "검사의뢰할 항목이 없습니다.", vbInformation, "정보확인"
    Else
        If Save = True Then
        End If
        Call ClassInitialize
    End If
End Sub

Private Function Save() As Boolean
    Dim strOrdDt    As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim lngAccSeq   As Integer
    Dim blnSuccess  As Boolean
    Dim objSQL      As clsBBSSQLStatement

    Dim donorid     As String
    Dim accdt       As String
    Dim SSQL        As String
    Dim ii          As Integer
    
    donorid = lblDonorID.Caption
    If donorid = "" Then
        MsgBox "Donor를 선택한후 진행하세요.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    accdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    Call TblSort
    
On Error GoTo ErrOther
    
    '처방 루틴
    If SaveOrder = False Then GoTo ErrOther
    
    'objCollect.InitRtn
    
    Call ReadyToCollect              '채혈준비
    
    '원래 막혔슴(2001/09/20)
    If objMyCollection.DoCollection = False Then GoTo ErrOther    '채혈수행

 '----- Begin Transaction -----
    DBConn.BeginTrans
   
On Error GoTo ErrSave

    '처방내역 생성
    blnSuccess = objMyOrder.ExecuteSqlStmt
    If blnSuccess = False Then GoTo ErrSave
    '원래 막혔슴(2001/09/20)
    
    '채혈내역 생성
    blnSuccess = objMyCollection.ExecuteSqlStmt
    If blnSuccess = False Then GoTo ErrSave
'==============2001/09/20===================
    '채혈 루틴
'    objCollect.SetTrans = False
'    If objCollect.DoCollection = False Then GoTo ErrSave
'    For ii = 1 To objCollect.ColCount
'        Call objCollect.GetLabNumbers(ii, strWorkArea, strAccdt, lngAccSeq)
'        sSql = objMySQL.SetDonorAccHistoryUpdateByTmpID2(donorid, accdt, lblTmpPtId.Caption)
'        DBConn.Execute sSql
'        sSql = objMySQL.SetTestRequest(donorid, accdt, _
'                            Format(GetSystemDate, PRESENTDATE_FORMAT), ii, strWorkArea, strAccdt, lngAccSeq)
'        DBConn.Execute sSql
'    Next
'============================================
    For ii = 1 To objMyCollection.ColCount
        objMyCollection.GetBarcodeLabel (ii)
    Next

'----- Commit Transaction -----

    Set objSQL = New clsBBSSQLStatement
    If objSQL.SetDonorStatus(donorid, accdt, DonorStatus.stsDonation, False) = False Then GoTo ErrSave
    
    SSQL = objSQL.SetDonorAcc(donorid, accdt)
    DBConn.Execute SSQL
    
    Set objSQL = Nothing
    
    DBConn.CommitTrans
    Save = True
    MsgBox "정상적으로 처리되었습니다.", vbInformation, "정보확인"
    
    Call FrameInitialize
    Exit Function
    
ErrSave:
'----- Rollback Transaction -----
    DBConn.RollbackTrans
    Save = False
    MsgBox Err.Description, vbExclamation
    Exit Function
    
ErrOther:
    
    MsgBox "정상적으로 처리되지 않았습니다.", vbInformation, "정보확인"

    Save = False

End Function

Private Function SaveOrder() As Boolean
    Dim i As Long
    Dim lngStartOrdNo As Long
    Dim strTmpPtID As String
    Dim strDonorAccdt As String
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

        .spcyy = LIS_BarDiv & Mid(Format(datDateTime, "YYYY"), 4)         '검체년도
       
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
        .OrgBuildCd = ObjSysInfo.BuildingCd                         '** 채혈이 수행되는 건물코드
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
    
    Call FrameInitialize
    Call ClassInitialize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objMySQL = Nothing
    Set objMyOrder = Nothing
    Set objMyCollection = Nothing
    Set objCollect = Nothing
End Sub

Private Sub tabAccDt_Click()
    
    Dim donorid As String
    Dim canEdit As Boolean
    
    donorid = lblDonorID.Caption
    Call tabAccdtClickCode(donorid, Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT))
    Call SetDonorStatus(donorid, Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT))
    
'    canEdit = GetCanEdit
'    fraDonation.Enabled = canEdit

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



Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ii As Integer
    If lblTestChk.Visible = True Then Exit Sub
    
    If Row = 0 And Col = 6 Then
        With tblResult
            
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = 6
                If .CellType = CellTypeCheckBox Then .value = IIf(.value = 0, 1, 0)
            Next
        End With
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

            lblDOB.Caption = .mDOB
            lblSex.Caption = .mSEX
            lblABO.Caption = .mABO
            lblCnt.Caption = .Mcnt
            lblTotVol.Caption = .mTotVol
        
            Call ShowAccList
'            cmdNew.Enabled = True
        End If
    End With
    Set objDonor = Nothing
End Sub

Private Sub ShowAccList()
    Dim strAccDt As String
    Dim Rs As Recordset
    '헌혈자에 대해서 접수된 정보가 있을 경우에 접수 내역을 보여준다.

'    objMySQL.setDbConn DBConn
    'Set Rs = objMySQL.GetDonorAccHistory(Trim(lblDonorID.Caption))
    
    
    '성분헌혈을 제외한 헌혈만 검사의뢰할수 있게 조회(2001/10/04, 울산 동강병원)
    If TEST_FOR_PHERSIS Then
        Set Rs = objMySQL.GetDonorAccHistory(Trim(lblDonorID.Caption))
    Else
        Set Rs = objMySQL.GetDonorAccdtHistoryDivPheresis(Trim(lblDonorID.Caption))
    End If
    
    If Rs.EOF Then
        MsgBox "검사의뢰대상이 없습니다.", vbInformation + vbOKOnly, "헌혈자검사의뢰"
        
        tabAccDt.Tabs.Clear
        tabAccDt.Visible = False
        
    Else
        tabAccDt.Tabs.Clear
        tabAccDt.Visible = True
        
        Do Until Rs.EOF
            strAccDt = Format(Rs.Fields("donoraccdt").value & "", "####-##-##")
            tabAccDt.Tabs.Add , , strAccDt
            Rs.MoveNext
        Loop
        
        cmdSave.Enabled = True
        Call tabAccDt_Click
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
        lblTmpPtId.Caption = RsTestReq.Fields("tmpid").value & ""
        txtReservedID = RsTestReq.Fields("reservedid").value & ""
        lblReservedNm.Caption = objMySQL.GetPtntNm(txtReservedID)
        
        '검사의뢰내역을 읽는다-----------------------------------------
        Set RsDonorTest = objMySQL.Get_TestHistory(donorid, donoraccdt)
        If RsDonorTest.EOF Then
'            'dbconn.DisplayErrors
            Exit Sub
        End If
        
        
        If RsDonorTest.RecordCount > 0 Then
            '검사의뢰내역을 조회하여 보여준다.
            '이미 검사의뢰가 진행된 상태의 헌혈자정보
            
'            If RsTestReq.Fields("donationdt") <> "" Then
                cmdSave.Enabled = False
                cmdCancel.Enabled = True
'            Else
'                cmdSave.Enabled = True
'                cmdCancel.Enabled = False
'            End If
                        
            Set QueryTest = objMySQL.GetDonorTestDt(donorid, donoraccdt)
            Dim strTmpID As String
            lblTmpPtId.Caption = RsTestReq.Fields("tmpid").value & ""
            strTmpID = QueryTest.Fields("tmpid").value & ""
            
            'h7lab102에서 검사의뢰 정보를 불러온다.
            lblTestChk.Visible = True
            Call QueryInformation(strTmpID)
            Set QueryTest = Nothing
        Else
            lblTestChk.Visible = False
        End If
        
        
        
''''''''''        If RsDonorTest.RecordCount < 1 Then
''''''''''            '검사의뢰가 되지 않은 환자에 대해서는 검사항목 마스터에등록된 검사항목을
''''''''''            '무조건 보여준다.
''''''''''            Set NewTest = objMySQL.GetTestSpc
''''''''''            If Not NewTest.EOF Then
''''''''''                Dim objdic As New clsDictionary
''''''''''                Dim lngseq As Long
''''''''''
''''''''''                objdic.Clear
''''''''''                objdic.FieldInialize "seq", "testcd,spccd"
''''''''''                Do Until NewTest.EOF
''''''''''                    lngseq = lngseq + 1
''''''''''                    objdic.AddNew lngseq, Join(Array(NewTest.Fields("cdval1").value, NewTest.Fields("field1").value), COL_DIV)
''''''''''                    NewTest.MoveNext
''''''''''                Loop
''''''''''                lblTestChk.Visible = False
''''''''''                Call Default_Test(objdic)
''''''''''                Set NewTest = Nothing
''''''''''                Set objdic = Nothing
''''''''''                cmdSave.Enabled = True
''''''''''                cmdCancel.Enabled = False
''''''''''            End If
''''''''''        Else
''''''''''            '검사의뢰내역을 조회하여 보여준다.
''''''''''            '이미 검사의뢰가 진행된 상태의 헌혈자정보
''''''''''
'''''''''''            If RsTestReq.Fields("donationdt") <> "" Then
''''''''''                cmdSave.Enabled = False
''''''''''                cmdCancel.Enabled = True
'''''''''''            Else
'''''''''''                cmdSave.Enabled = True
'''''''''''                cmdCancel.Enabled = False
'''''''''''            End If
''''''''''
''''''''''            Set QueryTest = objMySQL.GetDonorTestDt(donorid, donoraccdt)
''''''''''            Dim strTmpID As String
''''''''''            lblTmpPtId.Caption = RsTestReq.Fields("tmpid").value
''''''''''            strTmpID = QueryTest.Fields("tmpid").value
''''''''''
''''''''''            'h7lab102에서 검사의뢰 정보를 불러온다.
''''''''''            lblTestChk.Visible = True
''''''''''            Call QueryInformation(strTmpID)
''''''''''            Set QueryTest = Nothing
''''''''''
''''''''''        End If
        
        
        
        
        
        Set RsDonorTest = Nothing
        Set RsTestReq = Nothing
    End If
    
End Sub

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

Private Sub Default_Test(objDefault As clsDictionary)
    Dim objQueryTest As New clsBBSSQLStatement
    Dim objGDic As New clsDictionary
    Dim DefaultTest As Recordset
    Dim strTmp As String
    Dim lngseq As Long
    
'    objQueryTest.setDbConn DBConn
'SpcNm, " & _
                        "        d.field2 as LabDiv, e.field2 as LabRange, '1' InsurFg " & _
    objGDic.Clear
    objGDic.FieldInialize "seq", "ordno1,testnm,testcd,spccd,gubyu,stat,reqdt,statfg,workarea," & _
                          "storecd,rndfg,testdiv,multifg,spcgrp,ordseq1,abbrnm5,labelcnt,statflag,location,testlocation,spcnm,labdiv,labrange,insurfg"
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
                                          .Fields("labelcnt").value & "", .Fields("statflags").value & "", "location", "중앙", _
                                          .Fields("spcnm").value & "", .Fields("labdiv").value & "", .Fields("labrange").value & "", .Fields("insurfg").value & ""), COL_DIV)
            End If
        End With
        objDefault.MoveNext
    Loop
    
    '화면에 보여주자......
    Call TblResult_Display(objGDic)
    Set objGDic = Nothing
    Set objQueryTest = Nothing
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
            .Col = 6: .CellType = CellTypeCheckBox: .TypeCheckCenter = True
            If ObjDic.Fields("statfg") = "1" Then
                .value = 1
            Else
                .value = 0
            End If
            

'            If objdic.Fields("statfg") = "1" Then
'                .Col = 6: .CellType = CellTypeCheckBox
'                   .TypeCheckCenter = True
'            Else
'                .Col = 6: .CellType = CellTypeStaticText
'            End If
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
            '
            .Col = 21: .value = ObjDic.Fields("spcnm")
            .Col = 22: .value = ObjDic.Fields("labdiv")
            .Col = 23: .value = ObjDic.Fields("labrange")
            .Col = 24: .value = ObjDic.Fields("insurfg")
            
            tmpStatFg = medGetP(ObjDic.Fields("statflag"), 1, ";")  '건물별 응급가능 여부
            tmpTestFg = medGetP(ObjDic.Fields("statflag"), 2, ";")  '건물별 검사가능 여부
'    '***건물정보 사용
'        If ObjSysInfo.UseBuildingInfo = "1" Then
'
'            .Col = enORDSHEET.tcSTATCHK
'            If .value = 1 Then   '응급선택
'                .Col = enORDSHEET.tcSTATFG
'                If .value = "1" Then
'
'                    ' ** 중앙/안이검사실에서 응급검사가 발생하면 --> 응급센터로...
'                    If ObjSysInfo.BuildingCd = CentralLab Or ObjSysInfo.BuildingCd = AneLab Then
'                        .Col = enORDSHEET.tcBUILDCD: .value = EmergencyLab
'                        .Col = enORDSHEET.tcBUILDNM: .value = EmergencyLabNm
'
'                    ' ** 해당건물에서 응급검사 가능함
'                    Else
'                        .Col = enORDSHEET.tcBUILDCD: .value = ObjSysInfo.BuildingCd
'                        .Col = enORDSHEET.tcBUILDNM: .value = ObjSysInfo.BuildingNm
'                    End If
'                    Exit Sub
'
'                Else
'                    ' ** 해당건물에서 응급검사 불가능...
'                    .Col = enORDSHEET.tcSTATCHK
'                    .CellType = CellTypeStaticText
'                    .Text = ""
'                End If
'            End If
'
'            '** 일반검사 가능여부
'            .Col = enORDSHEET.tcTESTFLAG
'
'            ' ** 해당건물에서 일반검사 가능함
'            If .value = "1" Then
'                .Col = enORDSHEET.tcBUILDCD: .value = ObjSysInfo.BuildingCd
'                .Col = enORDSHEET.tcBUILDNM: .value = ObjSysInfo.BuildingNm
'
'            ' ** 해당건물에서 일반검사 불가능함 --> 중앙검사실로...
'            Else
'                .Col = enORDSHEET.tcBUILDCD: .value = CentralLab
'                .Col = enORDSHEET.tcBUILDNM: .value = CentralLabNm
'            End If
'
'    '***건물정보 사용하지 않음
'        Else
'            .Col = enORDSHEET.tcBUILDCD: .value = ObjSysInfo.BuildingCd
'            .Col = enORDSHEET.tcBUILDNM: .value = ObjSysInfo.BuildingNm
'        End If
             If ObjSysInfo.UseBuildingInfo = "1" Then
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
            Else
                .Col = 19: .value = ObjSysInfo.BuildingCd
                .Col = 20: .value = ObjSysInfo.BuildingNm
            End If
        
            ObjDic.MoveNext
        Loop
    End With
    
            
End Sub

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
    
'    objMySQL.setDbConn DBConn
    
    objDicT.Clear
    objDicT.FieldInialize "ptid,orddt,ordno,ordseq", "ordcd,spccd,reqdate,statfg"
    
    
    Set RsDonorTest = objMySQL.GetDonnorTest(tmpid)
    
    If Not RsDonorTest.EOF Then
        Do Until RsDonorTest.EOF
            objDicT.AddNew Join(Array(RsDonorTest.Fields("ptid").value & "", RsDonorTest.Fields("orddt").value & "", RsDonorTest.Fields("ordno").value & "", _
                                RsDonorTest.Fields("ordseq").value & ""), COL_DIV), Join(Array(RsDonorTest.Fields("ordcd").value & "", _
                                RsDonorTest.Fields("spccd").value & "", RsDonorTest.Fields("reqdt").value & "" & RsDonorTest.Fields("reqtm").value & "", RsDonorTest.Fields("statfg").value & ""), COL_DIV)
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
            If Not RsDisplay.EOF Then
                With RsDisplay
                    objDicD.AddNew Join(Array(objDicT.Fields("ptid"), objDicT.Fields("orddt"), objDicT.Fields("ordno"), objDicT.Fields("ordseq")), COL_DIV), _
                                   Join(Array(objDicT.Fields("ordno"), .Fields("testnm").value & "", .Fields("testcd").value & "", .Fields("spccd").value & "", _
                                              "1", "", objDicT.Fields("reqdate"), objDicT.Fields("statfg"), .Fields("workarea").value & "", _
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



