VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm301QCMaster_N 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   14955
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   Begin MedControls1.LisLabel LisLabel6 
      Height          =   300
      Left            =   75
      TabIndex        =   27
      Top             =   45
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   529
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
      Caption         =   "◈ 컨트롤 정보"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel8 
      Height          =   300
      Left            =   10215
      TabIndex        =   29
      Top             =   45
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   529
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
      Caption         =   "◈ 부가 기능"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "저장(&S)"
      Height          =   510
      Left            =   9180
      Style           =   1  '그래픽
      TabIndex        =   42
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "삭제(&D)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   41
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   40
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   39
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1380
      Left            =   75
      TabIndex        =   0
      Top             =   285
      Width           =   10125
      Begin VB.TextBox txtCtrlCd 
         Height          =   375
         Left            =   1410
         MaxLength       =   9
         TabIndex        =   2
         Text            =   "하둘셋넷다여일여아"
         Top             =   135
         Width           =   2145
      End
      Begin VB.CommandButton cmdPopCtrl 
         BackColor       =   &H00F4F0F2&
         Height          =   360
         Left            =   3555
         Picture         =   "frm301QCMaster_N.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   135
         Width           =   330
      End
      Begin MedControls1.LisLabel lblCtrlNm 
         Height          =   330
         Left            =   3900
         TabIndex        =   3
         Top             =   150
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   582
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
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCtrlDiv 
         Height          =   360
         Left            =   5175
         TabIndex        =   8
         Top             =   540
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "내부정도관리"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblEqp 
         Height          =   360
         Left            =   7695
         TabIndex        =   9
         Top             =   540
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C001 Coulter Stks"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBuilding 
         Height          =   360
         Left            =   1410
         TabIndex        =   10
         Top             =   945
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "10 본원"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSection 
         Height          =   360
         Left            =   5175
         TabIndex        =   11
         Top             =   945
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "HE Hematology"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWorkarea 
         Height          =   360
         Left            =   7695
         TabIndex        =   12
         Top             =   945
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "03 Hematology"
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   465
         Left            =   1410
         TabIndex        =   4
         Top             =   450
         Width           =   2490
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "High"
            Height          =   180
            Index           =   2
            Left            =   1740
            TabIndex        =   7
            Top             =   150
            Width           =   705
         End
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Normal"
            Height          =   180
            Index           =   1
            Left            =   765
            TabIndex        =   6
            Top             =   150
            Width           =   960
         End
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Low"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   150
            Value           =   -1  'True
            Width           =   705
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   30
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "Control 정보"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   30
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   945
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "건물구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   30
         TabIndex        =   45
         Top             =   540
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "Level 구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   3900
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   540
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "정도관리구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   3900
         TabIndex        =   47
         Top             =   945
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "섹션구분"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   6675
         TabIndex        =   48
         Top             =   540
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "검사장비"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   6675
         TabIndex        =   49
         Top             =   945
         Width           =   1005
         _ExtentX        =   1773
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
         Caption         =   "Workarea"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   75
      TabIndex        =   13
      Top             =   1605
      Width           =   10125
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   8
         Left            =   3915
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   135
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "시작일"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   3915
         TabIndex        =   53
         Top             =   540
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "비  고"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   30
         TabIndex        =   50
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "Lot No."
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   30
         TabIndex        =   51
         Top             =   540
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "제조사"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdPopLotNo 
         BackColor       =   &H00F4F0F2&
         Height          =   360
         Left            =   3555
         Picture         =   "frm301QCMaster_N.frx":00B2
         Style           =   1  '그래픽
         TabIndex        =   37
         Top             =   120
         Width           =   345
      End
      Begin VB.TextBox txtLotNo 
         Height          =   360
         Left            =   1410
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   135
         Width           =   2160
      End
      Begin VB.TextBox txtRemark 
         Height          =   375
         Left            =   5190
         MaxLength       =   50
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   540
         Width           =   4200
      End
      Begin VB.TextBox txtMakeCd 
         Height          =   375
         Left            =   1410
         MaxLength       =   20
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   540
         Width           =   2490
      End
      Begin MSComCtl2.DTPicker dtpOpenDt 
         Height          =   375
         Left            =   5190
         TabIndex        =   14
         Top             =   135
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   661
         _Version        =   393216
         Format          =   167313409
         CurrentDate     =   37917
      End
      Begin MSComCtl2.DTPicker dtpExpDt 
         Height          =   375
         Left            =   7935
         TabIndex        =   15
         Top             =   135
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   661
         _Version        =   393216
         Format          =   167313409
         CurrentDate     =   37917
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   10
         Left            =   6660
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   135
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "만료일"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel7 
      Height          =   300
      Left            =   11895
      TabIndex        =   28
      Top             =   45
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   529
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
      Caption         =   "◈ Multi-Rule 설정"
      Appearance      =   0
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   11880
      TabIndex        =   18
      Top             =   255
      Width           =   2580
      Begin VB.CheckBox chkRule 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Step 1 : Once 3 SD"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   285
         Width           =   2100
      End
      Begin VB.CheckBox chkRule 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Step 2 : Once 4 SD"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   25
         Top             =   525
         Width           =   2100
      End
      Begin VB.CheckBox chkRule 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Step 3 : Twice 2 SD"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   24
         Top             =   765
         Width           =   2100
      End
      Begin VB.CheckBox chkRule 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Step 4 : 4 - 1 SD"
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   23
         Top             =   1005
         Width           =   2400
      End
      Begin VB.CheckBox chkRule 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Step 5 : 10 Times Trend"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   22
         Top             =   1230
         Width           =   2355
      End
      Begin VB.CommandButton cmdApply 
         BackColor       =   &H00F4F0F2&
         Caption         =   "적용"
         Height          =   330
         Left            =   1755
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   1755
         Width           =   780
      End
      Begin VB.CommandButton cmdDeSel 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전체해제"
         Height          =   330
         Left            =   915
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   1755
         Width           =   840
      End
      Begin VB.CommandButton cmdSelAll 
         BackColor       =   &H00F4F0F2&
         Caption         =   "전체선택"
         Height          =   330
         Left            =   75
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   1755
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   135
         X2              =   2490
         Y1              =   1590
         Y2              =   1590
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   2295
      Left            =   10200
      TabIndex        =   30
      Top             =   285
      Width           =   1680
      Begin VB.CommandButton cmdSchedule 
         BackColor       =   &H00E0E0E0&
         Caption         =   "QC 스케쥴"
         Height          =   495
         Left            =   195
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   1200
         Width           =   1320
      End
      Begin VB.CommandButton cmdPrevData 
         BackColor       =   &H00E0E0E0&
         Caption         =   "이전 데이터"
         Enabled         =   0   'False
         Height          =   495
         Left            =   195
         Style           =   1  '그래픽
         TabIndex        =   31
         Top             =   585
         Width           =   1320
      End
      Begin VB.CommandButton cmdCalculation 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Calculation"
         Height          =   495
         Left            =   195
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   1560
         Visible         =   0   'False
         Width           =   1320
      End
   End
   Begin FPSpread.vaSpread tblQcMst 
      Height          =   5550
      Left            =   75
      TabIndex        =   34
      Tag             =   "30111"
      Top             =   2895
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   9790
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
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
      MaxCols         =   15
      MaxRows         =   18
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      SpreadDesigner  =   "frm301QCMaster_N.frx":0164
   End
   Begin MedControls1.LisLabel LisLabel9 
      Height          =   300
      Left            =   75
      TabIndex        =   35
      Top             =   2580
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   529
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
      Caption         =   "◈ 컨트롤 검사 정보"
      Appearance      =   0
   End
   Begin VB.Label lblLotNo 
      Height          =   495
      Left            =   540
      TabIndex        =   38
      Top             =   8355
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frm301QCMaster_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Coding By Legends

Public Event LastFormUnload()

Private mvarParentHwnd As Long

Public Property Let ParentHwnd(ByVal vData As Long)
    mvarParentHwnd = vData
End Property

Public Property Get ParentHwnd() As Long
    ParentHwnd = mvarParentHwnd
End Property

Private Sub cmdApply_Click()
    Dim i As Long
    Dim j As Long
    
    With tblQcMst
        .ReDraw = False
        For i = chkRule.LBound To chkRule.UBound
            .Row = 1: .Row2 = .DataRowCnt
            .Col = i + 11: .Col2 = i + 15
            .BlockMode = True
            .Value = chkRule(i).Value
            .BlockMode = False
        Next
        
        For i = 1 To .DataRowCnt
            .Row = i
            For j = 11 To 15
                .Col = j
                If .CellType <> CellTypeCheckBox Then
                    .Value = ""
                End If
            Next j
        Next i
        .ReDraw = True
    End With
End Sub

Private Sub cmdCalculation_Click()
'    If CheckValidation = False Then Exit Sub
'
'    Call LoadForm(frm330Calculation_N, Me)
'    Call frm330Calculation_N.CallByExternal(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")))
End Sub

Private Sub cmdClear_Click()
    txtCtrlCd.Text = ""
    Call InitControl
    txtLotNo.Text = ""
    Call InitLotNo
    Call medClearTable(tblQcMst)
    tblQcMst.MaxRows = 18
End Sub

Private Sub cmdDelete_Click()
'023지우고
'024지운다
    Dim objSQL As clsLISSqlQc
    Dim strMsg As VbMsgBoxResult
    
    If CheckValidation = False Then Exit Sub
    
    strMsg = MsgBox("현재 작성된 자료를  삭제합니다." & vbNewLine & _
                    "이 LotNo가 삭제되면 이 LotNo로 수행했던 모든 작업이 삭제됩니다." & vbNewLine & vbNewLine & _
                    "계속 진행하시겠습니까?", vbExclamation + vbYesNo)
    
    If strMsg = vbNo Then Exit Sub
    
    Set objSQL = New clsLISSqlQc
    
    On Error GoTo ErrTrap
    DBConn.BeginTrans
    DBConn.Execute objSQL.SqlMstDeleteAll(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), Trim(txtLotNo.Text), "1")
    DBConn.Execute objSQL.SqlMstDeleteAll(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), Trim(txtLotNo.Text), "2")
    DBConn.CommitTrans
    Set objSQL = Nothing
    
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    txtLotNo.Text = ""
    Call InitLotNo
    Call medClearTable(tblQcMst)
    
    Exit Sub
ErrTrap:
    Set objSQL = Nothing
    DBConn.RollbackTrans
    MsgBox "처리도중 오류가 발생하였습니다." & vbNewLine & _
           Err.Description, vbCritical
End Sub

Private Sub cmdDeSel_Click()
    Dim i  As Long
    
    For i = chkRule.LBound To chkRule.UBound
        chkRule(i).Value = 0
    Next
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
'    If IsLastForm Then Call UnloadForm(Me)
End Sub

Private Sub cmdPopCtrl_Click()
    If lblCtrlNm.Caption <> "" Then
        DoEvents
        txtCtrlCd.Text = ""
        Call InitControl
        txtLotNo.Text = ""
        Call InitLotNo
        Call medClearTable(tblQcMst)
        tblQcMst.MaxRows = 18
    End If
    
    DoEvents
    Call LoadControlInfo
    DoEvents
    Call LoadLotNo
    DoEvents
    Call LoadControlMaster
End Sub

Private Sub LoadControlInfo(Optional ByVal pCtrlCd As String = "")
'컨트롤의 일반 정보를 불러온다..
    Dim objPop As clsPopUpList
    Dim i As Long
    
    Set objPop = New clsPopUpList

    With objPop
        .Recordset = GetControlInfo(pCtrlCd)
        .FormCaption = "컨트롤 찾기"
        .Delimiter = COL_DIV
        .FormWidth = 4470
        .ColumnHeaderText = "코드컨트롤명Level구분장비코드장비명건물코드건물명섹션코드섹션명워크애랴코드워크애랴명"
        .ColumnHeaderWidth = "854.92922475.213629.8583000000000"
        .ColumnHeaderAlign = "002"
        '0 왼쪽, 1 오른쪽, 2 가운데
        
        Call .LoadPopUp
        
        DoEvents
        
        txtCtrlCd.Text = medGetP(.SelectedString, 1, .Delimiter)
        lblCtrlNm.Caption = medGetP(.SelectedString, 2, .Delimiter)
        
        If medGetP(.SelectedString, 3, .Delimiter) = "L" Then
            optLevelCd(0).Value = True
        ElseIf medGetP(.SelectedString, 3, .Delimiter) = "N" Then
            optLevelCd(1).Value = True
        ElseIf medGetP(.SelectedString, 3, .Delimiter) = "H" Then
            optLevelCd(2).Value = True
        End If
        
        lblCtrlDiv.Caption = IIf(medGetP(.SelectedString, 4, .Delimiter) = "I", "내부정도관리", "외부정도관리")
        lblEqp.Caption = Format(medGetP(.SelectedString, 5, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 6, .Delimiter)
        lblEqp.ToolTipText = Format(medGetP(.SelectedString, 5, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 6, .Delimiter)
        lblBuilding.Caption = Format(medGetP(.SelectedString, 7, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 8, .Delimiter)
        lblBuilding.ToolTipText = Format(medGetP(.SelectedString, 7, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 8, .Delimiter)
        lblSection.Caption = Format(medGetP(.SelectedString, 9, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 10, .Delimiter)
        lblSection.ToolTipText = Format(medGetP(.SelectedString, 9, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 10, .Delimiter)
        lblWorkarea.Caption = Format(medGetP(.SelectedString, 11, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 12, .Delimiter)
        lblWorkarea.ToolTipText = Format(medGetP(.SelectedString, 11, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 12, .Delimiter)
        
    End With
    
    Set objPop = Nothing
End Sub

Private Function GetControlInfo(Optional ByVal pCtrlCd As String = "") As Recordset
    Dim strSql As String
    
    strSql = " select a.ctrlcd,a.ctrlnm,a.levelcd,a.ctrldiv,a.eqpcd,c.eqpnm, a.buildcd,d.field1 as buildnm, " & _
            " a.sectcd,e.field1 as sectnm, a.workarea, f.field1 as workareanm " & _
            " from " & T_LAB021 & " a, " & T_LAB006 & " c, " & T_LAB032 & " d, " & T_LAB032 & " e, " & T_LAB032 & " f " & _
            " where " & DBJ("a.eqpcd*=c.eqpcd") & _
            " and " & DBW("d.cdindex=", LC3_Buildings) & _
            " and a.buildcd=d.cdval1 " & _
            " and " & DBW("e.cdindex=", LC3_Section) & _
            " and a.sectcd=e.cdval1 " & _
            " and " & DBW("f.cdindex=", LC3_WorkArea) & _
            " and a.workarea=f.cdval1 "

'and a.ctrlcd='test'
'and a.levelcd='N'
    If pCtrlCd <> "" Then
        strSql = strSql & " and " & DBW("a.ctrlcd=", pCtrlCd)
    End If
    
    strSql = strSql & " order by a.ctrlcd,ctrlnm,levelcd"
            
    Set GetControlInfo = New Recordset
    GetControlInfo.Open strSql, DBConn
End Function

Private Function GetLotNo(Optional ByVal pLotNo As String = "") As Recordset
    Dim strSql As String
    
    strSql = " select a.lotno,a.opendt,a.expdt,a.makecd,a.remark,b.ctrlnm from " & T_LAB023 & " a, " & T_LAB021 & " b " & _
            " where " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text)) & _
            " and " & DBW("a.levelcd=", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))) & _
            " and a.ctrlcd=b.ctrlcd " & _
            " and a.levelcd=b.levelcd "
    
    If pLotNo <> "" Then
        strSql = strSql & " and " & DBW("lotno=", pLotNo)
    End If
    
    strSql = strSql & " order by opendt desc"
    
    Set GetLotNo = New Recordset
    GetLotNo.Open strSql, DBConn
End Function

Private Sub cmdPopLotNo_Click()
    Dim objPop As clsPopUpList
    Dim i As Long
    
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    
    cmdPrevData.Enabled = False
    
    Set objPop = New clsPopUpList

    With objPop
        .Recordset = GetLotNo
        .FormCaption = "Lot Number 찾기"
        .Delimiter = COL_DIV
        .ColumnHeaderText = "Lot No시작일만료일제조사리마크컨트롤명"
        .ColumnHeaderWidth = "108010201020000"
        .ColumnHeaderAlign = "022"
        
        Call .LoadPopUp
        
        On Error Resume Next
        txtLotNo.Text = medGetP(.SelectedString, 1, .Delimiter)
        dtpOpenDt.Value = Format(medGetP(.SelectedString, 2, .Delimiter), "####-##-##")
        dtpExpDt.Value = Format(medGetP(.SelectedString, 3, .Delimiter), "####-##-##")
        txtMakeCd.Text = medGetP(.SelectedString, 4, .Delimiter)
        txtRemark.Text = medGetP(.SelectedString, 5, .Delimiter)
        
        DoEvents
    End With
    
    Call LoadControlMaster
    
    Set objPop = Nothing
End Sub

Private Sub cmdPrevData_Click()
    Dim objPop As clsPopUpList
    Dim i As Long
    
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    
    Set objPop = New clsPopUpList

    With objPop
        .Recordset = GetLotNo
        .FormCaption = "Lot Number 찾기"
        .Delimiter = COL_DIV
        .ColumnHeaderText = "Lot No시작일만료일제조사리마크컨트롤명"
        .ColumnHeaderWidth = "108010201020000"
        .ColumnHeaderAlign = "022"
        
        Call .LoadPopUp
        
        On Error Resume Next
'        txtLotNo.Text = medGetP(.SelectedString, 1, .DELIMITER)
        dtpOpenDt.Value = Format(medGetP(.SelectedString, 2, .Delimiter), "####-##-##")
        dtpExpDt.Value = Format(medGetP(.SelectedString, 3, .Delimiter), "####-##-##")
        txtMakeCd.Text = medGetP(.SelectedString, 4, .Delimiter)
        txtRemark.Text = medGetP(.SelectedString, 5, .Delimiter)
        
        DoEvents
    End With
    
    Call LoadControlMaster(Trim(medGetP(objPop.SelectedString, 1, objPop.Delimiter)))
    
    Set objPop = Nothing
End Sub

Private Sub cmdSave_Click()
'023지우고
'024지우고
'023인써트 하고
'024인써트 한다.
    Dim objSQL As clsLISSqlQc
    Dim strMsg As VbMsgBoxResult
    Dim arySQL() As String
    Dim i As Long, j As Long
    Dim strTestcd As String
    Dim strMeanVal As String
    Dim strSdVal As String
    Dim strAvalval As String
    Dim strRefCd As String
    Dim strRstUnit As String
    Dim strCvVal As String
    Dim strMinVal As String
    Dim strMaxVal As String
    Dim strWmSet As String

    If CheckValidation = False Then Exit Sub
    
    strMsg = MsgBox("현재 작성된 데이터를 저장합니다." & vbNewLine & _
                    "과거의 자료가 존재했을 경우 현재의 자료로 대치됩니다." & vbNewLine & vbNewLine & _
                    "계속 진행하시겠습니까?", vbExclamation + vbYesNo)
    
    If strMsg = vbNo Then Exit Sub
    
    Set objSQL = New clsLISSqlQc
    
    ReDim arySQL(2)
    
    arySQL(0) = objSQL.SqlMstDeleteAll(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), Trim(txtLotNo.Text), "1")
    arySQL(1) = objSQL.SqlMstDeleteAll(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), Trim(txtLotNo.Text), "2")
    
'    arySQL(0) = " delete from " & T_LAB023 & _
'                " where " & DBW("ctrlcd=", Trim(txtCtrlCd.Text)) & _
'                " and " & DBW("levelcd=", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))) & _
'                " and " & DBW("lotno=", Trim(txtLotNo.Text))
'
'    arySQL(1) = " delete from " & T_LAB024 & _
'                " where " & DBW("ctrlcd=", Trim(txtCtrlCd.Text)) & _
'                " and " & DBW("levelcd=", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))) & _
                " and " & DBW("lotno=", Trim(txtLotNo.Text))
    
    arySQL(2) = " insert into " & T_LAB023 & _
                " (ctrlcd,levelcd,lotno,opendt,expdt,makecd,remark) values ( " & _
                DBV("ctrlcd", Trim(txtCtrlCd.Text), 1) & DBV("levelcd", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), 1) & _
                DBV("lotno", Trim(txtLotNo.Text), 1) & DBV("opendt", Format(dtpOpenDt.Value, "yyyyMMdd"), 1) & _
                DBV("expdt", Format(dtpExpDt.Value, "yyyyMMdd"), 1) & DBV("makecd", Trim(txtMakeCd.Text), 1) & _
                DBV("remark", Trim(txtRemark.Text)) & " ) "

    For i = 1 To tblQcMst.DataRowCnt
        With tblQcMst
            .Row = i
            .Col = 1
            If .Value <> "" Then
                strTestcd = .Value
                .Col = 3: strMeanVal = .Value
                .Col = 4: strSdVal = .Value
                .Col = 5: strAvalval = .Value
                .Col = 6: strRefCd = .Value
                .Col = 7: strRstUnit = .Value
                .Col = 8: strCvVal = .Value
                .Col = 9: strMinVal = .Value
                .Col = 10: strMaxVal = .Value
                .Col = 11: strWmSet = .Value
                .Col = 12: strWmSet = strWmSet & .Value
                .Col = 13: strWmSet = strWmSet & .Value
                .Col = 14: strWmSet = strWmSet & .Value
                .Col = 15: strWmSet = strWmSet & .Value
                
                ReDim Preserve arySQL(UBound(arySQL) + 1)
                
                arySQL(UBound(arySQL)) = " insert into " & T_LAB024 & _
                                         " (ctrlcd,levelcd,lotno,testcd,meanval,sdval,avalval,refcd,rstunit,cvval,minval, " & _
                                         " maxval,schedfg,wmset,calfg,entdt,entid ) values ( " & _
                                         DBV("ctrlcd", Trim(txtCtrlCd.Text), 1) & DBV("levelcd", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")), 1) & _
                                         DBV("lotno", Trim(txtLotNo.Text), 1) & DBV("testcd", strTestcd, 1) & DBV("meanval", strMeanVal, 1) & _
                                         DBV("sdval", strSdVal, 1) & DBV("avalval", strAvalval, 1) & DBV("refcd", strRefCd, 1) & DBV("rstunit", strRstUnit, 1) & _
                                         DBV("cvval", strCvVal, 1) & DBV("minval", strMinVal, 1) & DBV("maxval", strMaxVal, 1) & DBV("schedfg", "", 1) & _
                                         DBV("wmset", strWmSet, 1) & DBV("calfg", "", 1) & DBV("entdt", "", 1) & DBV("entid", "") & " ) "
            End If
        End With
    Next
    
    On Error GoTo ErrTrap
    
    DBConn.BeginTrans
    For j = LBound(arySQL) To UBound(arySQL)
        If arySQL(j) <> "" Then
            DBConn.Execute arySQL(j)
        End If
    Next
    DBConn.CommitTrans
    Set objSQL = Nothing
    
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    Exit Sub
    
ErrTrap:
    Set objSQL = Nothing
    DBConn.RollbackTrans
    MsgBox "처리도중 오류가 발생하였습니다." & vbNewLine & _
           Err.Description, vbCritical
End Sub

Private Function CheckValidation() As Boolean
    CheckValidation = False
    
    If Trim(txtCtrlCd.Text) = "" Then
        MsgBox "컨트롤 코드를 입력하거나 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If optLevelCd(0).Value = False And optLevelCd(1).Value = False And optLevelCd(2).Value = False Then
        MsgBox "컨트롤 레벨을 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If Trim(txtLotNo.Text) = "" Then
        MsgBox "LotNo를 입력하거나 선택하십시오.", vbExclamation
        Exit Function
    End If
    
    If DateDiff("d", dtpOpenDt.Value, dtpExpDt.Value) < 0 Then
        MsgBox "날짜 입력이 잘못되었습니다.", vbExclamation
        Exit Function
    End If
    
    tblQcMst.Row = 1
    tblQcMst.Col = 1
    If tblQcMst.Value = "" Then
        MsgBox "등록된 컨트롤이 아닙니다. 컨트롤을 먼저 등록하십시오.", vbExclamation
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Sub cmdSchedule_Click()
    If CheckValidation = False Then Exit Sub
    
    Call LoadForm(frm312QCSchedule_N, Me)
    Call frm312QCSchedule_N.CallByExternal(Trim(txtCtrlCd.Text), IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H")))
End Sub

Private Sub cmdSelAll_Click()
    Dim i  As Long
    
    For i = chkRule.LBound To chkRule.UBound
        chkRule(i).Value = 1
    Next
End Sub

Private Sub Form_Load()

    txtCtrlCd.Text = ""
    Call InitControl
    txtLotNo.Text = ""
    Call InitLotNo
    Call medClearTable(tblQcMst)
    tblQcMst.MaxRows = 18
End Sub

Private Sub InitControl()
    lblCtrlNm.Caption = ""
    lblCtrlDiv.Caption = ""
    lblEqp.Caption = ""
    lblBuilding.Caption = ""
    lblSection.Caption = ""
    lblWorkarea.Caption = ""
End Sub

Private Sub InitLotNo()
    lblLotNo.Caption = ""
    dtpOpenDt.Value = GetSystemDate
    dtpExpDt.Value = GetSystemDate
    txtMakeCd.Text = ""
    txtRemark.Text = ""
    
    cmdPrevData.Enabled = False
End Sub

Private Sub optLevelcd_Click(Index As Integer)
    On Error Resume Next
    If Screen.ActiveControl.Name <> optLevelCd(Index).Name Then Exit Sub
    
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    
    txtLotNo.Text = ""
    Call InitLotNo
    Call medClearTable(tblQcMst)
    tblQcMst.MaxRows = 18
    
    Call LoadLotNo
    Call LoadControlMaster
    
    If tblQcMst.DataRowCnt = 0 Then
        MsgBox "해당 컨트롤이 존재하지 않습니다.", vbExclamation
    End If
End Sub

Private Sub tblQcMst_Click(ByVal Col As Long, ByVal Row As Long)
   
    Dim i As Long
    Static blnToggle(11 To 15) As Boolean
    
    If Col < 11 Then Exit Sub
    If Row > 0 Then Exit Sub
    
    blnToggle(Col) = IIf(blnToggle(Col), False, True)
    
    With tblQcMst
        .Col = Col
        For i = 1 To .DataRowCnt
            .Row = i
            If .CellType = CellTypeCheckBox Then
                .Value = IIf(blnToggle(Col), 0, 1)
            End If
        Next
    End With
End Sub

Private Sub tblQcMst_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim lngDecimalPlace As Long
    
    If Col = 5 Then
        With tblQcMst
            .Row = Row
            .Col = 5
            lngDecimalPlace = .Value
            
            .Col = 3
            .TypeFloatDecimalPlaces = lngDecimalPlace
            .Col = 4
            .TypeFloatDecimalPlaces = lngDecimalPlace
            .Col = 8
            .TypeFloatDecimalPlaces = lngDecimalPlace
            .Col = 9
            .TypeFloatDecimalPlaces = lngDecimalPlace
            .Col = 10
            .TypeFloatDecimalPlaces = lngDecimalPlace
        End With
    End If
    
    If Col = 6 Then
        With tblQcMst
            .Row = Row
            .Col = 6
            If .Value <> "" Then
                .Row = Row: .Row2 = Row
                .Col = 11: .Col2 = 15
                .BlockMode = True
                .Value = ""
                .CellType = CellTypeStaticText
                .BlockMode = False
            Else
                .Row = Row: .Row2 = Row
                .Col = 11: .Col2 = 15
                .BlockMode = True
                .CellType = CellTypeCheckBox
                .TypeHAlign = TypeHAlignCenter
                .BlockMode = False
            End If
            
        End With
    End If
End Sub

' 2015.01.05 온승호 : CV, MIN, MAX 자동계산

Private Sub tblQcMst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTmp
    Dim tmpMean As String
    Dim tmpSD As String
    Dim tmpCV As String
    Dim tmpMin As String
    Dim tmpMax As String
    
    With tblQcMst
        If Col = 4 Then
            .GetText 3, Row, varTmp: tmpMean = varTmp
            .GetText 4, Row, varTmp: tmpSD = varTmp
            If tmpSD <> 0 And tmpMean <> 0 Then
                tmpCV = Round((tmpSD / tmpMean) * 100, 2)
                tmpMin = tmpMean - (tmpSD * 2)
                tmpMax = tmpMean + (tmpSD * 2)
            End If
            .SetText 8, Row, tmpCV
            .SetText 9, Row, tmpMin
            .SetText 10, Row, tmpMax
        End If
    End With
    
End Sub

Private Sub tblQcMst_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim objPop As clsPopUpList
    Dim Rs As Recordset
    Dim strSql As String
    Dim strTestcd As String
    
    If Col <> 6 Then Exit Sub
    
    With tblQcMst
        .Row = Row
        .Col = 1
        strTestcd = .Value
            
        strSql = " select cdval2,field1 from " & T_LAB031 & _
                 " where " & DBW("cdindex=", LC2_ItemResult) & _
                 " and " & DBW("cdval1=", strTestcd)
                
        Set Rs = New Recordset
        Rs.Open strSql, DBConn
        
        If Rs.EOF = False Then
            Set objPop = New clsPopUpList
            objPop.Recordset = Rs
            objPop.HideSearchTool = True
            objPop.SelectByClick = True
            objPop.FormWidth = 4635
            objPop.FormHeight = 2880
            objPop.FormCaption = "결과코드 찾기"
            objPop.ColumnHeaderText = "결과코드;결과코드명"
            objPop.ColumnHeaderWidth = "1110.047;3075.024"
            objPop.LoadPopUp
            
            .Col = 6
            .Value = medGetP(objPop.SelectedString, 1, objPop.Delimiter)
        End If
            
    End With
    
    Set Rs = Nothing
    Set objPop = Nothing
End Sub

Private Sub txtCtrlCd_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> txtCtrlCd.Name Then Exit Sub
    
    If lblCtrlNm.Caption <> "" Then
        Call InitControl
        txtLotNo.Text = ""
        Call InitLotNo
        Call medClearTable(tblQcMst)
        tblQcMst.MaxRows = 18
    End If
End Sub

Private Sub txtCtrlCd_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtCtrlCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCtrlCd_LostFocus()
    Dim Rs As Recordset
'이따구루 밖에 못할까? 나중에 다른 방법으로 고쳐야지...

    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If Trim(lblCtrlNm.Caption) <> "" Then Exit Sub
    
    DoEvents
    Set Rs = GetControlInfo(Trim(txtCtrlCd.Text))
    
    If Rs.EOF = False Then
        DoEvents
        Call LoadControlInfo(Trim(txtCtrlCd.Text))
        DoEvents
        Call LoadLotNo
        DoEvents
        Call LoadControlMaster
    End If
    
    Set Rs = Nothing
End Sub

Private Sub txtLotNo_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> txtLotNo.Name Then Exit Sub
    
    Call InitLotNo
    Call medClearTable(tblQcMst)
    tblQcMst.MaxRows = 18
End Sub

Private Sub txtLotNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtLotNo.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtLotNo_LostFocus()
    
    If Trim(txtLotNo.Text) = "" Then Exit Sub
    If lblLotNo.Caption <> "" Then Exit Sub
    
    Call LoadLotNo
    Call LoadControlMaster
    
    lblLotNo.Caption = txtLotNo.Text
End Sub

Private Sub LoadLotNo()
    Dim Rs As Recordset
    
    Set Rs = GetLotNo(Trim(txtLotNo.Text))
    
    If Rs.EOF = False Then
        txtLotNo.Text = Rs.Fields("lotno").Value & ""
        lblLotNo.Caption = txtLotNo.Text
        dtpOpenDt.Value = Format(Rs.Fields("opendt").Value & "", "####-##-##")
        dtpExpDt.Value = Format(Rs.Fields("expdt").Value & "", "####-##-##")
        txtMakeCd.Text = Rs.Fields("makecd").Value & ""
        txtRemark.Text = Rs.Fields("remark").Value & ""
        
        lblCtrlNm.Caption = Rs.Fields("ctrlnm").Value & ""
    Else
        cmdPrevData.Enabled = True
    End If
    
    Set Rs = Nothing
End Sub

Private Sub LoadControlMaster(Optional ByVal pLotNo As String = "")
    Dim Rs As Recordset
    Dim strSql As String
    Dim lngDecimalPlace As Long
    
    '## 5.0.1: 이상대(2005-04-13)
    '   - 지정검체 마스터에 폐기된 코드와 폐기되지 않은 코드가 seq가 동일할 경우
    '     중복된 검사항목을 조회하여 쿼리수정
    '## 5.0.2: 이상대(2005-05-14)
    '   - 지정검체 마스터에서 폐기된 코드가 조회되어 쿼리수정
    strSql = " select a.ctrlcd,a.levelcd,b.lotno, a.testcd,c.testnm, b.meanval,b.sdval, " & _
            " b.avalval,d.avalval as avalval004,b.refcd,b.rstunit,d.rstunit as rstunit004 , " & _
            " b.cvVal , b.minval, b.maxval, b.wmset " & _
            " from " & T_LAB004 & " d, " & T_LAB001 & " c, " & T_LAB024 & " b, " & T_LAB022 & " a" & _
            " where " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text)) & _
            " and " & DBW("a.levelcd=", IIf(optLevelCd(0).Value, "L", IIf(optLevelCd(1).Value, "N", "H"))) & _
            " and " & DBJ("a.ctrlcd*=b.ctrlcd") & _
            " and " & DBJ("a.levelcd*=b.levelcd") & _
            " and " & DBJ("a.testcd*=b.testcd") & _
            " and " & DBJ("b.lotno=*" & DBS(IIf(pLotNo = "", Trim(txtLotNo.Text), pLotNo))) & _
            " and a.testcd=c.testcd " & _
            " and c.applydt=(select max(applydt) from " & T_LAB001 & " where testcd=a.testcd and (expdt='' or expdt is null)) " & _
            " and a.testcd=d.testcd " & _
            " and d.seq=(select min(seq) from " & T_LAB004 & " where testcd=a.testcd and (expdt='' or expdt is null)) " & _
            " and (d.expdt='' or d.expdt is null)" & _
            " order by a.testcd "

    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    Call medClearTable(tblQcMst)
    tblQcMst.MaxRows = 18
    
    With tblQcMst
        .ReDraw = False
        
        Do Until Rs.EOF
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                End If
                
                .Row = .DataRowCnt + 1
                
                lngDecimalPlace = IIf(Rs.Fields("avalval").Value & "" = "", Rs.Fields("avalval004").Value & "", Rs.Fields("avalval").Value & "")
                
                .Col = 1: .Value = Rs.Fields("testcd").Value & ""
                .Col = 2: .Value = Rs.Fields("testnm").Value & ""
                .Col = 3: .TypeFloatDecimalPlaces = IIf(lngDecimalPlace = 9, 0, lngDecimalPlace)
                          .Value = Rs.Fields("meanval").Value & ""
                .Col = 4: .TypeFloatDecimalPlaces = IIf(lngDecimalPlace = 9, 0, lngDecimalPlace)
                          .Value = Rs.Fields("sdval").Value & ""
                .Col = 5: .Value = IIf(Rs.Fields("avalval").Value & "" = "", Rs.Fields("avalval004").Value & "", Rs.Fields("avalval").Value & "")
                .Col = 6: .Value = Rs.Fields("refcd").Value & ""
                .Col = 7: .Value = IIf(Rs.Fields("rstunit").Value & "" = "", Rs.Fields("rstunit004").Value & "", Rs.Fields("rstunit").Value & "")
                .Col = 8: .TypeFloatDecimalPlaces = IIf(lngDecimalPlace = 9, 0, lngDecimalPlace)
                          .Value = Rs.Fields("cvval").Value & ""
                .Col = 9: .TypeFloatDecimalPlaces = IIf(lngDecimalPlace = 9, 0, lngDecimalPlace)
                          .Value = Rs.Fields("minval").Value & ""
                .Col = 10: .TypeFloatDecimalPlaces = IIf(lngDecimalPlace = 9, 0, lngDecimalPlace)
                           .Value = Rs.Fields("maxval").Value & ""
                .Col = 11: .Value = Val(Mid(Rs.Fields("wmset").Value & "", 1, 1))
                .Col = 12: .Value = Val(Mid(Rs.Fields("wmset").Value & "", 2, 1))
                .Col = 13: .Value = Val(Mid(Rs.Fields("wmset").Value & "", 3, 1))
                .Col = 14: .Value = Val(Mid(Rs.Fields("wmset").Value & "", 4, 1))
                .Col = 15: .Value = Val(Mid(Rs.Fields("wmset").Value & "", 5, 1))
                
                .Row = -1
                .Col = 2: .Col2 = 2
                .BlockMode = True
                .ForeColor = &H864B24
                .BlockMode = False
        
            Rs.MoveNext
        Loop
        
        .ReDraw = True
    End With
    
    Set Rs = Nothing
End Sub
