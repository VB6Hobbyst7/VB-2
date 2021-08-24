VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIIS602 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "검사코드 관리"
   ClientHeight    =   8925
   ClientLeft      =   4080
   ClientTop       =   285
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TabStrip tabTestCd 
      Height          =   315
      Left            =   75
      TabIndex        =   43
      Top             =   1275
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   556
      Style           =   2
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2003-12-12"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   6630
      Left            =   5445
      TabIndex        =   30
      Top             =   2055
      Width           =   5670
      Begin MSComctlLib.ListView lvwRefList 
         Height          =   2010
         Left            =   90
         TabIndex        =   57
         Top             =   4560
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   3545
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16252919
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "성별"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "일령"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "기준치"
            Object.Width           =   3616
         EndProperty
      End
      Begin VB.CommandButton cmdSpc 
         BackColor       =   &H00DBE6E6&
         Caption         =   "지정검체등록"
         Height          =   420
         Left            =   4320
         Style           =   1  '그래픽
         TabIndex        =   51
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00DBE6E6&
         Caption         =   "참고치 등록"
         Height          =   420
         Left            =   4320
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   3690
         Width           =   1215
      End
      Begin MedControls1.LisLabel lblSpcCd 
         Height          =   345
         Left            =   1860
         TabIndex        =   45
         Top             =   285
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   345
         Left            =   1860
         TabIndex        =   46
         Top             =   717
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPanicTo 
         Height          =   345
         Left            =   3735
         TabIndex        =   47
         Top             =   2877
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeltaTo 
         Height          =   345
         Left            =   4035
         TabIndex        =   48
         Top             =   3315
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSpcUnit 
         Height          =   345
         Left            =   1860
         TabIndex        =   52
         Top             =   2013
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblAvalVal 
         Height          =   345
         Left            =   1860
         TabIndex        =   53
         Top             =   2445
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPanicFr 
         Height          =   345
         Left            =   1860
         TabIndex        =   54
         Top             =   2877
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeltaFr 
         Height          =   345
         Left            =   2160
         TabIndex        =   55
         Top             =   3315
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MSComctlLib.TabStrip tabRefList 
         Height          =   315
         Left            =   135
         TabIndex        =   56
         Top             =   4185
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   556
         Style           =   2
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "2003-10-10"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MedControls1.LisLabel lblSpcApplyDt 
         Height          =   345
         Left            =   1860
         TabIndex        =   71
         Top             =   1149
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSpcExpDt 
         Height          =   345
         Left            =   1860
         TabIndex        =   72
         Top             =   1581
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         BackColor       =   16252919
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
         LeftGab         =   100
      End
      Begin VB.Label Label27 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "%"
         Height          =   180
         Left            =   5010
         TabIndex        =   76
         Top             =   3405
         Width           =   300
      End
      Begin VB.Label Label26 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "%"
         Height          =   180
         Left            =   3135
         TabIndex        =   75
         Top             =   3405
         Width           =   300
      End
      Begin VB.Label Label25 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "(+)"
         Height          =   180
         Left            =   3720
         TabIndex        =   74
         Top             =   3405
         Width           =   300
      End
      Begin VB.Label Label24 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "(-)"
         Height          =   180
         Left            =   1845
         TabIndex        =   73
         Top             =   3405
         Width           =   300
      End
      Begin VB.Label Label23 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참 고 치 정 보"
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
         Left            =   1260
         TabIndex        =   70
         Top             =   3840
         Width           =   1305
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체코드 :"
         Height          =   180
         Left            =   240
         TabIndex        =   69
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체명 :"
         Height          =   180
         Left            =   240
         TabIndex        =   68
         Top             =   792
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "적용일 :"
         Height          =   180
         Left            =   240
         TabIndex        =   67
         Top             =   1224
         Width           =   660
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "폐기일 :"
         Height          =   180
         Left            =   240
         TabIndex        =   66
         Top             =   1656
         Width           =   660
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과단위 :"
         Height          =   180
         Left            =   240
         TabIndex        =   65
         Top             =   2088
         Width           =   840
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "유효숫자 :"
         Height          =   180
         Left            =   240
         TabIndex        =   64
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Panic Check :"
         Height          =   180
         Left            =   240
         TabIndex        =   63
         Top             =   2952
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Delta Check :"
         Height          =   180
         Left            =   240
         TabIndex        =   62
         Top             =   3390
         Width           =   1140
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00DBE6E6&
         FillStyle       =   0  '단색
         Height          =   390
         Left            =   120
         Top             =   4155
         Width           =   5415
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "~"
         Height          =   180
         Left            =   3375
         TabIndex        =   49
         Top             =   2985
         Width           =   135
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   105
         Top             =   3735
         Width           =   3495
      End
   End
   Begin MSComctlLib.TabStrip tabSpcList 
      Height          =   315
      Left            =   5475
      TabIndex        =   44
      Top             =   1275
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   556
      Style           =   2
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Serum"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Urine"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   7050
      Left            =   45
      TabIndex        =   29
      Top             =   1635
      Width           =   5340
      Begin VB.CommandButton cmdGroup 
         BackColor       =   &H00DBE6E6&
         Caption         =   "그룹항목설정"
         Height          =   495
         Left            =   3855
         Style           =   1  '그래픽
         TabIndex        =   78
         Top             =   5655
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저 장(&S)"
         Height          =   495
         Left            =   270
         Style           =   1  '그래픽
         TabIndex        =   77
         Top             =   195
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "취 소(&C)"
         Height          =   495
         Left            =   4230
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   195
         Width           =   990
      End
      Begin VB.ComboBox cboWorkarea 
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmIIS602.frx":0000
         Left            =   1770
         List            =   "frmIIS602.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   3285
         Width           =   2835
      End
      Begin VB.CheckBox chkDetailFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "상세항목"
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Tag             =   "35102"
         Top             =   5760
         Width           =   1110
      End
      Begin VB.CommandButton cmdDetail 
         BackColor       =   &H00DBE6E6&
         Caption         =   "상세항목설정"
         Height          =   495
         Left            =   2490
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   5655
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "삭 제(&D)"
         Height          =   495
         Left            =   3240
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   195
         Width           =   990
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00DBE6E6&
         Caption         =   "수 정(&M)"
         Height          =   495
         Left            =   2250
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   195
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "추 가(&A)"
         Height          =   495
         Left            =   1260
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   195
         Width           =   990
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   2865
         TabIndex        =   42
         Top             =   5130
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtItemSeq 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   13
         Text            =   "0"
         Top             =   5160
         Width           =   1110
      End
      Begin VB.ComboBox cboRstType 
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmIIS602.frx":0004
         Left            =   1770
         List            =   "frmIIS602.frx":0014
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   4710
         Width           =   1680
      End
      Begin VB.PictureBox picRstDiv 
         BackColor       =   &H00F7FFF7&
         Height          =   360
         Left            =   1770
         ScaleHeight     =   300
         ScaleWidth      =   3300
         TabIndex        =   41
         Top             =   4215
         Width           =   3360
         Begin VB.OptionButton optRstDiv 
            BackColor       =   &H00F7FFF7&
            Caption         =   "Required"
            Height          =   300
            Index           =   0
            Left            =   75
            TabIndex        =   10
            Tag             =   "35136"
            Top             =   15
            Width           =   1560
         End
         Begin VB.OptionButton optRstDiv 
            BackColor       =   &H00F7FFF7&
            Caption         =   "Alternative"
            Height          =   300
            Index           =   1
            Left            =   1740
            TabIndex        =   11
            Tag             =   "35135"
            Top             =   15
            Width           =   1350
         End
      End
      Begin VB.PictureBox picPanelFg 
         BackColor       =   &H00F7FFF7&
         Height          =   360
         Left            =   1770
         ScaleHeight     =   300
         ScaleWidth      =   3300
         TabIndex        =   40
         Top             =   3720
         Width           =   3360
         Begin VB.OptionButton optPanelFg 
            BackColor       =   &H00F7FFF7&
            Caption         =   "없음"
            Height          =   300
            Index           =   0
            Left            =   75
            TabIndex        =   7
            Tag             =   "35135"
            Top             =   15
            Width           =   780
         End
         Begin VB.OptionButton optPanelFg 
            BackColor       =   &H00F7FFF7&
            Caption         =   "그룹처방"
            Height          =   300
            Index           =   1
            Left            =   975
            TabIndex        =   8
            Tag             =   "35135"
            Top             =   15
            Width           =   1095
         End
         Begin VB.OptionButton optPanelFg 
            BackColor       =   &H00F7FFF7&
            Caption         =   "상세검사"
            Height          =   300
            Index           =   2
            Left            =   2160
            TabIndex        =   9
            Tag             =   "35136"
            Top             =   15
            Width           =   1095
         End
      End
      Begin VB.TextBox txtTestNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2805
         Width           =   3360
      End
      Begin VB.TextBox txtTestNm10 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2340
         Width           =   2130
      End
      Begin VB.TextBox txtTestNm5 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   1770
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1875
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpApplyDt 
         Height          =   330
         Left            =   1770
         TabIndex        =   1
         Top             =   930
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   16252919
         Format          =   63373313
         CurrentDate     =   37994
      End
      Begin MSComCtl2.DTPicker dtpExpDt 
         Height          =   330
         Left            =   1770
         TabIndex        =   2
         Top             =   1395
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   63373313
         CurrentDate     =   38001
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "WorkArea :"
         Height          =   180
         Left            =   240
         TabIndex        =   59
         Top             =   3345
         Width           =   915
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "출력순서 :"
         Height          =   180
         Left            =   240
         TabIndex        =   39
         Top             =   5220
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과유형 :"
         Height          =   180
         Left            =   240
         TabIndex        =   38
         Top             =   4755
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과구분 :"
         Height          =   180
         Left            =   240
         TabIndex        =   37
         Top             =   4275
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "처방구분 :"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   3810
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사명(전체) :"
         Height          =   180
         Left            =   240
         TabIndex        =   35
         Top             =   2865
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사명(약어10) :"
         Height          =   180
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사명(약어5) :"
         Height          =   180
         Left            =   240
         TabIndex        =   33
         Top             =   1935
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "폐기일 :"
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   1455
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "적용일 :"
         Height          =   180
         Left            =   240
         TabIndex        =   31
         Top             =   990
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   765
      Left            =   45
      TabIndex        =   23
      Top             =   0
      Width           =   11085
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H00DBE6E6&
         Caption         =   "<< 이전(&P)"
         Height          =   495
         Left            =   7260
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdHopSrh 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Left            =   6885
         Picture         =   "frmIIS602.frx":004C
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   270
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtHopTestCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   5325
         MaxLength       =   20
         TabIndex        =   27
         Top             =   285
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdTestSrh 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Left            =   2925
         Picture         =   "frmIIS602.frx":0E8E
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   270
         Width           =   405
      End
      Begin VB.TextBox txtTestCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   0
         Top             =   285
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00DBE6E6&
         Caption         =   "종 료(&X)"
         Height          =   495
         Left            =   9690
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00DBE6E6&
         Caption         =   "다음(&N) >>"
         Height          =   495
         Left            =   8475
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "병원 검사코드 :"
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
         Left            =   3840
         TabIndex        =   25
         Top             =   345
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사코드 :"
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
         Left            =   315
         TabIndex        =   24
         Top             =   345
         Width           =   930
      End
   End
   Begin MSComctlLib.TabStrip tabSpcDt 
      Height          =   315
      Left            =   5475
      TabIndex        =   61
      Top             =   1695
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   556
      Style           =   2
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2003-10-10"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2004-01-10"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00DBE6E6&
      FillStyle       =   0  '단색
      Height          =   390
      Left            =   5460
      Top             =   1665
      Width           =   5655
   End
   Begin VB.Label Label22 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "지 정 검 체 정 보"
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
      Left            =   6525
      TabIndex        =   60
      Top             =   930
      Width           =   1575
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검 사 항 목 정 보"
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
      Left            =   1125
      TabIndex        =   58
      Top             =   930
      Width           =   1575
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Left            =   45
      Top             =   825
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00DBE6E6&
      FillStyle       =   0  '단색
      Height          =   390
      Left            =   60
      Top             =   1245
      Width           =   5325
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00DBE6E6&
      FillStyle       =   0  '단색
      Height          =   390
      Left            =   5460
      Top             =   1245
      Width           =   5655
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Left            =   5445
      Top             =   825
      Width           =   3495
   End
End
Attribute VB_Name = "frmIIS602"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS602.frm
'   작성자  : 이상대
'   내  용  : 검사코드 마스터
'   작성일  : 2004-01-27
'   버  전  :
'   메  모  :
'-----------------------------------------------------------------------------'

Option Explicit

'## StateEnum
Private Enum StateEnum
    ccInit              '최초상태
    ccSave              '저장할수 있는 상태
    ccAdd               '새 적용일을 추가하는 상태
    ccModify            '수정, 삭제할수 있는상태
End Enum

'## ClearEnum
Private Enum ClearEnum
    ccAll               '전체 컨트롤 초기화
    ccCmdAdd            'cmdAdd 클릭시 컨트롤 초기화
    ccTabSpcNm          'tabSpcNm 클릭시 컨트롤 초기화
    ccTabSpcDt          'tabSpcDt 클릭시 컨트롤 초기화
End Enum
    
Private mTMaster As clsIISTMaster       '검사코드 마스터 클래스
Private mTestCd  As String              '현재 검사코드
Private mState   As StateEnum           '폼의 버튼상태

Private WithEvents mCode As clsIISCodeList      'CodeList 클래스
Attribute mCode.VB_VarHelpID = -1

Private Sub Form_Load()
    Set mTMaster = New clsIISTMaster
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight: .Width = 11270
    End With
    
    '## 화면지움, Form Show
    Call CtlClear(ccAll)
    Call CtlLock(ccInit)
    Call Me.Show
    DoEvents
    
    '## 검사코드 로드, Workarea 로드
    Call mTMaster.GetTestCdList
    Call GetWorkarea
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS602").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTMaster = Nothing
    Set frmIIS602 = Nothing
End Sub

Private Sub cmdNext_Click()
    Dim strTestCd As String     '검사코드
    
    strTestCd = UCase(Trim(txtTestCd.Text))
    If strTestCd = "" Then Exit Sub
    If mTMaster.Exist(strTestCd) = False Then Exit Sub

    mTestCd = mTMaster.NextTestCd(strTestCd)
    If mTestCd = strTestCd Then Exit Sub
    
    Call CtlClear(ccAll)
    txtTestCd.Text = mTestCd
    Call GetTestCds
    Call CtlLock(ccModify)
End Sub

Private Sub cmdPrev_Click()
    Dim strTestCd As String     '검사코드
    
    strTestCd = UCase(Trim(txtTestCd.Text))
    If strTestCd = "" Then Exit Sub
    If mTMaster.Exist(strTestCd) = False Then Exit Sub

    mTestCd = mTMaster.PrevTestCd(strTestCd)
    If mTestCd = strTestCd Then Exit Sub
    
    Call CtlClear(ccAll)
    txtTestCd.Text = mTestCd
    Call GetTestCds
    Call CtlLock(ccModify)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdTestSrh_Click()
    Set mCode = New clsIISCodeList
    With mCode
        .Caption = "검사코드 리스트"
        .HeaderCd = "검사코드"
        .HeaderCdNm = "검사명"
        .CodeListByCol mTMaster.TestCds
    End With
    Set mCode = Nothing
    
    Call txtTestCd_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdHopSrh_Click()
'
End Sub

Private Sub cmdSave_Click()
    Dim strOldDt        As String   '수정하려는 검사코드의 적용일
    Dim strLastDt       As String   '검사코드의 적용일중 최근 적용일
    Dim strApplyDt      As String   '적용일
    Dim strExpireDt     As String   '폐기일
    Dim strTestNm5      As String   '검사명(5)
    Dim strTestNm10     As String   '검사명(10)
    Dim strTestNm       As String   '검사명(전체)
    Dim strWorkarea     As String   'Workarea
    Dim strPanelFg      As String   '처방구분(Null: 단일항목, G: 그룹항목, D: 상세항목)
    Dim strRstDiv       As String   'Alternative, Require 유형(A: Alternative, R: Require)
    Dim strRstType      As String   '결과유형(Null: 일반, F: Free, N: Numeric, A: Alpha)
    Dim strDetailFg     As String   'Detail 항목여부(Null: 없음, *:상세항목 자코드)
    Dim lngRptSeq       As String   '출력순서
    Dim blnReturn       As Boolean
    
    '## 1.기존 검사코드를 수정후 저장하는 경우
    '     - 수정
    '     - 적용일 체크 불필요
    '     - mTmaster.Exist()==True, mState==ccModify
    '## 2.기존 검사코드에 새 적용일을 저장하는 경우
    '     - 입력
    '     - 적용일 체크
    '     - mTmaster.Exist()==True, mState==ccAdd
    '## 3.신규 검사코드를 저장하는 경우
    '     - 입력
    '     - 적용일 체크 불필요
    '     - mTMaster.Exist()==False, mState==ccAdd
    
    '## 적용일 Check
    '## 기존 검사코드에 새 적용일을 추가할경우 새 적용일은
    '## 이전 적용일, 현재일자보다 반드시 커야한다.
    
    strLastDt = mTMaster.GetTestCdLastApplyDt(mTestCd)
    strApplyDt = Format$(dtpApplyDt.Value, "YYYYMMDD")
    If mTMaster.Exist(mTestCd) And mState = ccAdd Then
        If strApplyDt <= strLastDt Then
            MsgBox "적용일 잘못되었습니다. 수정후 다시 저장하세요.", vbInformation, "정보"
            Exit Sub
        End If
    End If

    '## 검사명 Check
    strTestNm5 = Trim(txtTestNm5.Text)
    strTestNm10 = Trim(txtTestNm10.Text)
    strTestNm = Trim(txtTestNm.Text)
    If strTestNm5 = "" Or strTestNm10 = "" Or strTestNm = "" Then
        MsgBox "검사명(약어포함)을 모두 입력하세요.", vbInformation, "정보"
        Exit Sub
    End If
    
    '## Workarea Check
    If cboWorkarea.ListIndex = -1 Then
        MsgBox "WorkArea를 선택하세요.", vbInformation, "정보"
        Exit Sub
    End If
    strWorkarea = Trim(mGetP(cboWorkarea.Text, 1, Space(5)))
    
    '## 나머지 정보
    strExpireDt = Format$(dtpExpDt.Value, "YYYYMMDD")
    strRstType = Trim(mGetP(cboRstType.Text, 1, Space(5)))
    
    If Trim(txtItemSeq.Text) = "" Then
        lngRptSeq = 0
    Else
        lngRptSeq = CLng(Trim(txtItemSeq.Text))
    End If
    
    If optPanelFg(0).Value = True Then
        strPanelFg = ""
    ElseIf optPanelFg(1).Value = True Then
        strPanelFg = "G"
    Else
        strPanelFg = "D"
    End If
    
    If optRstDiv(0).Value = True Then
        strRstDiv = "R"
    Else
        strRstDiv = "A"
    End If
    
    If chkDetailFg.Value = "1" Then
        strDetailFg = "*"
    Else
        strDetailFg = ""
    End If
    
    '## DB에 저장
    Me.MousePointer = vbHourglass
    If mState = ccModify Then
        '## Update
        blnReturn = mTMaster.ModifyTestCd(mTestCd, strApplyDt, strExpireDt, strTestNm5, strTestNm10, _
            strTestNm, strWorkarea, strRstType, strRstDiv, strPanelFg, strDetailFg, lngRptSeq)
    Else
        '## Insert
        blnReturn = mTMaster.AddTestCd(mTestCd, strApplyDt, strExpireDt, strTestNm5, strTestNm10, _
            strTestNm, strWorkarea, strRstType, strRstDiv, strPanelFg, strDetailFg, lngRptSeq)
    End If
    
    '## 현재 입력된 검사코드에 대해 수정상태로 변경
    If blnReturn = True Then
        Call GetTestCds
        Call CtlLock(ccModify)
        mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
    Else
        mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
    Call CtlClear(ccCmdAdd)
    Call CtlLock(ccAdd)
    dtpExpDt.SetFocus
End Sub

Private Sub cmdModify_Click()
    Call CtlLock(ccSave)
    dtpExpDt.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim strApplyDt  As String       '적용일
    Dim lngReturn   As Long
    Dim intTemp     As Integer
    
    intTemp = MsgBox("정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub
    
    strApplyDt = Format$(tabTestCd.SelectedItem.Caption, "YYYYMMDD")
    
    '## 해당 검사코드에 적용일이 남어있으면 첫번째 적용일에 대한 정보를 표시
    '## 적용이 없으면 모두 화면지움
    lngReturn = mTMaster.RemoveTestCd(mTestCd, strApplyDt)
    If lngReturn = -1 Then
        mdiIISMain.sbrStatus.Panels(2).Text = "삭제중 에러가 발생했습니다."
    ElseIf lngReturn = 0 Then
        Call CtlClear(ccAll)
        Call CtlLock(ccInit)
        mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
        txtTestCd.SetFocus
    Else
        Call GetTestCds
        Call CtlLock(ccModify)
        mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim intTemp As Integer
    
    intTemp = MsgBox("변경된 내용을 취소할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub
    
    '## 기존 검사코드인 경우 현재검사코드+적용일에 대한 정보를 표시
    '## 신규 검사코드인 경우 컨트롤 초기화
    If mTMaster.Exist(mTestCd) Then
        tabTestCd.Tabs(1).Selected = True
        Call CtlLock(ccModify)
    Else
        Call CtlClear(ccAll)
        Call CtlLock(ccInit)
    End If
    txtTestCd.SetFocus
End Sub

Private Sub cmdDetail_Click()
    If mTestCd = "" Then Exit Sub
    
    With frmIIS606
        .TestCd = mTestCd
        .Show
        .ZOrder 0
    End With
End Sub

Private Sub cmdGroup_Click()
    If mTestCd = "" Then Exit Sub
    
    With frmIIS607
        .TestCd = mTestCd
        .Show
        .ZOrder 0
    End With
End Sub

Private Sub cmdSpc_Click()
    If mTestCd = "" Then Exit Sub
    
    With frmIIS603
        .TestCd = mTestCd
        .TMaster = mTMaster
        .Show
        .ZOrder 0
    End With
End Sub

Private Sub cmdRef_Click()
    If mTestCd = "" Or lblSpcCd.Caption = "" Then Exit Sub
    
    With frmIIS604
        .TestCd = mTestCd
        .TMaster = mTMaster
        .Show
        .ZOrder 0
    End With
End Sub

Private Sub txtTestCd_GotFocus()
    With txtTestCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestCd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn    '## Enter키가 입력되면 정보표시
            '## 신규/기존 검사코드인지 판단하여 기존 검사코드이면 정보표시
            '## 신규 검사코드이면 입력할수 있는 상태로 변경
            mTestCd = UCase(Trim(txtTestCd.Text))
            If mTestCd = "" Then Exit Sub
            
            Call CtlClear(ccAll)
            txtTestCd.Text = mTestCd
            If mTMaster.Exist(mTestCd) Then
                Call GetTestCds
                Call CtlLock(ccModify)
            Else
                Call CtlLock(ccAdd)
            End If
            SendKeys "{TAB}"
        Case vbKeyDown      '## 화살표 Down키가 입력되면 팝업 코드리스트를 표시
            Call cmdTestSrh_Click
    End Select
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    
    '## 숫자, 문자, Enter, Backspcace만 입력할수 있도록함
    If KeyAscii >= 65 And KeyAscii <= 90 Then Exit Sub
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    
    KeyAscii = 0
End Sub

Private Sub txtTestNm5_GotFocus()
    With txtTestNm5
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestNm5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTestNm10_GotFocus()
    With txtTestNm10
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestNm10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTestNm_GotFocus()
    With txtTestNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtItemSeq_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub optPanelFg_Click(Index As Integer)
    Select Case Index
        Case 0
            cmdDetail.Visible = False
            cmdGroup.Visible = False
            chkDetailFg.Visible = True
        Case 1
            cmdDetail.Visible = False
            cmdGroup.Visible = True
            chkDetailFg.Visible = False
        Case 2
            cmdDetail.Visible = True
            cmdGroup.Visible = False
            chkDetailFg.Visible = False
    End Select
End Sub

Private Sub tabTestCd_Click()
    Dim objFTestCds As clsIISTestCdFulls    '검사코드 컬렉션
    Dim strApplyDt  As String               '적용일
    
    '## 검사코드+적용일에 해당하는 검사코드정보 표시
    Set objFTestCds = mTMaster.FTestCds
    strApplyDt = Format(tabTestCd.SelectedItem.Caption, "YYYYMMDD")
    
    With objFTestCds(mTestCd, strApplyDt)
        dtpApplyDt.Value = Format$(.Applydt, "####-##-##")
        dtpExpDt.Value = Format$(.ExpireDt, "####-##-##")
        txtTestNm5.Text = .TestNm5
        txtTestNm10.Text = .TestNm10
        txtTestNm.Text = .TestNm
        txtItemSeq.Text = CStr(.RptSeq)
        cboWorkarea.ListIndex = mFindCombo(cboWorkarea, .Workarea)
        cboRstType.ListIndex = mFindCombo(cboRstType, .RstType)
        Select Case .PanelFg
            Case "G": optPanelFg(1).Value = True
            Case "D": optPanelFg(2).Value = True
            Case Else: optPanelFg(0).Value = True
        End Select
        
        Select Case .RstDiv
            Case "R": optRstDiv(0).Value = True
            Case "A": optRstDiv(1).Value = True
        End Select
        
        If .DetailFg = "*" Then
            chkDetailFg.Visible = True
            chkDetailFg.Value = "1"
        Else
            chkDetailFg.Visible = False
            chkDetailFg.Value = "0"
        End If
    End With
    Set objFTestCds = Nothing
End Sub

Private Sub tabSpcList_Click()
    Dim objTSpcs As clsIISTSpcs     '지정검체 컬렉션
    Dim objTSpc  As clsIISTSpc      '지정검체 클래스
    Dim strSpcCd As String          '검체코드
    
    '## 현재 검사코드, 검체에 대한 적용일 리스트를 tabSpcDt에 표시
    strSpcCd = tabSpcList.SelectedItem.Tag
    Set objTSpcs = mTMaster.TSpcs
    
    tabSpcDt.Tabs.Clear
    Call CtlClear(ccTabSpcNm)
    For Each objTSpc In objTSpcs
        If mTestCd = objTSpc.TestCd And strSpcCd = objTSpc.SpcCd Then
            tabSpcDt.Tabs.Add , , Format$(objTSpc.Applydt, "####-##-##")
        End If
    Next
    Set objTSpc = Nothing
    Set objTSpcs = Nothing
    
    '## 지정검체 정보를 표시함수 콜
    '## 참고치 정보 표시함수 콜
    tabSpcDt.Tabs(1).Selected = True
    Call GetRefList
End Sub

Private Sub tabSpcDt_Click()
    Dim objTSpcs    As clsIISTSpcs      '지정검체 컬렉션
    Dim strSpcCd    As String           '검사코드
    Dim strApplyDt  As String           '적용일
    
    '## 현재 검사코드, 검체, 적용일(검체)에 대한 지정검체 정보를 표시
    strSpcCd = tabSpcList.SelectedItem.Tag
    strApplyDt = Format$(tabSpcDt.SelectedItem.Caption, "YYYYMMDD")
    Call CtlClear(ccTabSpcDt)
    Set objTSpcs = mTMaster.TSpcs
    
    With objTSpcs(mTestCd, strSpcCd, strApplyDt)
        lblSpcCd.Caption = strSpcCd
        lblSpcNm.Caption = .SpcNm
        lblSpcApplyDt.Caption = Format$(.Applydt, "####-##-##")
        lblSpcExpDt.Caption = Format$(.ExpireDt, "####-##-##")
        lblSpcUnit.Caption = .Unit
        lblAvalVal.Caption = CStr(.AvalVal)
        
        If .PanicFg = "1" Then
            lblPanicFr.Caption = CStr(.PanicFrVal)
            lblPanicTo.Caption = CStr(.PanicToVal)
        End If
        
        If .DeltaFg = "1" Then
            lblDeltaFr.Caption = CStr(.DeltaFrVal)
            lblDeltaTo.Caption = CStr(.DeltaToVal)
        End If
    End With
    Set objTSpcs = Nothing
End Sub

Private Sub tabRefList_Click()
    Dim itmX        As ListItem
    Dim objRefs     As clsIISRefs   '참고치 컬렉션
    Dim objRef      As clsIISRef    '참고치 클래스
    Dim strSpcCd    As String       '검체코드
    Dim strApplyDt  As String       '적용일
    Dim strSex      As String       '적용성별
    
    Me.MousePointer = vbHourglass
    
    '## 현재 검사코드, 검체, 적용일(참고치)에 대한 참고치 정보를 표시
    strSpcCd = tabSpcList.SelectedItem.Tag
    strApplyDt = Format$(tabRefList.SelectedItem.Caption, "YYYYMMDD")
    
    lvwRefList.ListItems.Clear
    Set objRefs = mTMaster.Refs
    
    For Each objRef In objRefs
        With objRef
            If mTestCd = .TestCd And strSpcCd = .SpcCd And strApplyDt = .Applydt Then
                Select Case .Sex
                    Case "M": strSex = "남자"
                    Case "F": strSex = "여자"
                    Case "B": strSex = "Both"
                    Case "U": strSex = "중성"
                End Select
                
                Set itmX = lvwRefList.ListItems.Add(, , strSex)
                itmX.SubItems(1) = CStr(.AgeFr) & " - " & CStr(.AgeTo)
                itmX.SubItems(2) = CStr(.RefFrVal) & " - " & CStr(.RefToVal)
            End If
        End With
    Next
    
    If lvwRefList.ListItems.Count > 8 Then
        lvwRefList.ColumnHeaders(2).Width = 2080
    Else
        lvwRefList.ColumnHeaders(2).Width = 2300
    End If
    
    Set objRef = Nothing
    Set objRefs = Nothing
    Set itmX = Nothing
    
    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Workarea 코드, 코드명을 cboWorkarea에 로드
'-----------------------------------------------------------------------------'
Private Sub GetWorkarea()
    Dim objWA   As clsIISWorkarea       'Workarea 클래스
    Dim Rs      As ADODB.Recordset
    
On Error GoTo Errors
    Set objWA = New clsIISWorkarea
    Set Rs = objWA.GetWorkarea
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    With cboWorkarea
        Do Until Rs.EOF
            .AddItem Rs.Fields("WACD").Value & Space(5) & Rs.Fields("WAENGNM").Value
            Rs.MoveNext
        Loop
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Set objWA = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Set objWA = Nothing
    Error.SetLog App.EXEName, "frmIIS602", "GetWorkarea", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검사코드의 적용일 리스트 tabTestCd에 표시
'-----------------------------------------------------------------------------'
Private Sub GetTestCds()
    Dim objFTestCds As clsIISTestCdFulls    '검사코드 컬렉션
    Dim objFTestCd  As clsIISTestCdFull     '검사코드 클래스
    
    Me.MousePointer = vbHourglass
    
    '## 검사코드의 적용일 리스트를 표시
    tabTestCd.Tabs.Clear
    Set objFTestCds = mTMaster.GetTestCdInfo(mTestCd)
    If objFTestCds.Count = 0 Then
        Set objFTestCds = Nothing
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    For Each objFTestCd In objFTestCds
        tabTestCd.Tabs.Add , , Format$(objFTestCd.Applydt, "####-##-##")
    Next
    
    Set objFTestCds = Nothing
    Set objFTestCd = Nothing
    
    '## 검사코드+적용일에 해당하는 검사코드정보 표시함수 콜
    '## 검사코드에 대한 검체리스트 표시함수 콜
    tabTestCd.Tabs(1).Selected = True
    Call GetSpcNms
    
    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 현재 검사코드의 검체명을 tabSpcList에 표시
'-----------------------------------------------------------------------------'
Private Sub GetSpcNms()
    Dim tabX     As MSComctlLib.Tab
    Dim objTSpcs As clsIISTSpcs     '지정검체 컬렉션
    Dim objTSpc  As clsIISTSpc      '지정검체 클래스
    Dim strSpcCd As String          '검체코드
    
    '## 검체명 리스트 표시
    Me.MousePointer = vbHourglass
    
    tabSpcList.Tabs.Clear
    Set objTSpcs = mTMaster.GetSpcInfo(mTestCd)
    If objTSpcs.Count = 0 Then
        Set objTSpcs = Nothing
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    For Each objTSpc In objTSpcs
        If strSpcCd <> objTSpc.SpcCd Then
            strSpcCd = objTSpc.SpcCd
            Set tabX = tabSpcList.Tabs.Add(, , objTSpc.SpcNm)
            tabX.Tag = strSpcCd
        End If
    Next
    
    Set tabX = Nothing
    Set objTSpc = Nothing
    Set objTSpcs = Nothing
    
    '## 해당검체의 적용일 리스트 표시
    tabSpcList.Tabs(1).Selected = True
    
    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 현재 검사코드, 검체에 대한 참고치 적용일을 tabRefList에 표시
'-----------------------------------------------------------------------------'
Private Sub GetRefList()
    Dim objRefs     As clsIISRefs       '참고치 컬렉션
    Dim objRef      As clsIISRef        '참고치 클래스
    Dim strSpcCd    As String           '검체코드
    Dim strApplyDt  As String           '적용일
    
    Me.MousePointer = vbHourglass
    
    strSpcCd = tabSpcList.SelectedItem.Tag
    tabRefList.Tabs.Clear
    Set objRefs = mTMaster.GetRefList(mTestCd, strSpcCd)
    If objRefs.Count = 0 Then
        Set objRefs = Nothing
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    For Each objRef In objRefs
        If strApplyDt <> objRef.Applydt Then
            strApplyDt = objRef.Applydt
            tabRefList.Tabs.Add , , Format$(strApplyDt, "####-##-##")
        End If
    Next
    
    Set objRef = Nothing
    Set objRefs = Nothing
    tabRefList.Tabs(1).Selected = True
    
    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 Lock, Enable 유무결정, 저장/수정,삭제 버튼의 활성,비활성 결정
'   인수 :
'       1.pState : StateEnum 상수
'-----------------------------------------------------------------------------'
Private Sub CtlLock(ByVal pState As StateEnum)
    Dim blnEnable   As Boolean      'DTP Picker, PictureBox의 Enable 유무
    Dim blnLock     As Boolean      '이외의 다른 컨트롤의 Locked 유무
    
    Select Case pState
        Case StateEnum.ccInit                       '## 최초상태
            blnEnable = False
            blnLock = True
            cmdSave.Enabled = False
            cmdAdd.Enabled = False
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
            cmdCancel.Enabled = False
            dtpApplyDt.Enabled = False
            chkDetailFg.Visible = False
            mState = ccInit
        Case StateEnum.ccSave, StateEnum.ccAdd      '## 저장할수 있는 상태
            blnEnable = True
            blnLock = False
            cmdSave.Enabled = True
            cmdAdd.Enabled = False
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
            cmdCancel.Enabled = True
            chkDetailFg.Visible = True
            If pState = ccSave Then
                mState = ccModify
                dtpApplyDt.Enabled = False
            Else
                mState = ccAdd
                dtpApplyDt.Enabled = True
            End If
        Case StateEnum.ccModify                     '## 수정, 삭제할수 있는상태
            blnEnable = False
            blnLock = True
            cmdSave.Enabled = False
            cmdAdd.Enabled = True
            cmdModify.Enabled = True
            cmdDelete.Enabled = True
            cmdCancel.Enabled = False
            dtpApplyDt.Enabled = False
            mState = ccModify
    End Select
    
    txtTestCd.Locked = Not (blnLock)
    tabTestCd.Enabled = blnLock
    dtpExpDt.Enabled = blnEnable
    txtTestNm5.Locked = blnLock
    txtTestNm10.Locked = blnLock
    txtTestNm.Locked = blnLock
    cboWorkarea.Locked = blnLock
    picPanelFg.Enabled = blnEnable
    picRstDiv.Enabled = blnEnable
    cboRstType.Locked = blnLock
    txtItemSeq.Locked = blnLock
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 화면 컨트롤의 초기화
'   인수 :
'       1.pFlag : ClearEnum 상수
'-----------------------------------------------------------------------------'
Private Sub CtlClear(ByVal pFlag As ClearEnum)
    Select Case pFlag
        Case ClearEnum.ccAll
            txtTestCd.Text = "":            txtHopTestCd.Text = ""
            tabTestCd.Tabs.Clear:           tabSpcList.Tabs.Clear
            tabSpcDt.Tabs.Clear:            tabRefList.Tabs.Clear
            dtpApplyDt.Value = Now:         dtpExpDt.Value = ""
            txtTestNm5.Text = "":           txtTestNm10.Text = ""
            txtTestNm.Text = "":            txtItemSeq.Text = ""
            cboWorkarea.ListIndex = -1:     cboRstType.ListIndex = -1
            lblSpcCd.Caption = "":          lblSpcNm.Caption = ""
            lblSpcApplyDt.Caption = "":     lblSpcExpDt.Caption = ""
            lblSpcUnit.Caption = "":        lblAvalVal.Caption = ""
            lblPanicFr.Caption = "":        lblPanicTo.Caption = ""
            lblDeltaFr.Caption = "":        lblDeltaTo.Caption = ""
            optPanelFg(0).Value = True:     optPanelFg(1).Value = False
            optPanelFg(2).Value = False:    optRstDiv(0).Value = True
            optRstDiv(1).Value = False:     lvwRefList.ListItems.Clear
            chkDetailFg.Value = "0":        chkDetailFg.Visible = False
        
        Case ClearEnum.ccCmdAdd
            dtpApplyDt.Value = Now:         dtpExpDt.Value = ""
            txtTestNm5.Text = "":           txtTestNm10.Text = ""
            txtTestNm.Text = "":            txtItemSeq.Text = ""
            cboWorkarea.ListIndex = -1:     cboRstType.ListIndex = -1
            optPanelFg(0).Value = True:     optRstDiv(0).Value = True
            chkDetailFg.Value = "0":        chkDetailFg.Visible = False
        
        Case ClearEnum.ccTabSpcNm, ClearEnum.ccTabSpcDt
            lblSpcCd.Caption = "":          lblSpcNm.Caption = ""
            lblSpcApplyDt.Caption = "":     lblSpcExpDt.Caption = ""
            lblSpcUnit.Caption = "":        lblAvalVal.Caption = ""
            lblPanicFr.Caption = "":        lblPanicTo.Caption = ""
            lblDeltaFr.Caption = "":        lblDeltaTo.Caption = ""
            lvwRefList.ListItems.Clear
    End Select
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리1
'-----------------------------------------------------------------------------'
Private Sub mCode_SelectedItem(ByRef pSelItem As String)
    txtTestCd.Text = mGetP(pSelItem, 1, DIV)
End Sub

