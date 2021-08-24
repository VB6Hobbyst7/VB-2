VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIIS603 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "지정검체 관리"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7845
      Left            =   5865
      TabIndex        =   32
      Top             =   765
      Width           =   5265
      Begin VB.CheckBox chkDeltaFg 
         BackColor       =   &H00DBE6E6&
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   5220
         Width           =   180
      End
      Begin VB.CheckBox chkPanicFg 
         BackColor       =   &H00DBE6E6&
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   4680
         Width           =   180
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "추 가(&A)"
         Height          =   495
         Left            =   1140
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   690
         Width           =   990
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00DBE6E6&
         Caption         =   "수 정(&M)"
         Height          =   495
         Left            =   2130
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   690
         Width           =   990
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "삭 제(&D)"
         Height          =   495
         Left            =   3120
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   690
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "취 소(&C)"
         Height          =   495
         Left            =   4110
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   690
         Width           =   990
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저 장(&S)"
         Height          =   495
         Left            =   150
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   690
         Width           =   990
      End
      Begin VB.TextBox txtDeltaTo 
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3870
         MaxLength       =   20
         TabIndex        =   16
         Top             =   5175
         Width           =   1050
      End
      Begin VB.TextBox txtDeltaFr 
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   15
         Top             =   5175
         Width           =   1050
      End
      Begin VB.TextBox txtPanicTo 
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3675
         MaxLength       =   20
         TabIndex        =   13
         Top             =   4635
         Width           =   1245
      End
      Begin VB.ComboBox cboUnit 
         BackColor       =   &H00F7FFF7&
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
         ItemData        =   "frmIIS603.frx":0000
         Left            =   1845
         List            =   "frmIIS603.frx":0019
         TabIndex        =   9
         Top             =   3600
         Width           =   1680
      End
      Begin VB.TextBox txtPanicFr 
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1845
         MaxLength       =   20
         TabIndex        =   12
         Top             =   4635
         Width           =   1245
      End
      Begin VB.TextBox txtAvalVal 
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1845
         MaxLength       =   1
         TabIndex        =   10
         Top             =   4125
         Width           =   645
      End
      Begin MedControls1.LisLabel lblSpcCd 
         Height          =   345
         Left            =   1845
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   1845
         TabIndex        =   6
         Top             =   1965
         Width           =   2820
         _ExtentX        =   4974
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
      Begin MSComCtl2.DTPicker dtpApplyDt 
         Height          =   330
         Left            =   1845
         TabIndex        =   7
         Top             =   2520
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         Format          =   23724033
         CurrentDate     =   37994
      End
      Begin MSComCtl2.DTPicker dtpExpireDt 
         Height          =   330
         Left            =   1845
         TabIndex        =   8
         Top             =   3060
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   23724033
         CurrentDate     =   37994
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   2475
         TabIndex        =   48
         Top             =   4095
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "(9: 적용안함)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2775
         TabIndex        =   52
         Top             =   4185
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4935
         TabIndex        =   47
         Top             =   5250
         Width           =   150
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3105
         TabIndex        =   46
         Top             =   5250
         Width           =   150
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "(-)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3615
         TabIndex        =   45
         Top             =   5250
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "(+)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1785
         TabIndex        =   44
         Top             =   5250
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3315
         TabIndex        =   43
         Top             =   4710
         Width           =   135
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "폐기일 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   330
         TabIndex        =   41
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "적용일 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   330
         TabIndex        =   40
         Top             =   2595
         Width           =   660
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체명 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   330
         TabIndex        =   39
         Top             =   2055
         Width           =   660
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체코드 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   330
         TabIndex        =   38
         Top             =   1530
         Width           =   840
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Delta Check :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   330
         TabIndex        =   37
         Top             =   5250
         Width           =   1140
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Panic Check :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   330
         TabIndex        =   36
         Top             =   4710
         Width           =   1200
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "유효숫자 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   330
         TabIndex        =   35
         Top             =   4185
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "결과단위 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   330
         TabIndex        =   34
         Top             =   3660
         Width           =   840
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "상 세 정 보"
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
         Left            =   1560
         TabIndex        =   33
         Top             =   285
         Width           =   1035
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   285
         Top             =   180
         Width           =   3495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   45
      TabIndex        =   30
      Top             =   4560
      Width           =   5775
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00DBE6E6&
         Caption         =   "참고치수정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4410
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   420
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwRefList 
         Height          =   2625
         Left            =   60
         TabIndex        =   22
         Top             =   1320
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   4630
         View            =   3
         LabelEdit       =   1
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
            Object.Width           =   3898
         EndProperty
      End
      Begin MSComctlLib.TabStrip tabRefList 
         Height          =   315
         Left            =   105
         TabIndex        =   50
         Top             =   945
         Width           =   5505
         _ExtentX        =   9710
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
      Begin VB.Shape Shape4 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00DBE6E6&
         FillStyle       =   0  '단색
         Height          =   390
         Left            =   90
         Top             =   915
         Width           =   5550
      End
      Begin VB.Label lblCd 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   120
         TabIndex        =   49
         Top             =   660
         Width           =   180
      End
      Begin VB.Label lblTestSpcNm 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "WBC Count - B(EDTA)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   360
         TabIndex        =   42
         Top             =   630
         Width           =   2190
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참 고 치"
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
         Left            =   1500
         TabIndex        =   31
         Top             =   285
         Width           =   765
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   75
         Top             =   180
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   45
      TabIndex        =   26
      Top             =   0
      Width           =   11085
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00DBE6E6&
         Caption         =   "다음(&N) >>"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8655
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00DBE6E6&
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
         Height          =   495
         Left            =   9780
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   180
         Width           =   1125
      End
      Begin VB.TextBox txtTestCd 
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   0
         Top             =   285
         Width           =   1575
      End
      Begin VB.CommandButton cmdTestSrh 
         BackColor       =   &H00DBE6E6&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2925
         Picture         =   "frmIIS603.frx":0048
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   270
         Width           =   405
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H00DBE6E6&
         Caption         =   "<< 이전(&P)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7530
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   180
         Width           =   1125
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   345
         Left            =   3405
         TabIndex        =   51
         Top             =   270
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   609
         BackColor       =   16252919
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "WBC Count"
         LeftGab         =   100
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
         TabIndex        =   27
         Top             =   345
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Left            =   45
      TabIndex        =   28
      Top             =   765
      Width           =   5775
      Begin VB.CommandButton cmdSpcAdd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체추가"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4410
         Style           =   1  '그래픽
         TabIndex        =   24
         Top             =   150
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwSpcList 
         Height          =   3045
         Left            =   75
         TabIndex        =   1
         Top             =   615
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   5371
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16252919
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "검체코드"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "검체명"
            Object.Width           =   3193
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "적용일"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "폐기일"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 체 리 스 트"
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
         Left            =   1230
         TabIndex        =   29
         Top             =   285
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   375
         Left            =   75
         Top             =   180
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmIIS603"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS603.frm
'   작성자  : 이상대
'   내  용  : 지정검체 마스터
'   작성일  : 2004-01-09
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private Enum ClearEnum
    ccAll           '전체 Clear
    ccLvwSpcList    'lvwSpcList 클릭시 컨트롤 Clear
    ccCmdAdd        'cmdAdd 클릭시 컨트롤 Clear
End Enum

Private Enum StateEnum
    ccInit          '초기상태
    ccAdd           '새 적용일을 추가하는 상태
    ccSave          '저장버튼을 누를수 있는 상태
    ccModify        '수정하는 상태
End Enum

Private mTestCd     As String           '검사코드
Private mState      As StateEnum        '현재폼의 상태
Private mTMaster    As clsIISTMaster    '검사코드 마스터 클래스

Private WithEvents mCode1 As clsIISCodeList      'CodeList 클래스
Attribute mCode1.VB_VarHelpID = -1
Private WithEvents mCode2 As clsIISCodeList      'CodeList 클래스
Attribute mCode2.VB_VarHelpID = -1

Public Property Let TestCd(ByVal pTestCd As String)
    mTestCd = pTestCd
End Property

Public Property Let TMaster(ByVal pTMaster As clsIISTMaster)
    Set mTMaster = pTMaster
End Property

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight: .Width = 11270
    End With
    
    '## 1.검사코드 화면에서 지정검체등록 버튼을 클릭한경우
    '   - mTMaster <> Nothing
    '   - mTestCd <> ""
    '## 2.바로 지정검체화면으로 들어온경우
    '   - mTMaster = Nothing
    '   - mTestCd = ""
    Call CtlClear(ccAll)
    Call CtlLock(ccInit)
    Me.Show
    DoEvents
    
    Me.MousePointer = vbHourglass
    
    If mTMaster Is Nothing And mTestCd = "" Then
        Set mTMaster = New clsIISTMaster
        Call mTMaster.GetTestCdList
    Else
        txtTestCd.Text = mTestCd
        Call txtTestCd_KeyDown(vbKeyReturn, 0)
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS603").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTMaster = Nothing
    Set frmIIS603 = Nothing
End Sub

Private Sub cmdTestSrh_Click()
    Set mCode1 = New clsIISCodeList
    With mCode1
        .Caption = "검사코드 리스트"
        .HeaderCd = "검사코드"
        .HeaderCdNm = "검사명"
        .CodeListByCol mTMaster.TestCds
    End With
    Set mCode1 = Nothing
    
    Call txtTestCd_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdSpcAdd_Click()
    Dim objSpc As clsIISSpc

    Set mCode2 = New clsIISCodeList
    Set objSpc = New clsIISSpc
    With mCode2
        .Caption = "검체 리스트"
        .HeaderCd = "검체코드"
        .HeaderCdNm = "검체명"
        .CodeListByRs objSpc.GetSpcCd
    End With

    Set objSpc = Nothing
    Set mCode2 = Nothing
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
    lblTestNm.Caption = mTMaster.GetTestNm(mTestCd)
    Call GetSpcList
    Call CtlLock(ccModify)
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
    lblTestNm.Caption = mTMaster.GetTestNm(mTestCd)
    Call GetSpcList
    Call CtlLock(ccModify)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSpcCd        As String       '검체코드
    Dim strApplyDt      As String       '적용일
    Dim strExpireDt     As String       '폐기일
    Dim strUnit         As String       '결과단위
    Dim lngAvalVal      As Long         '유효숫자
    Dim strPanicFg      As String       'Panic Check(0:No, 1:Yes)
    Dim sngPanicFrVal   As Single       'Panic To Value
    Dim sngPanicToVal   As Single       'Panic From Value
    Dim strDeltaFg      As String       'Delta Check(0:No, 1:Yes)
    Dim lngDeltaFrVal   As Long         'Delta To Value
    Dim lngDeltaToVal   As Long         'Delta From Value
    Dim strLastDt       As String       '적용일중 최근 적용일
    Dim blnReturn       As Boolean

    '## 1. 신규 검체코드 입력
    '   - mState=ccAdd, mTMaster.ExistSpcCD() = False
    '   - 적용일 체크 불필요
    '   - Insert
    '## 2. 기존 검체코드에 적용일 추가
    '   - mState=ccAdd, mTMaster.ExistSpcCd() = True
    '   - 적용일 체크
    '   - Insert
    '## 3. 기존 검체코드 수정
    '   - mState=ccModify, mTMaster.ExistSpcCD() = True
    '   - 적용일 체크 불필요
    '   - Update
    
    strSpcCd = lblSpcCd.Caption
    strApplyDt = Format$(dtpApplyDt.Value, "YYYYMMDD")
    
    '## 적용일 Check
    '   - 검체코드의 적용일 가장 최근 적용일 보다 커야함
    '   - 적용일이 오늘날짜여도 상관없음
    If mTMaster.ExistSpcCd(strSpcCd) And mState = ccAdd Then
        strLastDt = mTMaster.GetSpcCdLastApplyDt(strSpcCd)
        If strApplyDt <= strLastDt Then
            MsgBox "적용일은 이전 적용일 보다 커야 합니다.", vbInformation, "정보"
            Exit Sub
        End If
    End If
    
    strExpireDt = Format$(dtpExpireDt.Value, "YYYYMMDD")
    strUnit = cboUnit.Text
    strPanicFg = chkPanicFg.Value
    strDeltaFg = chkDeltaFg.Value
    
    lngAvalVal = IIf(txtAvalVal.Text = "", 9, CLng(txtAvalVal.Text))
    sngPanicFrVal = CSng(txtPanicFr.Text)
    sngPanicToVal = CSng(txtPanicTo.Text)
    lngDeltaFrVal = CLng(txtDeltaFr.Text)
    lngDeltaToVal = CLng(txtDeltaTo.Text)
    
    '## DB에 저장
    If mState = ccModify Then
        '## Update
        blnReturn = mTMaster.ModifySpcCd(mTestCd, strSpcCd, strApplyDt, strExpireDt, strUnit, _
            lngAvalVal, strPanicFg, sngPanicFrVal, sngPanicToVal, strDeltaFg, lngDeltaFrVal, lngDeltaToVal)
    Else
        '## Insert
        blnReturn = mTMaster.AddSpcCd(mTestCd, strSpcCd, strApplyDt, strExpireDt, strUnit, _
            lngAvalVal, strPanicFg, sngPanicFrVal, sngPanicToVal, strDeltaFg, lngDeltaFrVal, lngDeltaToVal)
    End If
    
    If blnReturn Then
        Call GetSpcList
        mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
    Else
        mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
    End If
End Sub

Private Sub cmdAdd_Click()
    Call CtlClear(ccCmdAdd)
    Call CtlLock(ccAdd)
    dtpExpireDt.SetFocus
End Sub

Private Sub cmdModify_Click()
    Call CtlLock(ccSave)
    dtpExpireDt.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim strSpcCd    As String       '검체코드
    Dim strApplyDt  As String       '적용일
    Dim intTemp     As Integer
    
    intTemp = MsgBox("정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub
    
    strSpcCd = lblSpcCd.Caption
    strApplyDt = Format$(dtpApplyDt.Value, "YYYYMMDD")
    
    If mTMaster.RemoveSpcCd(mTestCd, strSpcCd, strApplyDt) Then
        Call GetSpcList
        mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
    Else
        mdiIISMain.sbrStatus.Panels(2).Text = "삭제중 에러가 발생했습니다."
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim intTemp As Integer
    
    intTemp = MsgBox("변경된 내용을 취소할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub
    
    Call lvwSpcList_ItemClick(lvwSpcList.SelectedItem)
    Call CtlLock(ccModify)
    lvwSpcList.SetFocus
End Sub

Private Sub cmdRef_Click()
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
            '## 신규/기존 검사코드를 판단하여 기존코드이면 정보표시
            '## 신규코드이면 모든 컨트롤 Lock
            mTestCd = Trim(txtTestCd.Text)
            
            Call CtlClear(ccAll)
            Call CtlLock(ccInit)
            If mTestCd = "" Then Exit Sub
            
            If mTMaster.Exist(mTestCd) Then
                txtTestCd.Text = mTestCd
                lblTestNm.Caption = mTMaster.GetTestNm(mTestCd)
                Call GetSpcList
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

Private Sub txtAvalVal_GotFocus()
    With txtAvalVal
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAvalVal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtAvalVal_KeyPress(KeyAscii As Integer)
    If CheckNum(KeyAscii) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPanicFr_GotFocus()
    With txtPanicFr
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPanicFr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPanicFr_KeyPress(KeyAscii As Integer)
    If CheckNum(KeyAscii) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPanicTo_GotFocus()
    With txtPanicTo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPanicTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPanicTo_KeyPress(KeyAscii As Integer)
    If CheckNum(KeyAscii) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDeltaTo_GotFocus()
    With txtDeltaTo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDeltaTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDeltaTo_KeyPress(KeyAscii As Integer)
    If CheckNum(KeyAscii) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDeltaFr_GotFocus()
    With txtDeltaFr
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDeltaFr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDeltaFr_KeyPress(KeyAscii As Integer)
    If CheckNum(KeyAscii) = False Then
        KeyAscii = 0
    End If
End Sub

Private Sub lvwSpcList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim objTSpcs    As clsIISTSpcs
    Dim strSpcCd    As String       '검체코드
    Dim strApplyDt  As String       '적용일
    Dim strOldSpcCd As String       '이전 검체코드
    
    '## 해당 검사코드+검체코드+적용일에 대한 정보표시
    Call CtlClear(ccLvwSpcList)
    strSpcCd = Item.Text
    strApplyDt = Format$(Item.SubItems(2), "YYYYMMDD")
    strOldSpcCd = lblSpcCd.Caption
    
    Set objTSpcs = mTMaster.TSpcs
    With objTSpcs(mTestCd, strSpcCd, strApplyDt)
        lblSpcCd.Caption = strSpcCd
        lblSpcNm.Caption = .SpcNm
        dtpApplyDt.Value = Format$(.Applydt, "####-##-##")
        dtpExpireDt.Value = Format$(.ExpireDt, "####-##-##")
        cboUnit.Text = .Unit
        txtAvalVal.Text = CStr(.AvalVal)
        
        If .PanicFg = "1" Then
            chkPanicFg.Value = "1"
            txtPanicFr.Text = CStr(.PanicFrVal)
            txtPanicTo.Text = CStr(.PanicToVal)
        End If
        
        If .DeltaFg = "1" Then
            chkDeltaFg.Value = "1"
            txtDeltaFr.Text = CStr(.DeltaFrVal)
            txtDeltaTo.Text = CStr(.DeltaToVal)
        End If
        
        Call CtlLock(ccModify)
    End With
    Set objTSpcs = Nothing
    
    '## 참고치 정보 표시
    lblTestSpcNm.Caption = lblSpcNm.Caption & " - " & strSpcCd
    Call GetRefList
End Sub

Private Sub tabRefList_Click()
    Dim itmX        As ListItem
    Dim objRefs     As clsIISRefs   '참고치 컬렉션
    Dim objRef      As clsIISRef    '참고치 클래스
    Dim strSpcCd    As String       '검체코드
    Dim strApplyDt  As String       '적용일
    Dim strSex      As String       '적용성별
    
    '## 현재 검사코드, 검체, 적용일(참고치)에 대한 참고치 정보를 표시
    strSpcCd = lblSpcCd.Caption
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
    
    If lvwRefList.ListItems.Count > 11 Then
        lvwRefList.ColumnHeaders(2).Width = 2050
    Else
        lvwRefList.ColumnHeaders(2).Width = 2300
    End If
    
    Set objRef = Nothing
    Set objRefs = Nothing
    Set itmX = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검체리스트를 lvwSpcList에 표시
'-----------------------------------------------------------------------------'
Private Sub GetSpcList()
    Dim objTSpcs    As clsIISTSpcs     '지정검체 컬렉션
    Dim objTSpc     As clsIISTSpc      '지정검체 클래스
    Dim itmX        As ListItem
    
    '## 반드시 mTestCd <> "" 이어야 하고, 기존검사코드 이어야 한다.
    lvwSpcList.ListItems.Clear
    Set objTSpcs = mTMaster.GetSpcInfo(mTestCd)
    If objTSpcs.Count = 0 Then
        Set objTSpcs = Nothing
        Exit Sub
    End If
    
    For Each objTSpc In objTSpcs
        Set itmX = lvwSpcList.ListItems.Add(, , objTSpc.SpcCd)
        itmX.SubItems(1) = objTSpc.SpcNm
        itmX.SubItems(2) = Format$(objTSpc.Applydt, "####-##-##")
        itmX.SubItems(3) = Format$(objTSpc.ExpireDt, "####-##-##")
    Next
    
    If lvwSpcList.ListItems.Count > 14 Then
        lvwSpcList.ColumnHeaders(2).Width = 1590
    Else
        lvwSpcList.ColumnHeaders(2).Width = 1810
    End If
        
    Set itmX = Nothing
    Set objTSpc = Nothing
    Set objTSpcs = Nothing
    
    '## 검체정보 표시
    Call lvwSpcList_ItemClick(lvwSpcList.SelectedItem)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 참고치의 적용일 리스트를 tabRefList에 표시
'-----------------------------------------------------------------------------'
Private Sub GetRefList()
    Dim objRefs     As clsIISRefs       '참고치 컬렉션
    Dim objRef      As clsIISRef        '참고치 클래스
    Dim strSpcCd    As String           '검체코드
    Dim strApplyDt  As String           '적용일
    
    strSpcCd = lblSpcCd.Caption
    tabRefList.Tabs.Clear
    Set objRefs = mTMaster.GetRefList(mTestCd, strSpcCd)
    If objRefs.Count = 0 Then
        Set objRefs = Nothing
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
    
    '## 참고치 정보표시
    tabRefList.Tabs(1).Selected = True
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 숫자, Backspace키만 입력되도록 함
'   인수 :
'       1.KeyAscii : 입력된 키의 ASCII코드값
'   반환 : True(숫자,Backspace키), False(숫자,Backspace이외의 키)
'-----------------------------------------------------------------------------'
Private Function CheckNum(KeyAscii As Integer) As Boolean
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        CheckNum = False
    Else
        CheckNum = True
    End If
End Function

'-----------------------------------------------------------------------------'
'   기능 : 현재 상태에 따라 컨트롤 Lock, Enable 유무결정
'   인수 :
'       1.pState : StateEnum 상수
'-----------------------------------------------------------------------------'
Private Sub CtlLock(ByVal pState As StateEnum)
    Dim blnLock     As Boolean      'Locked 속성
    Dim blnEnable   As Boolean      'Enabled 속성
    
    Select Case pState
        Case StateEnum.ccInit
            cmdSave.Enabled = False
            cmdAdd.Enabled = False
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
            cmdCancel.Enabled = False
            cmdSpcAdd.Enabled = True
            cmdRef.Enabled = False
            dtpApplyDt.Enabled = False
            blnLock = True
            blnEnable = False
        Case StateEnum.ccSave, StateEnum.ccAdd
            cmdSave.Enabled = True
            cmdAdd.Enabled = False
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
            cmdCancel.Enabled = True
            cmdSpcAdd.Enabled = True
            cmdRef.Enabled = True
            blnLock = False
            blnEnable = True
            If pState = ccSave Then
                mState = ccModify
                dtpApplyDt.Enabled = False
            Else
                mState = ccAdd
                dtpApplyDt.Enabled = True
            End If
        Case StateEnum.ccModify
            cmdSave.Enabled = False
            cmdAdd.Enabled = True
            cmdModify.Enabled = True
            cmdDelete.Enabled = True
            cmdCancel.Enabled = False
            cmdSpcAdd.Enabled = True
            cmdRef.Enabled = True
            blnLock = True
            blnEnable = False
    End Select
    
    txtTestCd.Locked = Not (blnLock)
    dtpExpireDt.Enabled = blnEnable
    cboUnit.Locked = blnLock
    txtAvalVal.Locked = blnLock
    txtPanicFr.Locked = blnLock
    txtPanicTo.Locked = blnLock
    txtDeltaFr.Locked = blnLock
    txtDeltaTo.Locked = blnLock
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 화면 컨트롤의 초기화
'   인수 :
'       1.pFlag : ClearEnum 상수
'-----------------------------------------------------------------------------'
Private Sub CtlClear(ByVal pFlag As ClearEnum)
    Select Case pFlag
        Case ClearEnum.ccAll
            txtTestCd.Text = "":            txtAvalVal.Text = ""
            lvwSpcList.ListItems.Clear:     lvwRefList.ListItems.Clear
            lblSpcCd.Caption = "":          lblSpcNm.Caption = ""
            lblTestNm.Caption = "":         lblTestSpcNm.Caption = ""
            dtpApplyDt.Value = Now:         dtpExpireDt.Value = ""
            txtPanicFr.Text = "":           txtPanicTo.Text = ""
            txtDeltaFr.Text = "":           txtDeltaTo.Text = ""
            cboUnit.Text = "":              tabRefList.Tabs.Clear
            chkPanicFg.Value = "0":         chkDeltaFg.Value = "0"
        Case ClearEnum.ccLvwSpcList
            lblSpcCd.Caption = "":          lblSpcNm.Caption = ""
            dtpApplyDt.Value = Now:         dtpExpireDt.Value = ""
            txtPanicFr.Text = "":           txtPanicTo.Text = ""
            txtDeltaFr.Text = "":           txtDeltaTo.Text = ""
            lvwRefList.ListItems.Clear:     txtAvalVal.Text = ""
            cboUnit.Text = "":              lblTestSpcNm.Caption = ""
            chkPanicFg.Value = "0":         chkDeltaFg.Value = "0"
            tabRefList.Tabs.Clear
        Case ClearEnum.ccCmdAdd
            dtpApplyDt.Value = Now:         dtpExpireDt.Value = ""
            txtPanicFr.Text = "":           txtPanicTo.Text = ""
            txtDeltaFr.Text = "":           txtDeltaTo.Text = ""
            txtAvalVal.Text = "":           cboUnit.Text = ""
            chkPanicFg.Value = "0":         chkDeltaFg.Value = "0"
    End Select
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리1
'-----------------------------------------------------------------------------'
Private Sub mCode1_SelectedItem(ByRef pSelItem As String)
    txtTestCd.Text = mGetP(pSelItem, 1, DIV)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리2
'-----------------------------------------------------------------------------'
Private Sub mCode2_SelectedItem(ByRef pSelItem As String)
    Dim itmX     As ListItem
    Dim strSpcCd As String      '검체코드
    
    strSpcCd = mGetP(pSelItem, 1, DIV)
    
    '## 검체명 구하기
    Set itmX = lvwSpcList.FindItem(strSpcCd)
    If Not (itmX Is Nothing) Then
        MsgBox "이미 존재하는 검체입니다.", vbInformation, "정보"
        pSelItem = ""
        Exit Sub
    End If
    Set itmX = Nothing
    
    lblSpcCd.Caption = strSpcCd
    lblSpcNm.Caption = mGetP(pSelItem, 2, DIV)
    
    Call CtlClear(ccCmdAdd)
    Call CtlLock(ccAdd)
End Sub
