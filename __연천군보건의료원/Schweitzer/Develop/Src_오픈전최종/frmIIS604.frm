VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIIS604 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "참고치 관리"
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
      TabIndex        =   28
      Top             =   765
      Width           =   5265
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "저 장(&S)"
         Height          =   495
         Left            =   150
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   690
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "취 소(&C)"
         Height          =   495
         Left            =   4110
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   690
         Width           =   990
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "삭 제(&D)"
         Height          =   495
         Left            =   3120
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   690
         Width           =   990
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00DBE6E6&
         Caption         =   "수 정(&M)"
         Height          =   495
         Left            =   2130
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   690
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "추 가(&A)"
         Height          =   495
         Left            =   1140
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   690
         Width           =   990
      End
      Begin VB.TextBox txtFrAge 
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
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   8
         Top             =   4110
         Width           =   1245
      End
      Begin VB.TextBox txtToAge 
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
         Left            =   3630
         MaxLength       =   20
         TabIndex        =   9
         Top             =   4110
         Width           =   1245
      End
      Begin VB.TextBox txtAlpha 
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
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   12
         Top             =   5160
         Width           =   1050
      End
      Begin VB.TextBox txtToRef 
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
         Left            =   3630
         MaxLength       =   20
         TabIndex        =   11
         Top             =   4635
         Width           =   1245
      End
      Begin VB.ComboBox cboSex 
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
         ItemData        =   "frmIIS604.frx":0000
         Left            =   1995
         List            =   "frmIIS604.frx":0010
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   3615
         Width           =   1680
      End
      Begin VB.TextBox txtFrRef 
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
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   10
         Top             =   4635
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpApplyDt 
         Height          =   330
         Left            =   1995
         TabIndex        =   5
         Top             =   2565
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         Format          =   68681729
         CurrentDate     =   37994
      End
      Begin MSComCtl2.DTPicker dtpExpireDt 
         Height          =   330
         Left            =   1995
         TabIndex        =   6
         Top             =   3090
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   68681729
         CurrentDate     =   37994
      End
      Begin MedControls1.LisLabel lblSpcCd 
         Height          =   345
         Left            =   1995
         TabIndex        =   3
         Top             =   1485
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
         Left            =   1995
         TabIndex        =   4
         Top             =   2025
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "3. M : 최대값(일령:50000, 연령:137)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   510
         TabIndex        =   45
         Tag             =   "35214"
         Top             =   6690
         Width           =   3150
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "2. Y : 연령으로"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   510
         TabIndex        =   44
         Tag             =   "35214"
         Top             =   6380
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "1. D : 입력된 값을 일령으로"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   510
         TabIndex        =   43
         Tag             =   "35214"
         Top             =   6070
         Width           =   2430
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "※ 나이계산 단축키"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   285
         TabIndex        =   42
         Tag             =   "35214"
         Top             =   5760
         Width           =   1770
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
         Left            =   285
         TabIndex        =   41
         Top             =   1575
         Width           =   840
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
         Left            =   285
         TabIndex        =   40
         Top             =   2100
         Width           =   660
      End
      Begin VB.Label Label5 
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
         Left            =   3345
         TabIndex        =   39
         Top             =   4185
         Width           =   135
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
         Left            =   3345
         TabIndex        =   36
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
         Left            =   285
         TabIndex        =   35
         Top             =   3135
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
         Left            =   285
         TabIndex        =   34
         Top             =   2625
         Width           =   660
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Alpha결과 참고치 :"
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
         Left            =   285
         TabIndex        =   33
         Top             =   5235
         Width           =   1560
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "참고치 :"
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
         Left            =   285
         TabIndex        =   32
         Top             =   4710
         Width           =   660
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "일 령 :"
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
         Left            =   285
         TabIndex        =   31
         Top             =   4185
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "성 별 :"
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
         Left            =   285
         TabIndex        =   30
         Top             =   3660
         Width           =   540
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
         TabIndex        =   29
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
      TabIndex        =   26
      Top             =   4560
      Width           =   5775
      Begin VB.CommandButton cmdAddRef 
         BackColor       =   &H00DBE6E6&
         Caption         =   "참고치추가"
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
         TabIndex        =   46
         Top             =   150
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwRefList 
         Height          =   2925
         Left            =   60
         TabIndex        =   2
         Top             =   1020
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   5159
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
            Text            =   "성별"
            Object.Width           =   1605
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "일령"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "연령"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "기준치"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComctlLib.TabStrip tabRefList 
         Height          =   315
         Left            =   105
         TabIndex        =   37
         Top             =   645
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
         Top             =   615
         Width           =   5550
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
         TabIndex        =   27
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
      TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         Picture         =   "frmIIS604.frx":002C
         Style           =   1  '그래픽
         TabIndex        =   22
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
         TabIndex        =   18
         Top             =   180
         Width           =   1125
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   345
         Left            =   3405
         TabIndex        =   38
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
         TabIndex        =   23
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
      TabIndex        =   24
      Top             =   765
      Width           =   5775
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "검체코드"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "검체명"
            Object.Width           =   7056
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
         TabIndex        =   25
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
Attribute VB_Name = "frmIIS604"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS603.frm
'   작성자  : 이상대
'   내  용  : 참고치 마스터
'   작성일  : 2004-01-09
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private Enum ClearEnum
    ccAll           '전체 Clear
    ccLvwSpcList    'lvwSpcList 클릭시 컨트롤 Clear
    ccTabRefList    'tabRefList 클릭시 컨트롤 Clear
    ccLvwRefList    'lvwRefList 클릭시 컨트롤 Clear
    ccCmdAdd        'cmdAdd 클릭시 컨트롤 Clear
End Enum

Private Enum StateEnum
    ccInit          '초기상태
    ccRefAdd        '새 적용일을 추가하는 상태
    ccAdd           '해당 적용일에 참고치 정보를 추가하는 상태
    ccSave          '저장버튼을 누를수 있는 상태
    ccModify        '수정하는 상태
End Enum

Private mTestCd  As String           '검사코드
Private mState   As StateEnum        '현재폼의 상태
Private mTMaster As clsIISTMaster    '검사코드 마스터 클래스

Private WithEvents mCode As clsIISCodeList      'CodeList 클래스
Attribute mCode.VB_VarHelpID = -1

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
    
    '## 1.검사코드,지정검체 화면에서 참고치등록 버튼을 클릭한경우
    '   - mTMaster <> Nothing
    '   - mTestCd <> ""
    '   - mSpcCd <> ""
    '## 2.바로 참고치등록 화면으로 들어온경우
    '   - mTMaster = Nothing
    '   - mTestCd = ""
    '   - mSpcCD = ""
    
    Call CtlClear(ccAll)
    Call CtlLock(ccInit)
    Me.Show
    DoEvents
    
    Me.MousePointer = vbHourglass
    If mTMaster Is Nothing Then
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
    frmIIS600.tvwMenu.Nodes("IIS604").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTMaster = Nothing
    Set frmIIS604 = Nothing
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

Private Sub cmdPrev_Click()
    Dim strTestCd As String     '검사코드

    strTestCd = UCase(Trim(txtTestCd.Text))
    If strTestCd = "" Then Exit Sub
    If mTMaster.Exist(strTestCd) = False Then Exit Sub

    mTestCd = mTMaster.PrevTestCd(strTestCd)
    If mTestCd = strTestCd Then Exit Sub

    Call CtlClear(ccAll)
    Call CtlLock(ccInit)
    txtTestCd.Text = mTestCd
    lblTestNm.Caption = mTMaster.GetTestNm(mTestCd)
    Call GetSpcList
End Sub

Private Sub cmdNext_Click()
    Dim strTestCd As String     '검사코드

    strTestCd = UCase(Trim(txtTestCd.Text))
    If strTestCd = "" Then Exit Sub
    If mTMaster.Exist(strTestCd) = False Then Exit Sub

    mTestCd = mTMaster.NextTestCd(strTestCd)
    If mTestCd = strTestCd Then Exit Sub

    Call CtlClear(ccAll)
    Call CtlLock(ccInit)
    txtTestCd.Text = mTestCd
    lblTestNm.Caption = mTMaster.GetTestNm(mTestCd)
    Call GetSpcList
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdAddRef_Click()
    If lblSpcCd.Caption = "" Then Exit Sub
    
    mState = ccRefAdd
    Call CtlClear(ccLvwRefList)
    Call CtlLock(ccSave)
    dtpApplyDt.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim strSpcCd    As String       '검체코드
    Dim strApplyDt  As String       '적용일
    Dim strExpireDt As String       '폐기일
    Dim strSex      As String       '성별
    Dim lngAgeFr    As Long         'From Age
    Dim lngAgeTo    As Long         'To Age
    Dim sngRefFrVal As Single       'From Reference
    Dim sngRefToVal As Single       'To Reference
    Dim strRefCd    As String       'Alpha결과 참고치
    Dim strLastDt   As String       '참고치 적용일중 가장 최근적용일
    Dim blnReturn   As Boolean
    
    '## 1. 기존 참고치정보가 없는상태에서 적용일 생성
    '   - mState=ccRefAdd, mTMaster.ExistRef()=false
    '   - 적용일 체크 불필요
    '   - Insert
    '## 2. 기존 참고치정보가 있고, 새 적용일 생성
    '   - mState=ccRefAdd, mTMaster.ExistRef()=true
    '   - 적용일 체크
    '   - Insert
    '## 3. 기존 참고치정보가 있고, 적용일에 새 참고치정보를 입력
    '   - mState=ccAdd, mTMaster.ExistRef()=true
    '   - 적용일 체크 불필요
    '   - Insert
    '## 4. 기존 참고치정보가 있고, 기존 참고치 정보를 수정
    '   - mState=ccModify, mTMaster.Exist()=true
    '   - 적용일 체크 불필요
    '   - Update
    
    strSpcCd = lblSpcCd.Caption
    strApplyDt = Format$(dtpApplyDt.Value, "YYYYMMDD")
    
    '## 적용일 Check
    If mTMaster.ExistRef(strSpcCd) And mState = ccRefAdd Then
        strLastDt = mTMaster.GetRefLastApplyDt(strSpcCd)
        If strApplyDt <= strLastDt Then
            MsgBox "적용일은 이전 적용일 보다 커야 합니다.", vbInformation, "정보"
            Exit Sub
        End If
    End If
    strExpireDt = Format$(dtpExpireDt.Value, "YYYYMMDD")
    
    '## 성별 Check
    If cboSex.ListIndex = -1 Then
        MsgBox "성별을 입력해주세요.", vbInformation, "정보"
        Exit Sub
    End If
    Select Case cboSex.Text
        Case "남자": strSex = "M"
        Case "여자": strSex = "F"
        Case "Both": strSex = "B"
        Case "중성": strSex = "U"
    End Select
    
    '## 일령 Check
    If txtFrAge.Text = "" Or txtToAge.Text = "" Then
        MsgBox "일령을 입력해 주세요.", vbInformation, "정보"
        Exit Sub
    End If
    lngAgeFr = CLng(txtFrAge.Text)
    lngAgeTo = CLng(txtToAge.Text)
    
    '## 참고치 Check
    If txtFrRef.Text = "" Or txtToRef.Text = "" Then
        MsgBox "참고치를 입력해주세요.", vbInformation, "정보"
        Exit Sub
    End If
    sngRefFrVal = CSng(txtFrRef.Text)
    sngRefToVal = CSng(txtToRef.Text)
    strRefCd = Trim(txtAlpha.Text)
    
    '## DB에 저장
    If mState = ccModify Then
        '## Update
        blnReturn = mTMaster.ModifyRef(mTestCd, strSpcCd, strSex, lngAgeFr, lngAgeTo, strApplyDt, _
            strExpireDt, sngRefFrVal, sngRefToVal, strRefCd)
    Else
        '## Insert
        blnReturn = mTMaster.AddRef(mTestCd, strSpcCd, strSex, lngAgeFr, lngAgeTo, strApplyDt, _
            strExpireDt, sngRefFrVal, sngRefToVal, strRefCd)
    End If
    
    If blnReturn Then
        Call GetRefList
        mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
    Else
        mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
    End If
End Sub

Private Sub cmdAdd_Click()
    mState = ccAdd
    Call CtlClear(ccCmdAdd)
    Call CtlLock(ccSave)
    cboSex.SetFocus
End Sub

Private Sub cmdModify_Click()
    mState = ccModify
    Call CtlLock(ccSave)
    dtpExpireDt.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim strSpcCd    As String       '검체코드
    Dim strApplyDt  As String       '적용일
    Dim strSex      As String       '성별
    Dim lngAgeFr    As String       'From Age
    Dim lngAgeTo    As String       'To Age
    Dim intTemp     As Integer
    
    intTemp = MsgBox("정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub
    
    strSpcCd = lblSpcCd.Caption
    strApplyDt = Format$(dtpApplyDt.Value, "YYYYMMDD")
    Select Case cboSex.Text
        Case "남자": strSex = "M"
        Case "여자": strSex = "F"
        Case "Both": strSex = "B"
        Case "중성": strSex = "U"
    End Select
    lngAgeFr = CLng(txtFrAge.Text)
    lngAgeTo = CLng(txtToAge.Text)
    
    If mTMaster.RemoveRef(mTestCd, strSpcCd, strSex, lngAgeFr, lngAgeTo, strApplyDt) Then
        Call CtlLock(ccInit)
        Call GetRefList
        mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
    Else
        mdiIISMain.sbrStatus.Panels(2).Text = "삭제중 에러가 발생했습니다."
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim intTemp As Integer
    
    intTemp = MsgBox("변경된 내용을 취소할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub
    
    Call CtlLock(ccInit)
    Call lvwRefList_ItemClick(lvwRefList.SelectedItem)
    lvwRefList.SetFocus
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
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    
    '## 숫자, 문자, Enter, Backspcace만 입력할수 있도록함
    If KeyAscii >= 65 And KeyAscii <= 90 Then Exit Sub
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    
    KeyAscii = 0
End Sub

Private Sub txtFrAge_GotFocus()
    With txtFrAge
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFrAge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFrAge_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
On Error GoTo Errors
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        strTemp = UCase(Chr(KeyAscii))
        Select Case strTemp
            Case "D": txtFrAge.Text = CStr(CLng(txtFrAge.Text) * 365)
            Case "Y": txtFrAge.Text = CStr(CLng(txtFrAge.Text) / 365)
            Case "M": txtFrAge.Text = "50000"
        End Select
        
        KeyAscii = 0
    End If
    Exit Sub
    
Errors:
    MsgBox Err.Description, vbCritical, "오류"
    txtFrAge.Text = "0": KeyAscii = 0
End Sub

Private Sub txtToAge_GotFocus()
    With txtToAge
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtToAge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtToAge_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
On Error GoTo Errors
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        strTemp = UCase(Chr(KeyAscii))
        Select Case strTemp
            Case "D": txtToAge.Text = CStr(CLng(txtToAge.Text) * 365)
            Case "Y": txtToAge.Text = CStr(CLng(txtToAge.Text) / 365)
            Case "M": txtToAge.Text = "50000"
        End Select
        
        KeyAscii = 0
    End If
    Exit Sub
    
Errors:
    MsgBox Err.Description, vbCritical, "오류"
    txtFrAge.Text = "0": KeyAscii = 0
End Sub

Private Sub txtFrRef_GotFocus()
    With txtFrRef
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFrRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFrRef_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack _
        And KeyAscii <> vbKeyDecimal And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFrRef_Validate(Cancel As Boolean)
    txtFrRef.Text = Format$(txtFrRef.Text, ".0000")
    If Len(txtFrRef.Text) > 10 Then
        MsgBox "참고치는 최대 99999까지만 입력할수 있습니다.", vbInformation, "정보"
        Cancel = True
    End If
End Sub

Private Sub txtToRef_GotFocus()
    With txtToRef
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtToRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtToRef_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack _
        And KeyAscii <> vbKeyDecimal And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtToRef_Validate(Cancel As Boolean)
    txtToRef.Text = Format$(txtToRef.Text, ".0000")
    If Len(txtToRef.Text) > 10 Then
        MsgBox "참고치는 최대 99999까지만 입력할수 있습니다.", vbInformation, "정보"
        Cancel = True
    End If
End Sub

Private Sub txtAlpha_GotFocus()
    With txtAlpha
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAlpha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub lvwSpcList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## 검체코드, 검체명 표시
    Call CtlClear(ccLvwSpcList)
    lblSpcCd.Caption = Item.Text
    lblSpcNm.Caption = Item.SubItems(1)
    
    '## 참고치 적용일 표시
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
    Call CtlClear(ccTabRefList)
    
    strSpcCd = lblSpcCd.Caption
    strApplyDt = Format$(tabRefList.SelectedItem.Caption, "YYYYMMDD")
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
                itmX.SubItems(2) = CStr(.AgeFr / 365) & " - " & CStr(.AgeTo / 365)
                itmX.SubItems(3) = CStr(.RefFrVal) & " - " & CStr(.RefToVal)
            End If
        End With
    Next
    
    If lvwRefList.ListItems.Count > 13 Then
        lvwRefList.ColumnHeaders(1).Width = 690
    Else
        lvwRefList.ColumnHeaders(1).Width = 900
    End If
        
    Set objRef = Nothing
    Set objRefs = Nothing
    Set itmX = Nothing
    
    '## 참고치 정보표시
    Call lvwRefList_ItemClick(lvwRefList.SelectedItem)
End Sub

Private Sub lvwRefList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim objRefs     As clsIISRefs   '지정검체 컬렉션
    Dim strSpcCd    As String       '검체코드
    Dim strSex      As String       '성별
    Dim lngAgeFr    As Long         'From Age
    Dim lngAgeTo    As Long         'To Age
    Dim strApplyDt  As String       '적용일

    If Item Is Nothing Then Exit Sub
    
    strSpcCd = lblSpcCd.Caption
    Select Case Item.Text
        Case "남자": strSex = "M"
        Case "여자": strSex = "F"
        Case "Both": strSex = "B"
        Case "중성": strSex = "U"
    End Select
    lngAgeFr = CLng(Trim(mGetP(Item.SubItems(1), 1, "-")))
    lngAgeTo = CLng(Trim(mGetP(Item.SubItems(1), 2, "-")))
    strApplyDt = Format$(tabRefList.SelectedItem.Caption, "YYYYMMDD")
    
    '## 참고치 정보를 표시
    Set objRefs = mTMaster.Refs
    With objRefs(mTestCd, strSpcCd, strSex, lngAgeFr, lngAgeTo, strApplyDt)
        dtpApplyDt.Value = Format$(strApplyDt, "####-##-##")
        dtpExpireDt.Value = Format$(.ExpireDt, "####-##-##")
        cboSex.ListIndex = mFindCombo(cboSex, Item.Text)
        txtFrAge.Text = CStr(lngAgeFr)
        txtToAge.Text = CStr(lngAgeTo)
        txtFrRef.Text = CStr(.RefFrVal)
        txtToRef.Text = CStr(.RefToVal)
        txtAlpha.Text = .Refcd
    End With
    Set objRefs = Nothing
    Call CtlLock(ccModify)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 검사코드의 검체리스트를 lvwSpcList에 표시
'-----------------------------------------------------------------------------'
Private Sub GetSpcList()
    Dim itmX     As ListItem
    Dim objTSpcs As clsIISTSpcs     '지정검체 컬렉션
    Dim objTSpc  As clsIISTSpc      '지정검체 클래스
    Dim strSpcCd As String          '검체코드

    '## 검체코드, 검체명 표시
    lvwSpcList.ListItems.Clear
    Set objTSpcs = mTMaster.GetSpcInfo(mTestCd)
    If objTSpcs.Count = 0 Then
        Set objTSpcs = Nothing
        Exit Sub
    End If

    For Each objTSpc In objTSpcs
        If strSpcCd <> objTSpc.SpcCd Then
            strSpcCd = objTSpc.SpcCd
            Set itmX = lvwSpcList.ListItems.Add(, , strSpcCd)
            itmX.SubItems(1) = objTSpc.SpcNm
        End If
    Next

    If lvwSpcList.ListItems.Count > 14 Then
        lvwSpcList.ColumnHeaders(2).Width = 3780
    Else
        lvwSpcList.ColumnHeaders(2).Width = 4000
    End If
        
    Set itmX = Nothing
    Set objTSpc = Nothing
    Set objTSpcs = Nothing
    
    '## 해당검체의 참고치 리스트 표시
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
    
    Call CtlClear(ccTabRefList)
    tabRefList.Tabs.Clear
    
    strSpcCd = lblSpcCd.Caption
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
            dtpApplyDt.Enabled = False
            blnLock = True
            blnEnable = False
        Case StateEnum.ccSave
            cmdSave.Enabled = True
            cmdAdd.Enabled = False
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
            cmdCancel.Enabled = True
            blnLock = False
            blnEnable = True
            Select Case mState
                Case StateEnum.ccRefAdd
                    dtpApplyDt.Enabled = True
                    cboSex.Enabled = True
                    txtFrAge.Enabled = True
                    txtToAge.Enabled = True
                Case StateEnum.ccAdd
                    dtpApplyDt.Enabled = False
                    cboSex.Enabled = True
                    txtFrAge.Enabled = True
                    txtToAge.Enabled = True
                Case StateEnum.ccModify
                    dtpApplyDt.Enabled = False
                    cboSex.Enabled = False
                    txtFrAge.Enabled = False
                    txtToAge.Enabled = False
            End Select
        Case StateEnum.ccAdd
            cmdSave.Enabled = False
            cmdAdd.Enabled = True
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
            cmdCancel.Enabled = False
            dtpApplyDt.Enabled = False
            blnLock = True
            blnEnable = False
        Case StateEnum.ccModify
            cmdSave.Enabled = False
            cmdAdd.Enabled = True
            cmdModify.Enabled = True
            cmdDelete.Enabled = True
            cmdCancel.Enabled = False
            dtpApplyDt.Enabled = False
            blnLock = True
            blnEnable = False
    End Select
    
    txtTestCd.Locked = Not (blnLock)
    dtpExpireDt.Enabled = blnEnable
    cboSex.Locked = blnLock
    txtFrAge.Locked = blnLock
    txtToAge.Locked = blnLock
    txtFrRef.Locked = blnLock
    txtToRef.Locked = blnLock
    txtAlpha.Locked = blnLock
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 화면 컨트롤의 초기화
'   인수 :
'       1.pFlag : ClearEnum 상수
'-----------------------------------------------------------------------------'
Private Sub CtlClear(ByVal pFlag As ClearEnum)
    Select Case pFlag
        Case ClearEnum.ccAll
            txtTestCd.Text = "":        lblTestNm.Caption = ""
            lvwSpcList.ListItems.Clear: tabRefList.Tabs.Clear
            lvwRefList.ListItems.Clear: lblSpcCd.Caption = ""
            lblSpcNm.Caption = "":      dtpApplyDt.Value = Now
            dtpExpireDt.Value = "":     cboSex.ListIndex = -1
            txtFrAge.Text = "":         txtToAge.Text = ""
            txtFrRef.Text = "":         txtToRef.Text = ""
            txtAlpha.Text = ""
        Case ClearEnum.ccLvwSpcList
            tabRefList.Tabs.Clear:      txtAlpha.Text = ""
            lvwRefList.ListItems.Clear: lblSpcCd.Caption = ""
            lblSpcNm.Caption = "":      dtpApplyDt.Value = Now
            dtpExpireDt.Value = "":     cboSex.ListIndex = -1
            txtFrAge.Text = "":         txtToAge.Text = ""
            txtFrRef.Text = "":         txtToRef.Text = ""
        Case ClearEnum.ccTabRefList
            lvwRefList.ListItems.Clear:
            dtpApplyDt.Value = Now:     dtpExpireDt.Value = ""
            cboSex.ListIndex = -1:      txtAlpha.Text = ""
            txtFrAge.Text = "":         txtToAge.Text = ""
            txtFrRef.Text = "":         txtToRef.Text = ""
        Case ClearEnum.ccLvwRefList
            dtpApplyDt.Value = Now:     dtpExpireDt.Value = ""
            cboSex.ListIndex = -1:      txtAlpha.Text = ""
            txtFrAge.Text = "":         txtToAge.Text = ""
            txtFrRef.Text = "":         txtToRef.Text = ""
        Case ClearEnum.ccCmdAdd
            dtpExpireDt.Value = ""
            cboSex.ListIndex = -1:      txtAlpha.Text = ""
            txtFrAge.Text = "":         txtToAge.Text = ""
            txtFrRef.Text = "":         txtToRef.Text = ""
    End Select
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리1
'-----------------------------------------------------------------------------'
Private Sub mCode_SelectedItem(ByRef pSelItem As String)
    txtTestCd.Text = mGetP(pSelItem, 1, DIV)
End Sub

