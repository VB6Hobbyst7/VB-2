VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   11100
   ClientLeft      =   2670
   ClientTop       =   1290
   ClientWidth     =   17220
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11100
   ScaleWidth      =   17220
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   15600
      TabIndex        =   50
      Top             =   10350
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3765
      Left            =   12060
      TabIndex        =   32
      Top             =   6360
      Width           =   4875
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3960
         Picture         =   "frmTestSet.frx":1272
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   39
         Top             =   2460
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtOrder 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   38
         Top             =   2820
         Width           =   3495
      End
      Begin VB.TextBox txtResult 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   37
         Top             =   3210
         Width           =   3495
      End
      Begin VB.TextBox txtAT 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   36
         Top             =   1470
         Width           =   3195
      End
      Begin VB.TextBox txtFD 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   35
         Top             =   1080
         Width           =   3195
      End
      Begin VB.TextBox txtIN 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   34
         Top             =   690
         Width           =   3195
      End
      Begin VB.TextBox txtCommon 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   33
         Top             =   1860
         Width           =   3195
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "* 경로끝의 \는 제외"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   1830
         TabIndex        =   49
         Top             =   2520
         Width           =   1710
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "* 대소문자 정확히 입력하세요"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   1770
         TabIndex        =   48
         Top             =   390
         Width           =   2520
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "오더경로"
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
         Left            =   240
         TabIndex        =   47
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과경로"
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
         Left            =   240
         TabIndex        =   46
         Top             =   3285
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[XML경로 설정]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   270
         TabIndex        =   45
         Top             =   2490
         Width           =   1410
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[Assay명 설정]"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   270
         TabIndex        =   44
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "ATOPY"
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
         Left            =   240
         TabIndex        =   43
         Top             =   1545
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "FOOD"
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
         Left            =   240
         TabIndex        =   42
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "INHALANT"
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
         Left            =   240
         TabIndex        =   41
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "COMMON"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1935
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   12060
      TabIndex        =   7
      Top             =   510
      Width           =   4875
      Begin VB.CheckBox chkCommon 
         Caption         =   "선택"
         Height          =   345
         Left            =   1290
         TabIndex        =   21
         Top             =   3510
         Width           =   945
      End
      Begin VB.TextBox txtEquipCode 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   20
         Top             =   1335
         Width           =   2115
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   19
         Top             =   1770
         Width           =   2115
      End
      Begin VB.TextBox txtDec 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   18
         Top             =   2610
         Width           =   2115
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   17
         Top             =   2190
         Width           =   3195
      End
      Begin VB.TextBox txtMuch 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   16
         Top             =   360
         Width           =   2115
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   15
         Top             =   3060
         Width           =   585
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3570
         Picture         =   "frmTestSet.frx":13BC
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   14
         Top             =   1500
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtRefLow 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   13
         Top             =   3990
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtRefHigh 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   12
         Top             =   3990
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cboGubun 
         Height          =   300
         Left            =   1260
         Style           =   2  '드롭다운 목록
         TabIndex        =   11
         Top             =   870
         Width           =   2145
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1260
         TabIndex        =   10
         Top             =   5010
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2400
         TabIndex        =   9
         Top             =   5010
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3540
         TabIndex        =   8
         Top             =   5010
         Width           =   1095
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "공통코드"
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
         Left            =   390
         TabIndex        =   31
         Top             =   3570
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비채널"
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
         Left            =   390
         TabIndex        =   30
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
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
         Left            =   390
         TabIndex        =   29
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "소 수 점"
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
         Left            =   390
         TabIndex        =   28
         Top             =   2685
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 명"
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
         Left            =   390
         TabIndex        =   27
         Top             =   2265
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비구분"
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
         Left            =   390
         TabIndex        =   26
         Top             =   435
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    서"
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
         Left            =   390
         TabIndex        =   25
         Top             =   3150
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참 고 치"
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
         Left            =   420
         TabIndex        =   24
         Top             =   4080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2010
         TabIndex        =   23
         Top             =   3990
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사구분"
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
         Left            =   390
         TabIndex        =   22
         Top             =   945
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   11895
      Begin VB.OptionButton optGubun 
         Caption         =   "전체"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1530
         TabIndex        =   5
         Top             =   210
         Width           =   975
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "INHALANT"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2565
         TabIndex        =   4
         Top             =   195
         Width           =   1395
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "FOOD"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4065
         TabIndex        =   3
         Top             =   195
         Width           =   1005
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "ATOPY"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   5250
         TabIndex        =   2
         Top             =   180
         Width           =   1395
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "IN/FD"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   6630
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.Label Label20 
         Caption         =   "검사구분"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   270
         Width           =   915
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   10365
      Left            =   90
      TabIndex        =   51
      Top             =   660
      Width           =   11895
      _Version        =   393216
      _ExtentX        =   20981
      _ExtentY        =   18283
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   20
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmTestSet.frx":1506
   End
End
Attribute VB_Name = "frmTestSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearText()

    txtEquipCode = ""
    txtCode = ""
    txtName = ""
    txtDec = "1"
    txtSeq = ""
    txtRefLow = ""
    txtRefHigh = ""
    
    cboGubun.Clear
    cboGubun.AddItem "INHALANT"
    cboGubun.AddItem "FOOD"
    cboGubun.AddItem "ATOPY"
    cboGubun.AddItem "COMMON"
    cboGubun.ListIndex = 0
    
    txtIN = gAssayNM.INHALANT
    txtFD = gAssayNM.FOOD
    txtAT = gAssayNM.ATOPY
    txtCommon = gAssayNM.COMMON
    
    txtOrder = gAssayNM.OrderPath
    txtResult = gAssayNM.ResultPath
    
    cmdSave.Caption = "Save"
    chkCommon.Value = "0"
    
End Sub

Private Sub DisplayList()

    ClearSpread vasList

          SQL = "SELECT GUBUN, EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH, EXAMTYPE " & vbCrLf
    SQL = SQL & "  From EQPMASTER " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
    
    If optGubun(1).Value = True Then
        SQL = SQL & "  AND GUBUN = 'INHALANT' "
    ElseIf optGubun(2).Value = True Then
        SQL = SQL & "  AND GUBUN = 'FOOD' "
    ElseIf optGubun(3).Value = True Then
        SQL = SQL & "  AND GUBUN = 'ATOPY' "
    ElseIf optGubun(4).Value = True Then
        SQL = SQL & "  AND GUBUN = 'COMMON' "
    End If
    
    SQL = SQL & " GROUP BY GUBUN, EXAMCODE, EQUIPCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH,EXAMTYPE "
          
    SQL = SQL & " ORDER BY GUBUN, SEQNO * 10 "
          
'    SetRawData "[SQL]" & SQL

    Res = GetDBSelectVas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 10
    'Call vasList_Click(1, 0)
    
End Sub

'-- 장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure
Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH " & vbCrLf & _
          "  FROM EQPMASTER " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND GUBUN = '" & cboGubun.Text & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & asEquipCode & "' "
          
    If Trim(asSuga) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND EXAMCODE = '" & asSuga & "' "
    End If
    
    Res = GetDBSelectColumn(gLocal, SQL)
    If Res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf Res = -1 Then
        ExistOfEquipCode = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Or Trim(gReadBuf(1)) <> asSuga Then
        Exit Function
    End If
        
    ExistOfEquipCode = 1
End Function

'-- 검사구분  + 장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure
Function ExistOfEquipCode_Allergy(asGubun As String, asEquipCode As String, Optional asSuga As String = "") As Integer

    ExistOfEquipCode_Allergy = -1
    
    If asGubun = "" Then
        Exit Function
    End If
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER " & vbCrLf
    SQL = SQL & " WHERE GUBUN = '" & asGubun & "' " & vbCrLf
    SQL = SQL & "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "   AND EQUIPCODE = '" & asEquipCode & "' "
          
    If Trim(asSuga) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND EXAMCODE = '" & asSuga & "' "
    End If
    
    Res = GetDBSelectColumn(gLocal, SQL)
    If Res = 0 Then
        ExistOfEquipCode_Allergy = 0
        Exit Function
    ElseIf Res = -1 Then
        ExistOfEquipCode_Allergy = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Or Trim(gReadBuf(1)) <> asSuga Then
        Exit Function
    End If
        
    ExistOfEquipCode_Allergy = 1
End Function


Private Sub cmdCancel_Click()
    ClearText
    txtEquipCode.SetFocus
End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
    SQL = "DELETE FROM EQPMASTER " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND GUBUN = '" & Trim(cboGubun.Text) & "' " & vbCrLf & _
          "  AND EQUIPCODE = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
          "  AND EXAMCODE = '" & Trim(txtCode) & "' "
    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
        Exit Sub
    End If
    
    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
    Dim lsFlag As String
    Dim lsResFlag As String
    Dim liSeqNo As Integer

    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        MsgBox "장비코드를 입력하세요", vbInformation
        Exit Sub
    End If
    
    If Trim(txtDec) = "" Then
        txtDec.Text = 1

    End If
    
    If IsNumeric(txtSeq) Then
        liSeqNo = CInt(txtSeq)
    Else
        liSeqNo = 0
    End If
    
    Res = ExistOfEquipCode_Allergy(Trim(cboGubun.Text), Trim(txtEquipCode), Trim(txtCode))
    If Res = 1 Then
        SQL = "UPDATE EQPMASTER " & vbCrLf & _
              "SET RESPREC = '" & Trim(txtDec) & "', " & vbCrLf & _
              "    EXAMNAME = '" & Trim(txtName) & "', " & vbCrLf & _
              "    REFLOW = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    REFHIGH = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    SEQNO = " & liSeqNo & ", " & vbCrLf & _
              "    EXAMTYPE = '" & IIf(chkCommon.Value = "1", "공통", "") & "' " & vbCrLf & _
              "WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "  AND GUBUN = '" & Trim(cboGubun.Text) & "' " & vbCrLf & _
              "  AND EQUIPCODE = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  AND EXAMCODE = '" & Trim(txtCode) & "' "
    ElseIf Res = 0 Then
        SQL = "INSERT INTO EQPMASTER (EQUIPNO,GUBUN, EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO , REFLOW, REFHIGH, EXAMTYPE) " & vbCrLf & _
              "VALUES ('" & gEquip & "', '" & Trim(cboGubun.Text) & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', " & liSeqNo & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "','" & IIf(chkCommon.Value = "1", "공통", "") & "') "
    End If

    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    DisplayList
    
    cmdCancel_Click
    
End Sub


Private Sub Form_Load()
'    Me.Height = 7725
'    Me.Width = 9945
            
    ClearText
    DisplayList

    txtMuch = gEquip
End Sub

Private Sub optGubun_Click(Index As Integer)
    
    Call DisplayList

End Sub


Private Sub txtAT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtAT.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "AT", txtAT.Text, App.Path & "\Interface.ini")
    End If
End Sub

Private Sub txtCommon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtCommon.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "CM", txtCommon.Text, App.Path & "\Interface.ini")
    End If
End Sub

Private Sub txtEquipCode_GotFocus()
    SelectFocus txtEquipCode
End Sub

Private Sub txtEquipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtEquipCode = "" Then
            txtEquipCode.SetFocus
            Exit Sub
        End If
        txtCode.SetFocus
    End If
End Sub

Private Sub txtDec_GotFocus()
    SelectFocus txtDec
End Sub

Private Sub txtDec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtDec = "" Then
            txtDec.SetFocus
'            Exit Sub
        End If
        
        txtRefLow.SetFocus
    End If
End Sub

Private Sub txtcode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        Res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
        If Res = -1 Then
            txtCode.SetFocus
            Exit Sub
        ElseIf Res = 0 Then
            cmdSave.Caption = "Save"
            
        ElseIf Res = 1 Then
            cmdSave.Caption = "Edit"
            txtName = Trim(gReadBuf(2))
            txtDec = Trim(gReadBuf(3))
            txtSeq = Trim(gReadBuf(4))
            txtRefLow = Trim(gReadBuf(5))
            txtRefHigh = Trim(gReadBuf(6))
        End If
        
        txtName.SetFocus
    End If
    
End Sub



Private Sub txtFD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtFD.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "FD", txtFD.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtIN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtIN.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "IN", txtIN.Text, App.Path & "\Interface.ini")
    End If
End Sub

Private Sub txtMuch_GotFocus()

    SelectFocus txtMuch
    
End Sub

Private Sub txtMuch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtMuch.Text) = "" Then
            txtMuch.SetFocus
            Exit Sub
        End If
        txtEquipCode.SetFocus
    End If
End Sub

Private Sub txtName_GotFocus()
    SelectFocus txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtName.Text) = "" Then
            txtName.SetFocus
            Exit Sub
        End If
        txtDec.SetFocus
        
    End If
End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtOrder.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "ORDER", txtOrder.Text, App.Path & "\Interface.ini")
    End If
End Sub

Private Sub txtResult_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtOrder.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "RESULT", txtResult.Text, App.Path & "\Interface.ini")
    End If
End Sub

Private Sub txtSeq_GotFocus()
    SelectFocus txtSeq
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtSeq.Text) = "" Then
            txtSeq.SetFocus
            Exit Sub
        End If

        cmdSave.SetFocus
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    
    DoEvents
    
    If Row = 0 Then
        Select Case Col
        Case 1
            vasSort vasList, 1, 2
        Case 2
            vasSort vasList, 2, 1
        Case 3
            vasSort vasList, 3, 1
        Case 4
            vasSort vasList, 4, 1
        Case 5
            vasSort vasList, 5, 1
        Case 6
            vasSort vasList, 5, 1
        Case 7
            vasSort vasList, 7, 1
        End Select
        Exit Sub
    End If
    
    
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "Save"
        ClearText
        Exit Sub
    End If
    cboGubun.Text = Trim(GetText(vasList, Row, 1))
    txtEquipCode = Trim(GetText(vasList, Row, 2))
    txtCode = Trim(GetText(vasList, Row, 3))
    txtName = Trim(GetText(vasList, Row, 4))
    txtDec = Trim(GetText(vasList, Row, 5))
    txtSeq = Trim(GetText(vasList, Row, 6))
    'txtRefLow = Trim(GetText(vasList, Row, 7))
    'txtRefHigh = Trim(GetText(vasList, Row, 8))
    If GetText(vasList, Row, 9) = "공통" Then
        chkCommon.Value = "1"
    Else
        chkCommon.Value = "0"
    End If
    cmdSave.Caption = "Edit"

End Sub


