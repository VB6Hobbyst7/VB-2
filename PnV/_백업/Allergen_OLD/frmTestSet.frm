VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   11100
   ClientLeft      =   2670
   ClientTop       =   1290
   ClientWidth     =   17805
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11100
   ScaleWidth      =   17805
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   10935
      Begin VB.OptionButton optGubun 
         Caption         =   "Inhalant_A"
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
         Index           =   6
         Left            =   9060
         TabIndex        =   48
         Top             =   190
         Width           =   1455
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "Inhalant"
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
         Index           =   5
         Left            =   7605
         TabIndex        =   47
         Top             =   180
         Width           =   1395
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "Food_I"
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
         Left            =   6255
         TabIndex        =   46
         Top             =   185
         Width           =   1395
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "Food_A"
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
         Left            =   4800
         TabIndex        =   7
         Top             =   195
         Width           =   1395
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "Food"
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
         Left            =   3735
         TabIndex        =   6
         Top             =   205
         Width           =   1005
      End
      Begin VB.OptionButton optGubun 
         Caption         =   "Premium"
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
         Left            =   2265
         TabIndex        =   5
         Top             =   200
         Width           =   1395
      End
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
         Left            =   1230
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   975
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
         TabIndex        =   8
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   11160
      TabIndex        =   2
      Top             =   30
      Width           =   6555
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
         Height          =   465
         Left            =   5070
         TabIndex        =   63
         Top             =   180
         Width           =   1335
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
         Height          =   525
         Left            =   5280
         TabIndex        =   32
         Top             =   2280
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
         Height          =   525
         Left            =   5280
         TabIndex        =   31
         Top             =   1710
         Width           =   1095
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
         Height          =   525
         Left            =   5280
         TabIndex        =   30
         Top             =   1170
         Width           =   1095
      End
      Begin VB.ComboBox cboGubun 
         Height          =   300
         ItemData        =   "frmTestSet.frx":1272
         Left            =   1260
         List            =   "frmTestSet.frx":1274
         Style           =   2  '드롭다운 목록
         TabIndex        =   19
         Top             =   960
         Width           =   2145
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
         Left            =   4110
         TabIndex        =   18
         Top             =   3030
         Visible         =   0   'False
         Width           =   555
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
         Left            =   3090
         TabIndex        =   17
         Top             =   3030
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3570
         Picture         =   "frmTestSet.frx":1276
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   16
         Top             =   1650
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  '가운데 맞춤
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
         Left            =   3870
         TabIndex        =   15
         Top             =   2610
         Width           =   585
      End
      Begin VB.TextBox txtMuch 
         Appearance      =   0  '평면
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   1260
         TabIndex        =   14
         Top             =   540
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
         TabIndex        =   13
         Top             =   2220
         Width           =   3195
      End
      Begin VB.TextBox txtDec 
         Alignment       =   2  '가운데 맞춤
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
         TabIndex        =   12
         Top             =   2640
         Width           =   1215
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
         TabIndex        =   11
         Top             =   1800
         Width           =   2115
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
         TabIndex        =   10
         Top             =   1365
         Width           =   2115
      End
      Begin VB.CheckBox chkCommon 
         Caption         =   "선택"
         Height          =   345
         Left            =   1290
         TabIndex        =   9
         Top             =   3060
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[검사명 설정]"
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
         TabIndex        =   64
         Top             =   240
         Width           =   1290
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
         TabIndex        =   29
         Top             =   1035
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
         Left            =   3840
         TabIndex        =   28
         Top             =   3030
         Visible         =   0   'False
         Width           =   135
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
         Left            =   2250
         TabIndex        =   27
         Top             =   3120
         Visible         =   0   'False
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
         Left            =   3000
         TabIndex        =   26
         Top             =   2700
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
         TabIndex        =   25
         Top             =   615
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
         TabIndex        =   24
         Top             =   2325
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
         TabIndex        =   23
         Top             =   2745
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
         TabIndex        =   22
         Top             =   1890
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
         TabIndex        =   21
         Top             =   1440
         Width           =   720
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
         TabIndex        =   20
         Top             =   3120
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7905
      Left            =   11160
      TabIndex        =   1
      Top             =   3090
      Width           =   6555
      Begin VB.TextBox txtRstApi 
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
         Left            =   1530
         TabIndex        =   71
         Top             =   7410
         Width           =   4785
      End
      Begin VB.TextBox txtURL 
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
         Left            =   1530
         TabIndex        =   66
         Top             =   6630
         Width           =   4785
      End
      Begin VB.TextBox txtOrdApi 
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
         Left            =   1530
         TabIndex        =   65
         Top             =   7020
         Width           =   4785
      End
      Begin VB.TextBox txtIN1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
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
         Left            =   1530
         TabIndex        =   57
         Top             =   3570
         Width           =   4755
      End
      Begin VB.TextBox txtFI2 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
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
         Left            =   1530
         TabIndex        =   56
         Top             =   3180
         Width           =   4755
      End
      Begin VB.TextBox txtFI1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
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
         Left            =   1530
         TabIndex        =   55
         Top             =   2850
         Width           =   4755
      End
      Begin VB.TextBox txtIN2 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
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
         Left            =   1530
         TabIndex        =   54
         Top             =   3900
         Width           =   4755
      End
      Begin VB.TextBox txtIA1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
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
         Left            =   1530
         TabIndex        =   53
         Top             =   4290
         Width           =   4755
      End
      Begin VB.TextBox txtIA2 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
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
         Left            =   1530
         TabIndex        =   52
         Top             =   4620
         Width           =   4755
      End
      Begin VB.TextBox txtFA2 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
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
         Left            =   1530
         TabIndex        =   51
         Top             =   2460
         Width           =   4755
      End
      Begin VB.TextBox txtFA1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
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
         Left            =   1530
         TabIndex        =   50
         Top             =   2130
         Width           =   4755
      End
      Begin VB.TextBox txtFD2 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
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
         Left            =   1530
         TabIndex        =   49
         Top             =   1740
         Width           =   4755
      End
      Begin VB.TextBox txtPR1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
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
         Left            =   1530
         TabIndex        =   38
         Top             =   690
         Width           =   4755
      End
      Begin VB.TextBox txtPR2 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
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
         Left            =   1530
         TabIndex        =   37
         Top             =   1020
         Width           =   4755
      End
      Begin VB.TextBox txtFD1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
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
         Left            =   1530
         TabIndex        =   36
         Top             =   1410
         Width           =   4755
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
         Left            =   1530
         TabIndex        =   35
         Top             =   5940
         Width           =   4785
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
         Left            =   1530
         TabIndex        =   34
         Top             =   5550
         Width           =   4785
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3960
         Picture         =   "frmTestSet.frx":13C0
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   33
         Top             =   5190
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과 API"
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
         Left            =   450
         TabIndex        =   70
         Top             =   7485
         Width           =   720
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   7410
         Width           =   1305
      End
      Begin VB.Label Label27 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "URL"
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
         Left            =   450
         TabIndex        =   69
         Top             =   6705
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "오더 API"
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
         Left            =   450
         TabIndex        =   68
         Top             =   7095
         Width           =   720
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[URL 설정]"
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
         TabIndex        =   67
         Top             =   6360
         Width           =   1020
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   6630
         Width           =   1305
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   7020
         Width           =   1305
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   5940
         Width           =   1305
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   5550
         Width           =   1305
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   615
         Left            =   180
         Top             =   4290
         Width           =   1305
      End
      Begin VB.Label Label23 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Inhalant_A"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   62
         Top             =   4485
         Width           =   1050
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   615
         Left            =   180
         Top             =   3570
         Width           =   1305
      End
      Begin VB.Label Label22 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Inhalant"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   61
         Top             =   3765
         Width           =   840
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   615
         Left            =   180
         Top             =   2850
         Width           =   1305
      End
      Begin VB.Label Label21 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Food_I"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   60
         Top             =   3045
         Width           =   630
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   615
         Left            =   180
         Top             =   2130
         Width           =   1305
      End
      Begin VB.Label Label10 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Food_A"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   59
         Top             =   2325
         Width           =   630
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   615
         Left            =   180
         Top             =   1410
         Width           =   1305
      End
      Begin VB.Label Label9 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Food"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   58
         Top             =   1605
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   615
         Left            =   180
         Top             =   690
         Width           =   1305
      End
      Begin VB.Label Label13 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Prenium"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   45
         Top             =   885
         Width           =   735
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
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "[XLS경로 설정]"
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
         TabIndex        =   43
         Top             =   5220
         Width           =   1410
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
         Left            =   450
         TabIndex        =   42
         Top             =   6015
         Width           =   720
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
         Left            =   450
         TabIndex        =   41
         Top             =   5625
         Width           =   720
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
         TabIndex        =   40
         Top             =   390
         Width           =   2520
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
         TabIndex        =   39
         Top             =   5250
         Width           =   1710
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   10365
      Left            =   90
      TabIndex        =   0
      Top             =   660
      Width           =   10935
      _Version        =   393216
      _ExtentX        =   19288
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
      SpreadDesigner  =   "frmTestSet.frx":150A
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
    cboGubun.AddItem "PREMIUM"
    cboGubun.AddItem "FOOD"
    cboGubun.AddItem "FOOD_A"
    cboGubun.AddItem "FOOD_I"
    cboGubun.AddItem "INHALANT"
    cboGubun.AddItem "INHALANT_A"
    cboGubun.ListIndex = 0
    
    txtPR1 = gAssayNM.PR1
    txtPR2 = gAssayNM.PR2
    txtFD1 = gAssayNM.FD1
    txtFD2 = gAssayNM.FD2
    txtFA1 = gAssayNM.FA1
    txtFA2 = gAssayNM.FA2
    txtFI1 = gAssayNM.FI1
    txtFI2 = gAssayNM.FI2
    txtIN1 = gAssayNM.IN1
    txtIN2 = gAssayNM.IN2
    txtIA1 = gAssayNM.IA1
    txtIA2 = gAssayNM.IA2
    
    txtOrder = gAssayNM.OrderPath
    txtResult = gAssayNM.ResultPath
    
    txtURL = PnVAPI.APIURL
    txtOrdApi = PnVAPI.APIOrdPath
    txtRstApi = PnVAPI.APIRstPath
    
    cmdSave.Caption = "Save"
    chkCommon.Value = "0"
    
End Sub

Private Sub DisplayList()

    ClearSpread vasList

          SQL = "SELECT GUBUN, EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH, EXAMTYPE " & vbCrLf
    SQL = SQL & "  From EQPMASTER " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
    
    If optGubun(1).Value = True Then
        SQL = SQL & "  AND GUBUN = 'PREMIUM' "
    ElseIf optGubun(2).Value = True Then
        SQL = SQL & "  AND GUBUN = 'FOOD' "
    ElseIf optGubun(3).Value = True Then
        SQL = SQL & "  AND GUBUN = 'FOOD_A' "
    ElseIf optGubun(4).Value = True Then
        SQL = SQL & "  AND GUBUN = 'FOOD_I' "
    ElseIf optGubun(5).Value = True Then
        SQL = SQL & "  AND GUBUN = 'INHALANT' "
    ElseIf optGubun(6).Value = True Then
        SQL = SQL & "  AND GUBUN = 'INHALANT_A' "
    End If

    SQL = SQL & " GROUP BY GUBUN, EXAMCODE, EQUIPCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH,EXAMTYPE "
          
    SQL = SQL & " ORDER BY GUBUN, SEQNO * 10 "
          
    SetRawData "[SQL]" & SQL

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
        
'        txtRefLow.SetFocus
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

Private Sub txtFA1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtFA1.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "FA1", txtFA1.Text, App.Path & "\Interface.ini")
    End If

End Sub


Private Sub txtFA2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtFA2.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "FA2", txtFA2.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtFD1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtFD1.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "FD1", txtFD1.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtFD2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtFD2.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "FD2", txtFD2.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtFI1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtFI1.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "FI1", txtFI1.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtFI2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtFI2.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "FI2", txtFI2.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtIA1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Trim(txtIA1.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "IA1", txtIA1.Text, App.Path & "\Interface.ini")
    End If
    
End Sub

Private Sub txtIA2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Trim(txtIA2.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "IA2", txtIA2.Text, App.Path & "\Interface.ini")
    End If
    
End Sub

Private Sub txtIN1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Trim(txtIN1.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "IN1", txtIN1.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtIN2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Trim(txtIN2.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "IN2", txtIN2.Text, App.Path & "\Interface.ini")
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

Private Sub txtOrdApi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtOrdApi.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "ORDAPI", txtOrdApi.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtOrder.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "ORDER", txtOrder.Text, App.Path & "\Interface.ini")
    End If
End Sub



Private Sub txtPR1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtPR1.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "PR1", txtPR1.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtPR2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtPR2.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "PR2", txtPR2.Text, App.Path & "\Interface.ini")
    End If

End Sub

Private Sub txtResult_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtOrder.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "RESULT", txtResult.Text, App.Path & "\Interface.ini")
    End If
End Sub

Private Sub txtRstApi_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtRstApi.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "RSTAPI", txtRstApi.Text, App.Path & "\Interface.ini")
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

Private Sub txtURL_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 And Trim(txtURL.Text) <> "" Then
        Call WritePrivateProfileString("Assay", "URL", txtURL.Text, App.Path & "\Interface.ini")
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
