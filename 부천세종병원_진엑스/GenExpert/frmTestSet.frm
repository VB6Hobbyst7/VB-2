VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   Caption         =   "장비 검사코드 설정"
   ClientHeight    =   7305
   ClientLeft      =   2670
   ClientTop       =   1290
   ClientWidth     =   9765
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   9765
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8250
      TabIndex        =   49
      Top             =   6600
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Height          =   4785
      Left            =   6060
      TabIndex        =   26
      Top             =   120
      Width           =   3525
      Begin VB.TextBox txtEquipCode 
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
         TabIndex        =   39
         Top             =   735
         Width           =   2115
      End
      Begin VB.TextBox txtCode 
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
         Top             =   1170
         Width           =   2115
      End
      Begin VB.TextBox txtDec 
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
         Top             =   2010
         Width           =   2115
      End
      Begin VB.TextBox txtName 
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
         Top             =   1590
         Width           =   2115
      End
      Begin VB.TextBox txtMuch 
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
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   300
         Width           =   2115
      End
      Begin VB.TextBox txtSeq 
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
         Top             =   2430
         Width           =   585
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2820
         Picture         =   "frmTestSet.frx":1272
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   33
         Top             =   1170
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   150
         TabIndex        =   32
         Top             =   4110
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   1230
         TabIndex        =   31
         Top             =   4110
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         Height          =   495
         Left            =   2310
         TabIndex        =   30
         Top             =   4110
         Width           =   1035
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
         Left            =   1110
         TabIndex        =   29
         Top             =   2850
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
         Left            =   2130
         TabIndex        =   28
         Top             =   2850
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtGubun 
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
         Left            =   4080
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   2115
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
         Left            =   240
         TabIndex        =   48
         Top             =   810
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
         Left            =   240
         TabIndex        =   47
         Top             =   1230
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
         Left            =   240
         TabIndex        =   46
         Top             =   2085
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
         Left            =   240
         TabIndex        =   45
         Top             =   1665
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
         Left            =   240
         TabIndex        =   44
         Top             =   375
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
         Left            =   240
         TabIndex        =   43
         Top             =   2520
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
         Left            =   270
         TabIndex        =   42
         Top             =   2940
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
         Left            =   1860
         TabIndex        =   41
         Top             =   2850
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
         Left            =   3210
         TabIndex        =   40
         Top             =   795
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5715
      Left            =   10380
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
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
         TabIndex        =   23
         Top             =   4980
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.TextBox txtOrdApi2 
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
         TabIndex        =   22
         Top             =   4590
         Visible         =   0   'False
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
         TabIndex        =   17
         Top             =   3810
         Width           =   4785
      End
      Begin VB.TextBox txtOrdApi1 
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
         TabIndex        =   16
         Top             =   4200
         Visible         =   0   'False
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
         TabIndex        =   14
         Top             =   1410
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
         TabIndex        =   13
         Top             =   1740
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
         TabIndex        =   5
         Top             =   690
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
         TabIndex        =   4
         Top             =   1020
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
         TabIndex        =   3
         Top             =   3120
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
         TabIndex        =   2
         Top             =   2730
         Width           =   4785
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3960
         Picture         =   "frmTestSet.frx":13BC
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   1
         Top             =   2370
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   4980
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label10 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Result"
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
         TabIndex        =   24
         Top             =   5055
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label28 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Barcode"
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
         TabIndex        =   21
         Top             =   4665
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   4590
         Visible         =   0   'False
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
         TabIndex        =   20
         Top             =   3885
         Width           =   810
      End
      Begin VB.Label Label26 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "WorklIst"
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
         TabIndex        =   19
         Top             =   4275
         Visible         =   0   'False
         Width           =   900
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
         TabIndex        =   18
         Top             =   3540
         Width           =   1020
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   3810
         Width           =   1305
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   4200
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   315
         Left            =   210
         Top             =   2730
         Width           =   1305
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
         TabIndex        =   15
         Top             =   1605
         Width           =   840
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
         TabIndex        =   12
         Top             =   885
         Width           =   420
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   2400
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
         TabIndex        =   9
         Top             =   3195
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
         TabIndex        =   8
         Top             =   2805
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   2430
         Width           =   1710
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6945
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   5805
      _Version        =   393216
      _ExtentX        =   10239
      _ExtentY        =   12250
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
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
    cmdSave.Caption = "Save"
    
End Sub

Private Sub DisplayList()

    ClearSpread vasList

    SQL = "SELECT GUBUN, EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          " GROUP BY GUBUN,EXAMCODE, EQUIPCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH "
    SQL = SQL & " ORDER BY SEQNO * 10 "
          
    Res = GetDBSelectVas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 12
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
    
    Res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
    If Res = 1 Then
        SQL = "UPDATE EQPMASTER " & vbCrLf & _
              "SET RESPREC = '" & Trim(txtDec) & "', " & vbCrLf & _
              "    EXAMNAME = '" & Trim(txtName) & "', " & vbCrLf & _
              "    GUBUN = '" & Trim(txtGubun) & "', " & vbCrLf & _
              "    REFLOW = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    REFHIGH = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    SEQNO = " & liSeqNo & " " & vbCrLf & _
              "WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "  AND EQUIPCODE = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  AND EXAMCODE = '" & Trim(txtCode) & "' "
    ElseIf Res = 0 Then
        SQL = "INSERT INTO EQPMASTER (EQUIPNO,GUBUN, EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO , REFLOW, REFHIGH) " & vbCrLf & _
              "VALUES ('" & gEquip & "','" & Trim(txtGubun) & "','" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', " & liSeqNo & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "') "
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
    Me.Height = 7725
    Me.Width = 9945
            
    ClearText
    DisplayList

    txtMuch = gEquip
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
    If Row = 0 Then
        Select Case Col
        Case 1
            vasSort vasList, 1, 2
        Case 2
            vasSort vasList, 2, 1
        Case 5
            vasSort vasList, 5, 1
        End Select
        Exit Sub
    End If
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "Save"
        ClearText
        Exit Sub
    End If
    
    txtGubun = Trim(GetText(vasList, Row, 1))
    
    txtEquipCode = Trim(GetText(vasList, Row, 2))
    txtCode = Trim(GetText(vasList, Row, 3))
    txtName = Trim(GetText(vasList, Row, 4))
    txtDec = Trim(GetText(vasList, Row, 5))
    txtSeq = Trim(GetText(vasList, Row, 6))
    txtRefLow = Trim(GetText(vasList, Row, 7))
    txtRefHigh = Trim(GetText(vasList, Row, 8))

    
    
    cmdSave.Caption = "Edit"
End Sub


