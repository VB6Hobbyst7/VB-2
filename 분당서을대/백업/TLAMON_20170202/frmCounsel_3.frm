VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Begin VB.Form frmCounsel_3 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin ChartfxLibCtl.ChartFX chtSnack 
      Height          =   4335
      Left            =   7800
      TabIndex        =   21
      Top             =   3090
      Width           =   3885
      _cx             =   6853
      _cy             =   7646
      Build           =   20
      TypeMask        =   111673364
      LeftGap         =   0
      RightGap        =   0
      TopGap          =   0
      BottomGap       =   0
      WallWidth       =   3
      CylSides        =   30
      Volume          =   40
      AxesStyle       =   0
      Axis(0).Max     =   800
      Axis(0).Decimals=   0
      Axis(0).Style   =   14440
      Axis(2).Min     =   0
      Axis(2).Max     =   700
      Axis(2).Style   =   10344
      RGBBk           =   15724527
      nColors         =   16
      Colors          =   "frmCounsel_3.frx":0000
      nPts            =   10
      nSer            =   1
      NumPoint        =   10
      NumSer          =   1
      BorderS         =   0
      _Data_          =   "frmCounsel_3.frx":00A0
   End
   Begin VB.Image imgUnit5 
      Height          =   900
      Index           =   3
      Left            =   5400
      Top             =   7110
      Width           =   900
   End
   Begin VB.Image imgUnit5 
      Height          =   900
      Index           =   2
      Left            =   5400
      Top             =   5640
      Width           =   900
   End
   Begin VB.Image imgUnit5 
      Height          =   900
      Index           =   1
      Left            =   5400
      Top             =   4020
      Width           =   900
   End
   Begin VB.Image imgUnit5 
      Height          =   900
      Index           =   0
      Left            =   5400
      Top             =   2520
      Width           =   900
   End
   Begin VB.Image imgUnit4 
      Height          =   900
      Index           =   3
      Left            =   4320
      Top             =   7140
      Width           =   900
   End
   Begin VB.Image imgUnit4 
      Height          =   900
      Index           =   2
      Left            =   4290
      Top             =   5640
      Width           =   900
   End
   Begin VB.Image imgUnit4 
      Height          =   900
      Index           =   1
      Left            =   4320
      Top             =   4020
      Width           =   900
   End
   Begin VB.Image imgUnit4 
      Height          =   900
      Index           =   0
      Left            =   4320
      Top             =   2520
      Width           =   900
   End
   Begin VB.Image imgUnit3 
      Height          =   900
      Index           =   3
      Left            =   3270
      Top             =   7140
      Width           =   900
   End
   Begin VB.Image imgUnit3 
      Height          =   900
      Index           =   2
      Left            =   3270
      Top             =   5640
      Width           =   900
   End
   Begin VB.Image imgUnit3 
      Height          =   900
      Index           =   1
      Left            =   3270
      Top             =   4020
      Width           =   900
   End
   Begin VB.Image imgUnit3 
      Height          =   900
      Index           =   0
      Left            =   3270
      Top             =   2520
      Width           =   900
   End
   Begin VB.Image imgUnit2 
      Height          =   900
      Index           =   3
      Left            =   2190
      Top             =   7140
      Width           =   900
   End
   Begin VB.Image imgUnit2 
      Height          =   900
      Index           =   2
      Left            =   2190
      Top             =   5640
      Width           =   900
   End
   Begin VB.Image imgUnit2 
      Height          =   900
      Index           =   1
      Left            =   2190
      Top             =   4050
      Width           =   900
   End
   Begin VB.Image imgUnit2 
      Height          =   900
      Index           =   0
      Left            =   2190
      Top             =   2520
      Width           =   900
   End
   Begin VB.Image imgUnit1 
      Height          =   900
      Index           =   3
      Left            =   1110
      Top             =   7140
      Width           =   900
   End
   Begin VB.Image imgUnit1 
      Height          =   900
      Index           =   2
      Left            =   1110
      Top             =   5640
      Width           =   900
   End
   Begin VB.Image imgUnit1 
      Height          =   900
      Index           =   1
      Left            =   1110
      Top             =   4050
      Width           =   900
   End
   Begin VB.Image imgUnit1 
      Height          =   900
      Index           =   0
      Left            =   1110
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label lblExTime 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H0080FFFF&
      Caption         =   " 떡볶이1인분 = 500kcal = 달리기 30분"
      Height          =   240
      Left            =   6690
      TabIndex        =   54
      Top             =   7950
      Width           =   5445
   End
   Begin VB.Image TopImage 
      Height          =   960
      Left            =   -30
      Picture         =   "frmCounsel_3.frx":01AD
      Top             =   50
      Width           =   13140
   End
   Begin VB.Label lblFCal 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "150"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   11000
      TabIndex        =   53
      Top             =   2460
      Width           =   500
   End
   Begin VB.Label lblFMain 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "귤 1개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   6870
      TabIndex        =   52
      Top             =   2115
      Width           =   855
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   19
      Left            =   11130
      TabIndex        =   51
      Top             =   6800
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   18
      Left            =   10035
      TabIndex        =   50
      Top             =   6800
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   17
      Left            =   8935
      TabIndex        =   49
      Top             =   6800
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   7860
      TabIndex        =   48
      Top             =   6800
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   6750
      TabIndex        =   47
      Top             =   6800
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   11070
      TabIndex        =   46
      Top             =   5650
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   10035
      TabIndex        =   45
      Top             =   5650
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   8935
      TabIndex        =   44
      Top             =   5650
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   7860
      TabIndex        =   43
      Top             =   5650
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   6750
      TabIndex        =   42
      Top             =   5650
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   11100
      TabIndex        =   41
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   10035
      TabIndex        =   40
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   8935
      TabIndex        =   39
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   7860
      TabIndex        =   38
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   6750
      TabIndex        =   37
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   11070
      TabIndex        =   36
      Top             =   3310
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   10035
      TabIndex        =   35
      Top             =   3310
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   8935
      TabIndex        =   34
      Top             =   3310
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   7860
      TabIndex        =   33
      Top             =   3310
      Width           =   915
   End
   Begin VB.Label lblFruit 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "금귤 5개"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   6750
      TabIndex        =   32
      Top             =   3310
      Width           =   915
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   9
      Left            =   6750
      TabIndex        =   31
      Top             =   7050
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   8
      Left            =   6750
      TabIndex        =   30
      Top             =   6600
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   7
      Left            =   6750
      TabIndex        =   29
      Top             =   6180
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   6
      Left            =   6750
      TabIndex        =   28
      Top             =   5790
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   5
      Left            =   6750
      TabIndex        =   27
      Top             =   5370
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   4
      Left            =   6750
      TabIndex        =   26
      Top             =   4920
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   3
      Left            =   6750
      TabIndex        =   25
      Top             =   4530
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   2
      Left            =   6750
      TabIndex        =   24
      Top             =   4080
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   6750
      TabIndex        =   23
      Top             =   3660
      Width           =   885
   End
   Begin VB.Label lblSnack 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "떡볶이"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   6750
      TabIndex        =   22
      Top             =   3240
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,020 kcal"
      Height          =   400
      Index           =   3
      Left            =   11040
      TabIndex        =   20
      Top             =   3330
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "915 kcal"
      Height          =   400
      Index           =   4
      Left            =   11040
      TabIndex        =   19
      Top             =   5220
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "610 kcal"
      Height          =   400
      Index           =   5
      Left            =   11040
      TabIndex        =   18
      Top             =   7050
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "610 kcal"
      Height          =   255
      Index           =   2
      Left            =   10020
      TabIndex        =   17
      Top             =   7050
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "915 kcal"
      Height          =   255
      Index           =   1
      Left            =   10020
      TabIndex        =   16
      Top             =   5220
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,020 kcal"
      Height          =   255
      Index           =   0
      Left            =   10020
      TabIndex        =   15
      Top             =   3330
      Width           =   855
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   6730
      TabIndex        =   14
      Top             =   7830
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   6730
      TabIndex        =   13
      Top             =   7530
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   6730
      TabIndex        =   12
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   6730
      TabIndex        =   11
      Top             =   6900
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   6730
      TabIndex        =   10
      Top             =   6570
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   6730
      TabIndex        =   9
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   6730
      TabIndex        =   8
      Top             =   5940
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   6730
      TabIndex        =   7
      Top             =   5610
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   6730
      TabIndex        =   6
      Top             =   5190
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   6730
      TabIndex        =   5
      Top             =   4650
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   6730
      TabIndex        =   4
      Top             =   4110
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   6730
      TabIndex        =   3
      Top             =   3690
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6730
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6730
      TabIndex        =   1
      Top             =   3060
      Width           =   735
   End
   Begin VB.Label lblPortion 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.5인분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6730
      TabIndex        =   0
      Top             =   2730
      Width           =   735
   End
   Begin VB.Image imgSnack 
      Height          =   900
      Index           =   5
      Left            =   5400
      Top             =   7230
      Width           =   900
   End
   Begin VB.Image imgSnack 
      Height          =   900
      Index           =   4
      Left            =   4470
      Top             =   7230
      Width           =   900
   End
   Begin VB.Image imgSnack 
      Height          =   900
      Index           =   3
      Left            =   3540
      Top             =   7230
      Width           =   900
   End
   Begin VB.Image imgSnack 
      Height          =   900
      Index           =   2
      Left            =   2610
      Top             =   7230
      Width           =   900
   End
   Begin VB.Image imgSnack 
      Height          =   900
      Index           =   1
      Left            =   1680
      Top             =   7230
      Width           =   900
   End
   Begin VB.Image imgSnack 
      Height          =   900
      Index           =   0
      Left            =   750
      Top             =   7230
      Width           =   900
   End
   Begin VB.Image imgDinner 
      Height          =   900
      Index           =   5
      Left            =   5400
      Top             =   5490
      Width           =   900
   End
   Begin VB.Image imgDinner 
      Height          =   900
      Index           =   4
      Left            =   4470
      Top             =   5490
      Width           =   900
   End
   Begin VB.Image imgDinner 
      Height          =   900
      Index           =   3
      Left            =   3540
      Top             =   5490
      Width           =   900
   End
   Begin VB.Image imgDinner 
      Height          =   900
      Index           =   2
      Left            =   2610
      Top             =   5490
      Width           =   900
   End
   Begin VB.Image imgDinner 
      Height          =   900
      Index           =   1
      Left            =   1680
      Top             =   5490
      Width           =   900
   End
   Begin VB.Image imgDinner 
      Height          =   900
      Index           =   0
      Left            =   750
      Top             =   5490
      Width           =   900
   End
   Begin VB.Image imgLunch 
      Height          =   900
      Index           =   5
      Left            =   5400
      Top             =   3750
      Width           =   900
   End
   Begin VB.Image imgLunch 
      Height          =   900
      Index           =   4
      Left            =   4470
      Top             =   3750
      Width           =   900
   End
   Begin VB.Image imgLunch 
      Height          =   900
      Index           =   3
      Left            =   3540
      Top             =   3750
      Width           =   900
   End
   Begin VB.Image imgLunch 
      Height          =   900
      Index           =   2
      Left            =   2610
      Top             =   3750
      Width           =   900
   End
   Begin VB.Image imgLunch 
      Height          =   900
      Index           =   1
      Left            =   1680
      Top             =   3750
      Width           =   900
   End
   Begin VB.Image imgLunch 
      Height          =   900
      Index           =   0
      Left            =   750
      Top             =   3750
      Width           =   900
   End
   Begin VB.Image imgBreak 
      Height          =   900
      Index           =   5
      Left            =   5400
      Top             =   1980
      Width           =   900
   End
   Begin VB.Image imgBreak 
      Height          =   900
      Index           =   4
      Left            =   4470
      Top             =   1980
      Width           =   900
   End
   Begin VB.Image imgBreak 
      Height          =   900
      Index           =   3
      Left            =   3540
      Top             =   1980
      Width           =   900
   End
   Begin VB.Image imgBreak 
      Height          =   900
      Index           =   2
      Left            =   2610
      Top             =   1980
      Width           =   900
   End
   Begin VB.Image imgBreak 
      Height          =   900
      Index           =   1
      Left            =   1680
      Top             =   1980
      Width           =   900
   End
   Begin VB.Image imgBreak 
      Height          =   900
      Index           =   0
      Left            =   750
      Top             =   1980
      Width           =   900
   End
   Begin VB.Image imgSub 
      Height          =   1845
      Index           =   1
      Left            =   12330
      Picture         =   "frmCounsel_3.frx":1B33
      Top             =   6450
      Width           =   315
   End
   Begin VB.Image imgSub 
      Height          =   1845
      Index           =   0
      Left            =   12330
      Picture         =   "frmCounsel_3.frx":2647
      Top             =   4620
      Width           =   315
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   3
      Left            =   9180
      Picture         =   "frmCounsel_3.frx":3319
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   2
      Left            =   8280
      Picture         =   "frmCounsel_3.frx":3942
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   1
      Left            =   7380
      Picture         =   "frmCounsel_3.frx":3F57
      Top             =   1440
      Width           =   900
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   0
      Left            =   6480
      Picture         =   "frmCounsel_3.frx":4592
      Top             =   1440
      Width           =   900
   End
End
Attribute VB_Name = "frmCounsel_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'+---------------------------------------------------------------------------------+
'| 상담 > 식사 > 식사가이드
'+---------------------------------------------------------------------------------+
Private Const PATH03 As String = "\Back\Counsel\03\"
Private Const PATH03UNIT As String = "\Back\Counsel\03\Unit\"
Private Const IMG_BREAK As String = "상담-식사가이드 아침.jpg"
Private Const IMG_LUNCH As String = "상담-식사가이드 점심.jpg"
Private Const IMG_DINNER As String = "상담-식사가이드 저녁.jpg"
Private Const IMG_DINNER2 As String = "상담-식사가이드 저녁2.jpg"
Private Const IMG_SNACK As String = "간식back.jpg"
Private Const IMG_SNACK2 As String = "간식2 back.jpg"
Private Const IMG_BREAK_ON As String = "아침 on.jpg"
Private Const IMG_BREAK_OFF As String = "아침 off.jpg"
Private Const IMG_LUNCH_ON As String = "점심 on.jpg"
Private Const IMG_LUNCH_OFF As String = "점심 off.jpg"
Private Const IMG_DINNER_ON As String = "저녁 on.jpg"
Private Const IMG_DINNER_OFF As String = "저녁 off.jpg"
Private Const IMG_SNACK_ON As String = "간식 on.jpg"
Private Const IMG_SNACK_OFF As String = "간식 off.jpg"
Private Const IMG_SUB1_ON As String = "술과 안주의 열량ON.jpg"
Private Const IMG_SUB1_OFF As String = "술과 안주의 열량OFF.jpg"
Private Const IMG_SUB2_ON As String = "술종류별 열량ON.jpg"
Private Const IMG_SUB2_OFF As String = "술종류별 열량OFF.jpg"
Private Const IMG_SSUB1_ON As String = "대체과일 on.jpg"
Private Const IMG_SSUB1_OFF As String = "대체과일 off.jpg"
Private Const IMG_SSUB2_ON As String = "간식대 운동 on.jpg"
Private Const IMG_SSUB2_OFF As String = "간식대 운동 off.jpg"

Private Const IMG_MEAL As String = "\Back\Meal\normal\"

Private mintTab As Integer   ' 3: 저녁 / 4: 간식
Private intSet As Integer    ' 식단번호
Private mintConfig As Integer   '식단정보에 설정된 번호
Private intSnack(10) As Integer
Private sngGrain(3) As Single
Private sngFishMeat(3) As Single
Private sngSnackMilk As Single
Private sngSnackFruit As Single

Public Sub Form_Load()
    Set Me.Picture = LoadPicture(App.Path & "\Back\Counsel\03\상담-식사가이드 아침.jpg")
    Dim i As Integer
    Dim qrySelect As String, rValue As Variant
    Dim sngChMeal As Single
    
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.Height = FRM_HEIGHT
    Me.Width = FRM_WIDTH
    Me.BackColor = vbWhite

    '식단작성정보에 따라 화면구성 달라짐
    ' [1] 식단정보에 저장된 설정을 불러온다
'0:  한식
'1:  한식 빵 / 씨리얼(아침)
'2:  한식 빵 / 씨리얼(점심)
'3:  한식 빵 / 씨리얼(저녁)
'4:  한식 대용식(아침 + 간식)
'5:  한식 대용식(점심 + 간식)
'6:  한식 대용식(저녁 + 간식)
'7:  한식 대용식(아침)
'8:  한식 대용식(점심)
'9:  한식 대용식(저녁)
    '   [1]-1 한식+빵/씨리얼인 경우 - 1,2,3
    '       [1]-1-1 처방칼로리와 감량칼로리에 의해 식단번호와 해당 단위수를 불러온다(곡류군, 우유군, 과일군)
    '       [1]-1-2 설정된 끼니의 곡류군 단위가 2 이상이면
    '               해당 끼니의 곡류군 단위는 2로 하고, 나머지 끼니에 (아)35:(점)35:(저)30의 비율을 유지해 분배한다.
    '       [1]-1-3 설정된 끼니의 곡류군 단위가 2 보다 작으면 주어진 단위수만큼 빵을 준다.
    '       [1]-1-4 설정된 끼니에 우유군과 과일군을 각 1단위씩 사용한다.
    '       [1]-1-5 우유군과 과일군에서 1단위씩 빼고 남은 단위수만큼 간식에 분배해서 보여준다.
    '       [1]-1-6 나머지 끼니는 해당 단위수로 구성한다.
    
    '   [1]-2 한식+대용식인 경우 - 4~9
    '       [1]-2-1 처방칼로리-대용식열량(1회 혹은 2회)과 감량열량 0 으로 식단번호와 해당 단위수를 불러온다.
    '       [1]-2-2 설정된 끼니는 대용식+우유 제공
    '       [1]-2-3 나머지 끼니는 설정된 끼니의 단위수만큼을 분배, 추가해서 보여준다.
    '       [1]-2-4 간식에도 대용식을 먹는 경우(4~6) 대용식+남은 우유군+과일군으로 보여준다.
    '   [1]-3 한식인 경우-0
    '       [1]-3-1 처방칼로리와 감량칼로리에 의해 식단번호와 이미지를 불러와서 보여준다.
    qrySelect = "SELECT EatTime, ReMealName FROM LHmeal WHERE CustomerNum=" & glngCustomerNum
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
    '식단정보설정이 된 경우
        mintConfig = Is_Null(rValue(0, 0), 0)
    Else
        mintConfig = 0
    End If
    
    If mintConfig < 4 Then
        intSet = WhatMealSetNum
        If intSet = 0 Then
            MsgBox "식사가이드를 보시려면 식사칼로리를 먼저 처방하십시오." & vbNewLine & "식사칼로리는 <처방/비만상담> 화면에서 하실 수 있습니다.", vbOKOnly + vbInformation
            For i = 0 To 5
                Set imgBreak(i).Picture = LoadPicture("")
                Set imgLunch(i).Picture = LoadPicture("")
                Set imgDinner(i).Picture = LoadPicture("")
            Next i
            Set imgSnack(0).Picture = LoadPicture("")
            Set imgSnack(1).Picture = LoadPicture("")
            Call imgTab_Click(0)
            For i = 0 To 3
                imgTab(i).Enabled = False
            Next i
            Exit Sub
        End If
    Else
    '한식+대용식인 경우 대용식의 칼로리 부터 불러오기
        qrySelect = "SELECT ChMealCalory FROM ChangeMeal INNER JOIN LHmeal "
        qrySelect = qrySelect & "ON ChangeMeal.ChMealName=LHmeal.ReMealName "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        Set clsSelect = New clsSelect
        rValue = clsSelect.Query(qrySelect)
        Set clsSelect = Nothing
        If Not IsNull(rValue) Then
            sngChMeal = Is_Null(rValue(0, 0), 0)
            '테스트용
            sngChMeal = 125
            If mintConfig = 4 Or mintConfig = 5 Or mintConfig = 6 Then
                sngChMeal = sngChMeal * 2
            End If
        End If
        sngChMeal = 125
        If mintConfig = 4 Or mintConfig = 5 Or mintConfig = 6 Then
            sngChMeal = sngChMeal * 2
        End If
        intSet = WhatMealSetNum(sngChMeal)
    End If
    
    Select Case mintConfig
        Case 1 To 3:
            Call LoadMealUnitSet(intSet)
            Call ShowMealTable_Bread
        Case 4 To 9:
            Call LoadMealUnitSet(intSet)
            Call ShowMealTable_ChMeal
        Case Else
            Call ShowMealTable
    End Select
    
    Call InitialChart
    Call ShowPortion
    Call ShowSnack
    Call ShowFruit
    Call imgTab_Click(0)
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
    
Private Sub LoadMealUnitSet(intSetNum As Integer)
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    
    qrySelect = "SELECT B_Grain, (B_FishLow + B_FishMid + B_FishHigh), "
    qrySelect = qrySelect & "L_Grain, (L_FishLow + L_FishMid + L_FishHigh), "
    qrySelect = qrySelect & "D_Grain, (D_FishLow + D_FishMid + D_FishHigh), "
    qrySelect = qrySelect & "S_Milk, S_Fruit "
    qrySelect = qrySelect & "From tblMealUnitSet "
    qrySelect = qrySelect & "Where xno=" & intSetNum
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        For i = 0 To 2
            sngGrain(i) = rValue((2 * i) + 0, 0)
            sngFishMeat(i) = rValue((2 * i) + 1, 0)
        Next i
        sngSnackMilk = rValue(6, 0)
        sngSnackFruit = rValue(7, 0)
    End If
End Sub

Private Sub ShowMealTable_Bread()
'   [1]-1 한식+빵/씨리얼인 경우 - 1,2,3
'       [1]-1-1 처방칼로리와 감량칼로리에 의해 식단번호와 해당 단위수를 불러온다(곡류군, 우유군, 과일군)

'       [1]-1-2 설정된 끼니의 곡류군 단위가 2 이상이면
'               해당 끼니의 곡류군 단위는 2로 하고, 나머지 끼니에 (아)35:(점)35:(저)30의 비율을 유지해 분배한다.
'       [1]-1-3 설정된 끼니의 곡류군 단위가 2 보다 작으면 주어진 단위수만큼 빵을 준다.
'       [1]-1-4 설정된 끼니에 우유군과 과일군을 각 1단위씩 사용한다.
'       [1]-1-5 우유군과 과일군에서 1단위씩 빼고 남은 단위수만큼 간식에 분배해서 보여준다.
'       [1]-1-6 나머지 끼니는 해당 단위수로 구성한다.
    Dim intBread As Integer, intRice(2) As Integer
    Dim intMain As Integer, intSub1 As Integer, intSub2 As Integer
    Dim i As Integer
    Dim sngRemain As Single
    Dim strImage As String
    
    For i = 0 To 5
        imgBreak(i).Visible = False
        imgLunch(i).Visible = False
        imgDinner(i).Visible = False
        imgDinner(i).Visible = False
    Next i
    
    For i = 0 To 3
        Set imgUnit1(i).Picture = LoadPicture("")
        Set imgUnit2(i).Picture = LoadPicture("")
        Set imgUnit3(i).Picture = LoadPicture("")
        Set imgUnit4(i).Picture = LoadPicture("")
        Set imgUnit5(i).Picture = LoadPicture("")
        imgUnit1(i).Visible = True
        imgUnit2(i).Visible = True
        imgUnit3(i).Visible = True
        imgUnit4(i).Visible = True
        imgUnit5(i).Visible = True
    Next i
    
    Select Case mintConfig
        Case 1: intBread = 0: intRice(0) = 1: intRice(1) = 2
                intMain = 35: intSub1 = 35: intSub2 = 30
        Case 2: intBread = 1: intRice(0) = 0: intRice(1) = 2
                intMain = 35: intSub1 = 35: intSub2 = 30
        Case 3: intBread = 2: intRice(0) = 0: intRice(1) = 1
                intMain = 30: intSub1 = 35: intSub2 = 35
    End Select
    
    If sngGrain(intBread) >= 2 Then
        sngRemain = sngGrain(intBread) - 2
        sngGrain(intBread) = 2
        sngGrain(intRice(0)) = sngGrain(intRice(0)) + (sngRemain * intSub1 / (intSub1 + intSub2))
        sngGrain(intRice(1)) = sngGrain(intRice(1)) + (sngRemain * intSub2 / (intSub1 + intSub2))
    End If
    sngFishMeat(intRice(0)) = sngFishMeat(intRice(0)) + (sngFishMeat(intBread) * intSub1 / (intSub1 + intSub2))
    sngFishMeat(intRice(1)) = sngFishMeat(intRice(1)) + (sngFishMeat(intBread) * intSub2 / (intSub1 + intSub2))
    
    '==================================================
    ' 1. 곡류군
    '==================================================
    Select Case Format(sngGrain(intBread), "#.#")
        Case Is <= 1.3: strImage = "0084_13.jpg"
        Case 1.4: strImage = "0084_14.jpg"
        Case 1.5: strImage = "0084_15.jpg"
        Case 1.6: strImage = "0084_16.jpg"
        Case 1.7: strImage = "0084_17.jpg"
        Case 1.8: strImage = "0084_18.jpg"
        Case 1.9: strImage = "0084_19.jpg"
        Case Else: strImage = "0084_20.jpg"
    End Select
    Set imgUnit1(intBread).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)
    For i = 0 To 1
        Select Case Format(sngGrain(intRice(i)), "#.#")
            Case Is <= 0.3: strImage = "rice_3.jpg"
            Case Is <= 0.5: strImage = "rice_5.jpg"
            Case Is <= 0.7: strImage = "rice_7.jpg"
            Case Is <= 1: strImage = "rice_10.jpg"
            Case Is <= 1.3: strImage = "rice_13.jpg"
            Case Is <= 1.7: strImage = "rice_17.jpg"
            Case Else: strImage = "rice_20.jpg"
        End Select
        Set imgUnit1(intRice(i)).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)
    Next i
    '==================================================
    ' 2. 어육류군
    '==================================================
    For i = 0 To 1
        Select Case Format(sngFishMeat(intRice(i)), "#.#")
            Case Is <= 0.5: strImage = "beef_5.jpg"
            Case Is <= 1: strImage = "beef_10.jpg"
            Case Is <= 1.5: strImage = "beef_15.jpg"
            Case Is <= 2: strImage = "beef_20.jpg"
            Case Is <= 2.5: strImage = "beef_25.jpg"
            Case Else: strImage = "beef_30.jpg"
        End Select
        Set imgUnit2(intRice(i)).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)
    Next i
    '==================================================
    ' 3. 채소군
    '==================================================
    For i = 0 To 1
        Set imgUnit3(intRice(i)).Picture = LoadPicture(App.Path & IMG_MEAL & "vegetable_1.jpg")
    Next i
    '==================================================
    ' 4. 우유군
    '==================================================
    Set imgUnit4(intBread).Picture = LoadPicture(App.Path & IMG_MEAL & "1166_10.jpg")   '빵 먹을때 우유 1단위
    sngSnackMilk = sngSnackMilk - 1
    If sngSnackMilk = 0 Then
        Set imgUnit4(3).Picture = LoadPicture("")
    Else
        Select Case Format(sngSnackMilk, "#.#")
            Case Is <= 0.5: strImage = "1166_5.jpg"
            Case Is <= 1: strImage = "1166_10.jpg"
            Case Is <= 1.5: strImage = "1166_15.jpg"
            Case Else: strImage = "1166_20.jpg"
        End Select
        Set imgUnit4(3).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)
    End If
    '==================================================
    ' 5. 과일군
    '==================================================
    Set imgUnit5(intBread).Picture = LoadPicture(App.Path & IMG_MEAL & "1297_10.jpg")
    sngSnackFruit = sngSnackFruit - 1
    If sngSnackFruit = 0 Then
        Set imgUnit5(3).Picture = LoadPicture("")
    Else
        Select Case Format(sngSnackFruit, "#.#")
            Case Is <= 0.5: strImage = "1297_5.jpg"
            Case Is <= 1: strImage = "1297_10.jpg"
            Case Is <= 1.5: strImage = "1297_15.jpg"
            Case Is <= 2: strImage = "1297_20.jpg"
            Case Else: strImage = "1297_30.jpg"
        End Select
        Set imgUnit5(3).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)
    End If
End Sub

Private Sub ShowMealTable_ChMeal()
'   [1]-2 한식+대용식인 경우 - 4~9
'       [1]-2-1 처방칼로리-대용식열량(1회 혹은 2회)과 감량열량 0 으로 식단번호와 해당 단위수를 불러온다.
'       [1]-2-2 설정된 끼니는 대용식+우유 제공
'       [1]-2-3 나머지 끼니는 설정된 끼니의 단위수만큼을 분배, 추가해서 보여준다.
'       [1]-2-4 간식에도 대용식을 먹는 경우(4~6) 대용식+남은 우유군+과일군으로 보여준다.
    Dim intChMeal As Integer, intRice(2) As Integer
    Dim intMain As Integer, intSub1 As Integer, intSub2 As Integer
    Dim i As Integer
    Dim sngRemain As Single
    Dim strImage As String
    
    For i = 0 To 5
        imgBreak(i).Visible = False
        imgLunch(i).Visible = False
        imgDinner(i).Visible = False
        imgDinner(i).Visible = False
    Next i
    
    For i = 0 To 3
        Set imgUnit1(i).Picture = LoadPicture("")
        Set imgUnit2(i).Picture = LoadPicture("")
        Set imgUnit3(i).Picture = LoadPicture("")
        Set imgUnit4(i).Picture = LoadPicture("")
        Set imgUnit5(i).Picture = LoadPicture("")
        imgUnit1(i).Visible = True
        imgUnit2(i).Visible = True
        imgUnit3(i).Visible = True
        imgUnit4(i).Visible = True
        imgUnit5(i).Visible = True
    Next i
    
    Select Case mintConfig
        Case 4, 7: intChMeal = 0: intRice(0) = 1: intRice(1) = 2
                intMain = 35: intSub1 = 35: intSub2 = 30
        Case 5, 8: intChMeal = 1: intRice(0) = 0: intRice(1) = 2
                intMain = 35: intSub1 = 35: intSub2 = 30
        Case 6, 9: intChMeal = 2: intRice(0) = 0: intRice(1) = 1
                intMain = 30: intSub1 = 35: intSub2 = 35
    End Select
    
    sngGrain(intRice(0)) = sngGrain(intRice(0)) + (sngGrain(intChMeal) * intSub1 / (intSub1 + intSub2))
    sngGrain(intRice(1)) = sngGrain(intRice(1)) + (sngGrain(intChMeal) * intSub2 / (intSub1 + intSub2))
    sngFishMeat(intRice(0)) = sngFishMeat(intRice(0)) + (sngFishMeat(intChMeal) * intSub1 / (intSub1 + intSub2))
    sngFishMeat(intRice(1)) = sngFishMeat(intRice(1)) + (sngFishMeat(intChMeal) * intSub2 / (intSub1 + intSub2))
    
    '==================================================
    ' 0. 대용식
    '==================================================
    Set imgUnit1(intChMeal).Picture = LoadPicture(App.Path & IMG_MEAL & "sub_1.jpg")
    If mintConfig = 4 Or mintConfig = 5 Or mintConfig = 6 Then
        Set imgUnit1(3).Picture = LoadPicture(App.Path & IMG_MEAL & "sub_1.jpg")
    End If
    '==================================================
    ' 1. 곡류군
    '==================================================
    For i = 0 To 1
        Select Case Format(sngGrain(intRice(i)), "#.#")
            Case Is <= 0.3: strImage = "rice_3.jpg"
            Case Is <= 0.5: strImage = "rice_5.jpg"
            Case Is <= 0.7: strImage = "rice_7.jpg"
            Case Is <= 1: strImage = "rice_10.jpg"
            Case Is <= 1.3: strImage = "rice_13.jpg"
            Case Is <= 1.7: strImage = "rice_17.jpg"
            Case Else: strImage = "rice_20.jpg"
        End Select
        Set imgUnit1(intRice(i)).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)
    Next i
    '==================================================
    ' 2. 어육류군
    '==================================================
    For i = 0 To 1
        Select Case Format(sngFishMeat(intRice(i)), "#.#")
            Case Is <= 0.5: strImage = "beef_5.jpg"
            Case Is <= 1: strImage = "beef_10.jpg"
            Case Is <= 1.5: strImage = "beef_15.jpg"
            Case Is <= 2: strImage = "beef_20.jpg"
            Case Is <= 2.5: strImage = "beef_25.jpg"
            Case Else: strImage = "beef_30.jpg"
        End Select
        Set imgUnit2(intRice(i)).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)
    Next i
    '==================================================
    ' 3. 채소군
    '==================================================
    For i = 0 To 1
        Set imgUnit3(intRice(i)).Picture = LoadPicture(App.Path & IMG_MEAL & "vegetable_1.jpg")
    Next i
    '==================================================
    ' 4. 우유군
    '==================================================
    Set imgUnit4(intChMeal).Picture = LoadPicture(App.Path & IMG_MEAL & "1166_10.jpg")   '빵 먹을때 우유 1단위
    sngSnackMilk = sngSnackMilk - 1
    If sngSnackMilk = 0 Then
        Set imgUnit4(3).Picture = LoadPicture("")
    Else
        Select Case Format(sngSnackMilk, "#.#")
            Case Is <= 0.5: strImage = "1166_5.jpg"
            Case Is <= 1: strImage = "1166_10.jpg"
            Case Is <= 1.5: strImage = "1166_15.jpg"
            Case Else: strImage = "1166_20.jpg"
        End Select
        Set imgUnit4(3).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)
    End If
    '==================================================
    ' 5. 과일군
    '==================================================
    Select Case Format(sngSnackFruit, "#.#")
        Case Is <= 0.5: strImage = "1297_5.jpg"
        Case Is <= 1: strImage = "1297_10.jpg"
        Case Is <= 1.5: strImage = "1297_15.jpg"
        Case Is <= 2: strImage = "1297_20.jpg"
        Case Else: strImage = "1297_30.jpg"
    End Select
    Set imgUnit5(3).Picture = LoadPicture(App.Path & IMG_MEAL & strImage)

End Sub

Private Sub ShowMealTable()
'아침,점심,저녁,간식 미리 짜여진 식단 이미지 불러와서 보여주기
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    
On Error GoTo ShowErr
    For i = 0 To 3
        imgUnit1(i).Visible = False
        imgUnit2(i).Visible = False
        imgUnit3(i).Visible = False
        imgUnit4(i).Visible = False
        imgUnit5(i).Visible = False
    Next i
    
    For i = 0 To 5
        Set imgBreak(i).Picture = LoadPicture("")
        Set imgLunch(i).Picture = LoadPicture("")
        Set imgDinner(i).Picture = LoadPicture("")
        Set imgSnack(i).Picture = LoadPicture("")
        imgBreak(i).Visible = True
        imgLunch(i).Visible = True
        imgDinner(i).Visible = True
        imgDinner(i).Visible = True
    Next i
    
    qrySelect = "SELECT T1_1, T1_2, T1_3, T1_4, T1_5, T1_6, "
    qrySelect = qrySelect & "T2_1, T2_2, T2_3, T2_4, T2_5, T2_6, "
    qrySelect = qrySelect & "T3_1, T3_2, T3_3, T3_4, T3_5, T3_6, "
    qrySelect = qrySelect & "T4_1, T4_2 "
    qrySelect = qrySelect & "FROM tblMealTable "
    '테스트 아직 몇번 불러올지 받아야함.
    qrySelect = qrySelect & "WHERE xno=" & intSet
    
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        For i = 0 To 5
            If Not IsNull(rValue(i, 0)) Then
                Set imgBreak(i).Picture = LoadPicture(App.Path & IMG_MEAL & Trim(rValue(i, 0)) & ".jpg")
            Else
                Set imgBreak(i).Picture = LoadPicture("")
            End If
            If Not IsNull(rValue(i + 6, 0)) Then
                Set imgLunch(i).Picture = LoadPicture(App.Path & IMG_MEAL & Trim(rValue(i + 6, 0)) & ".jpg")
            Else
                Set imgLunch(i).Picture = LoadPicture("")
            End If
            If Not IsNull(rValue(i + 12, 0)) Then
                Set imgDinner(i).Picture = LoadPicture(App.Path & IMG_MEAL & Trim(rValue(i + 12, 0)) & ".jpg")
            Else
                Set imgDinner(i).Picture = LoadPicture("")
            End If
        Next i
        If Not IsNull(rValue(18, 0)) Then
            Set imgSnack(0).Picture = LoadPicture(App.Path & IMG_MEAL & Trim(rValue(18, 0)) & ".jpg")
        Else
            Set imgSnack(0).Picture = LoadPicture("")
        End If
        If Not IsNull(rValue(19, 0)) Then
            Set imgSnack(1).Picture = LoadPicture(App.Path & IMG_MEAL & Trim(rValue(19, 0)) & ".jpg")
        Else
            Set imgSnack(1).Picture = LoadPicture("")
        End If
    End If
    
    Exit Sub
ShowErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "ShowMealTable", "frmCounsel_3", Err.Number, Err.Description
    MsgBox Err.Description
    Resume Next
End Sub

Private Sub imgSub_Click(Index As Integer)
    Dim i As Integer
    
    chtSnack.Visible = False
    lblExTime.Visible = False
    For i = 0 To 9
        lblSnack(i).Visible = False
    Next i
    For i = 0 To 19
        lblFruit(i).Visible = False
    Next i
    lblFMain.Visible = False
    lblFCal.Visible = False
    imgSub(0).Visible = True
    imgSub(1).Visible = True
    Select Case Index
        Case 0:
            If mintTab = 3 Then
                If mintConfig = 0 Then
                    Set Me.Picture = LoadPicture(App.Path & PATH03 & IMG_DINNER)
                Else
                    Set Me.Picture = LoadPicture(App.Path & PATH03UNIT & IMG_DINNER)
                End If
                
                For i = 0 To 5
                    lbl(i).Visible = True
                Next i
                Set imgSub(0).Picture = LoadPicture(App.Path & PATH03 & IMG_SUB1_ON)
                Set imgSub(1).Picture = LoadPicture(App.Path & PATH03 & IMG_SUB2_OFF)
            ElseIf mintTab = 4 Then
                If mintConfig = 0 Then
                    Set Me.Picture = LoadPicture(App.Path & PATH03 & IMG_SNACK)
                Else
                    Set Me.Picture = LoadPicture(App.Path & PATH03UNIT & IMG_SNACK)
                End If
                
                For i = 0 To 19
                    lblFruit(i).Visible = True
                Next i
                lblFMain.Visible = True
                lblFCal.Visible = True
                For i = 0 To 5
                    lbl(i).Visible = False
                Next i
                Set imgSub(0).Picture = LoadPicture(App.Path & PATH03 & IMG_SSUB1_ON)
                Set imgSub(1).Picture = LoadPicture(App.Path & PATH03 & IMG_SSUB2_OFF)
            End If
        Case 1:
            If mintTab = 3 Then
                If mintConfig = 0 Then
                    Set Me.Picture = LoadPicture(App.Path & PATH03 & IMG_DINNER2)
                Else
                    Set Me.Picture = LoadPicture(App.Path & PATH03UNIT & IMG_DINNER2)
                End If
                
                For i = 0 To 5
                    lbl(i).Visible = False
                Next i
                Set imgSub(0).Picture = LoadPicture(App.Path & PATH03 & IMG_SUB1_OFF)
                Set imgSub(1).Picture = LoadPicture(App.Path & PATH03 & IMG_SUB2_ON)
            ElseIf mintTab = 4 Then
                If mintConfig = 0 Then
                    Set Me.Picture = LoadPicture(App.Path & PATH03 & IMG_SNACK2)
                Else
                    Set Me.Picture = LoadPicture(App.Path & PATH03UNIT & IMG_SNACK2)
                End If
                
                chtSnack.Visible = True
                lblExTime.Visible = True
                For i = 0 To 9
                    lblSnack(i).Visible = True
                Next i
                For i = 0 To 5
                    lbl(i).Visible = False
                Next i
                Set imgSub(0).Picture = LoadPicture(App.Path & PATH03 & IMG_SSUB1_OFF)
                Set imgSub(1).Picture = LoadPicture(App.Path & PATH03 & IMG_SSUB2_ON)
            End If
    End Select
End Sub

Private Sub imgTab_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        imgTab(i).Enabled = True
    Next i
    For i = 0 To 5
        lbl(i).Visible = False
    Next i
    chtSnack.Visible = False
    lblExTime.Visible = False
    For i = 0 To 9
        lblSnack(i).Visible = False
    Next i
    For i = 0 To 19
        lblFruit(i).Visible = False
    Next i
    lblFMain.Visible = False
    lblFCal.Visible = False
    Select Case Index
        Case 0:
            mintTab = 1
            For i = 0 To 14
                lblPortion(i).Visible = False
            Next i
            imgSub(0).Visible = False
            imgSub(1).Visible = False
            If mintConfig = 0 Then
                Set Me.Picture = LoadPicture(App.Path & PATH03 & IMG_BREAK)
            Else
                Set Me.Picture = LoadPicture(App.Path & PATH03UNIT & IMG_BREAK)
            End If
            
            Set imgTab(0).Picture = LoadPicture(App.Path & PATH03 & IMG_BREAK_ON)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH03 & IMG_LUNCH_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH03 & IMG_DINNER_OFF)
            Set imgTab(3).Picture = LoadPicture(App.Path & PATH03 & IMG_SNACK_OFF)
        Case 1:
            mintTab = 2
            imgSub(0).Visible = False
            imgSub(1).Visible = False
            If mintConfig = 0 Then
                Set Me.Picture = LoadPicture(App.Path & PATH03 & IMG_LUNCH)
            Else
                Set Me.Picture = LoadPicture(App.Path & PATH03UNIT & IMG_LUNCH)
            End If
            
            Set imgTab(0).Picture = LoadPicture(App.Path & PATH03 & IMG_BREAK_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH03 & IMG_LUNCH_ON)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH03 & IMG_DINNER_OFF)
            Set imgTab(3).Picture = LoadPicture(App.Path & PATH03 & IMG_SNACK_OFF)
            For i = 0 To 14
                lblPortion(i).Visible = True
            Next i
        Case 2:
            mintTab = 3
            Call imgSub_Click(0)
            
            For i = 0 To 14
                lblPortion(i).Visible = False
            Next i
            Set imgTab(0).Picture = LoadPicture(App.Path & PATH03 & IMG_BREAK_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH03 & IMG_LUNCH_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH03 & IMG_DINNER_ON)
            Set imgTab(3).Picture = LoadPicture(App.Path & PATH03 & IMG_SNACK_OFF)
        Case 3:
            mintTab = 4
            Call imgSub_Click(0)
                        
            For i = 0 To 14
                lblPortion(i).Visible = False
            Next i
            Set imgTab(0).Picture = LoadPicture(App.Path & PATH03 & IMG_BREAK_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH03 & IMG_LUNCH_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH03 & IMG_DINNER_OFF)
            Set imgTab(3).Picture = LoadPicture(App.Path & PATH03 & IMG_SNACK_ON)
    End Select
End Sub

Private Sub ShowPortion()
    Dim intLunch As Integer, intDinner As Integer
    Dim intGab As Integer
    Dim i As Integer

    intLunch = Cal_Calory(2)
    intDinner = Cal_Calory(3)
        
    If intLunch = 0 Then
        For i = 0 To 14
            lblPortion(i).Caption = ""
'            lblPortion(i).Visible = False
        Next i
        If intDinner = 0 Then
            lbl(3).Caption = ""
            lbl(4).Caption = ""
            lbl(5).Caption = ""
        End If
        Exit Sub
    End If

    '점심 외식 분량 ////////////////////////////////////////////////////////////////////////////
    lblPortion(0).Caption = Format((intLunch / 250), "0.0인분")
    lblPortion(1).Caption = Format((intLunch / 300), "0.0인분")
    lblPortion(2).Caption = Format((intLunch / 350), "0.0인분")
    lblPortion(3).Caption = Format((intLunch / 400), "0.0인분")
    lblPortion(4).Caption = Format((intLunch / 450), "0.0인분")
    lblPortion(5).Caption = Format((intLunch / 500), "0.0인분")
    lblPortion(6).Caption = Format((intLunch / 550), "0.0인분")
    lblPortion(7).Caption = Format((intLunch / 600), "0.0인분")
    lblPortion(8).Caption = Format((intLunch / 650), "0.0인분")
    lblPortion(9).Caption = Format((intLunch / 700), "0.0인분")
    lblPortion(10).Caption = Format((intLunch / 800), "0.0 인분")
    lblPortion(11).Caption = Format((intLunch / 850), "0.0 인분")
    lblPortion(12).Caption = Format((intLunch / 900), "0.0 인분")
    lblPortion(13).Caption = Format((intLunch / 950), "0.0 인분")
    lblPortion(14).Caption = Format((intLunch / 1000), "0.0 인분")
    
    '저녁칼로리 술,안주 열량 대비 ///////////////////////////////////////////////////////
    ' 1,020 kcal
    
    intGab = 1020 - intDinner
    If intGab > 0 Then
        lbl(3).Caption = intGab & "kcal" & vbNewLine & "초과"
    ElseIf intGab < 0 Then
        lbl(3).Caption = Abs(intGab) & "kcal" & vbNewLine & "부족"
    Else
        lbl(3).Caption = "적당량"
    End If
    ' 915 kcal
    intGab = 915 - intDinner
    If intGab > 0 Then
        lbl(4).Caption = intGab & "kcal" & vbNewLine & "초과"
    ElseIf intGab < 0 Then
        lbl(4).Caption = Abs(intGab) & "kcal" & vbNewLine & "부족"
    Else
        lbl(4).Caption = "적당량"
    End If
    ' 610 kcal
    intGab = 610 - intDinner
    If intGab > 0 Then
        lbl(5).Caption = intGab & "kcal" & vbNewLine & "초과"
    ElseIf intGab < 0 Then
        lbl(5).Caption = Abs(intGab) & "kcal" & vbNewLine & "부족"
    Else
        lbl(5).Caption = "적당량"
    End If
End Sub

Private Sub ShowFruit()
'간식으로 제공되는 과일의 대체과일 갯수(단위수) 보이기
    Dim sngFruit(20) As Single, strName(20) As String, strUnit(20) As String
    Dim sngCount As Single, strCount As String, sngSel As Single
    Dim i As Integer
    
    If intSet = 0 Then
        For i = 0 To 19
            lblFruit(i).Caption = ""
        Next i
        lblFMain.Caption = ""
        lblFCal.Caption = ""
        Exit Sub
    End If
    
    '0) 초기화 음식명, 단위수, 세는 단위 셋팅
    strName(0) = "금귤"
    strName(1) = "방울토마토"
    strName(2) = "딸기"
    strName(3) = "레몬"
    strName(4) = "천도복숭아"
    strName(5) = "자두"
    strName(6) = "토마토"
    strName(7) = "키위"
    strName(8) = "귤"
    strName(9) = "수박"
    strName(10) = "델라웨어"
    strName(11) = "오렌지"
    strName(12) = "바나나"
    strName(13) = "단감"
    strName(14) = "황도복숭아"
    strName(15) = "참외"
    strName(16) = "자몽"
    strName(17) = "포도"
    strName(18) = "사과"
    strName(19) = "배"
    
    ' 금귤/방울토마토/딸기/레몬/복숭아(천도)
    sngFruit(0) = 0.14
    sngFruit(1) = 0.05
    sngFruit(2) = 0.1
    sngFruit(3) = 0.3
    sngFruit(4) = 0.5
    ' 자두/토마토/키위/귤/수박
    For i = 5 To 9
        sngFruit(i) = 1
    Next i
    ' 델라웨어/오렌지/바나나/단감/복숭아(황도)
    sngFruit(10) = 1.36
'    sngFruit(11) = 1   '2004.07.08 변경
    For i = 11 To 16
        sngFruit(i) = 2
    Next i
    ' 참외/자몽/포도/사과/배
    sngFruit(17) = 2.8
    sngFruit(18) = 3
    sngFruit(19) = 4
    
    For i = 0 To 19
        strUnit(i) = "개"
    Next i
    strUnit(9) = "쪽"     '수박
    strUnit(10) = "송이"
    strUnit(17) = "송이"
    
    '1) 현재 간식으로 주어진 과일의 이름 , 단위수, 칼로리 알아내기
    '   --- 테스트용 2004.03.04
    Dim qrySelect As String, rValue As Variant
    Dim strFruit As String, strUnit2 As String
    
    qrySelect = "SELECT S_Fruit, T4_2Name "
    qrySelect = qrySelect & "FROM tblMealUnitSet INNER JOIN tblMealTable "
    qrySelect = qrySelect & "ON tblMealUnitSet.xno=tblMealTable.xno "
    qrySelect = qrySelect & "WHERE tblMealUnitSet.xno=" & intSet
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        sngSel = rValue(0, 0)
        strFruit = Trim(Is_Null(rValue(1, 0), ""))
    Else
        sngSel = 1
        strFruit = ""
    End If
    lblFMain.Caption = strFruit
    lblFCal.Caption = sngSel * 50
    
    '2) 갯수계산법=변경할과일단위수/1개단위수
    For i = 0 To 17
        sngCount = sngSel / sngFruit(i)
    
    '3) 소숫점처리
    '   - 0.9미만인 경우 : 소수점을 분수로 변경
    '   - 0.9, 1 초과한 경우 : 소수로 표현
        sngCount = Format(sngCount, "#.#")
        Select Case sngCount
            Case 0.1: strCount = "1/10"
            Case 0.2: strCount = "1/5"
            Case 0.3: strCount = "2/7"
            Case 0.4: strCount = "2/5"
            Case 0.5: strCount = "1/2"
            Case 0.6: strCount = "3/5"
            Case 0.7: strCount = "2/3"
            Case 0.8: strCount = "4/5"
            Case 0.9: strCount = "0.9"
            Case Else
                strCount = Format(sngCount, "#.#")
                If Right(strCount, 1) = "." Then
                    strCount = Int(sngCount)
                End If
        End Select
        lblFruit(i).Caption = strName(i) & vbNewLine & strCount & strUnit(i)
    Next i
    
    '사과와 배는 다른 갯수세는법 사용
'    사과 : 0.33 = 1/3개, 0.67 = 2/3개, 1 = 1개
'    배 : 0.25 = 1/4개, 0.5 = 1/2개, 0.75 = 3/4개, 1 = 1개
    sngCount = sngSel / sngFruit(18)
    sngCount = Format(sngCount, "#.##")
    Select Case sngCount
        Case 0.33: strCount = "1/3"
        Case 0.67: strCount = "2/3"
        Case 1: strCount = "1"
        Case Else
            strCount = Format(sngCount, "#.#")
            If Right(strCount, 1) = "." Then
                strCount = Int(sngCount)
            End If
    End Select
    lblFruit(18).Caption = strName(18) & vbNewLine & strCount & strUnit(18)
    
    sngCount = sngSel / sngFruit(19)
    sngCount = Format(sngCount, "#.00")
    Select Case sngCount
        Case 0.25: strCount = "1/4"
        Case 0.5: strCount = "1/2"
        Case 0.75: strCount = "3/4"
        Case 1: strCount = "1"
        Case Else
            strCount = Format(sngCount, "#.#")
            If Right(strCount, 1) = "." Then
                strCount = Int(sngCount)
            End If
    End Select
    lblFruit(19).Caption = strName(19) & vbNewLine & strCount & strUnit(19)
End Sub

Private Sub ShowSnack()
'기호간식의 열량 차트로 보이기
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    Dim colCalory As New Collection
    Dim cfxArray As Object
    
    Set clsSelect = New clsSelect
    Set cfxArray = CreateObject("cfxdata.array")
    
    qrySelect = "SELECT SnackCode, SnackName, Calory, Serving FROM CustomerInfo RIGHT JOIN tblSnack "
    qrySelect = qrySelect & "ON CustomerInfo.Snack=tblSnack.GroupCode "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND tblSnack.UseYn='Y';"
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        If UBound(rValue, 2) >= 9 Then
            For i = 0 To 9
                lblSnack(9 - i).Caption = Trim(rValue(1, i)) & Trim(rValue(3, i))
                colCalory.Add rValue(2, i)
                intSnack(i) = rValue(2, i)
            Next i
        Else
            For i = 0 To UBound(rValue, 2)
                lblSnack(9 - i).Caption = Trim(rValue(1, i)) & Trim(rValue(3, i))
                colCalory.Add rValue(2, i)
                intSnack(i) = rValue(2, i)
            Next i
        End If
        cfxArray.AddArray colCalory
        
        chtSnack.GetExternalData cfxArray
    Else
        chtSnack.ClearData CD_VALUES
    End If
    Call lblSnack_Click(0)
    
    Set colCalory = Nothing
    Set clsSelect = Nothing
End Sub

'최종 처방칼로리에서 끼니별 칼로리를 구해서 반환하는 함수
'1: 아침/ 2: 점심 / 3: 저녁
Private Function Cal_Calory(intTime As Integer) As Integer
    Dim qrySelect As String, rValue As Variant
    Dim TotalUnit(8) As Single
    Dim TimeCalory(4) As Single
    Dim GroupUnit(3, 7) As Single  '식사시간, 군'
    Dim i, j, cal, t As Integer, mintBase As Integer
    
    Dim sngCalory As Single, intAct As Integer
    '점심칼로리를 구하기 위해 TreatCalory, ActDegree를 알아냄
    Set clsSelect = New clsSelect
    
    If intSet = 0 Then
        Exit Function
    End If

    If intTime = 2 Then
        qrySelect = "SELECT L_Grain, L_FishLow, L_FishMid, L_FishHigh, L_Veg, L_Fat "
    ElseIf intTime = 3 Then
        qrySelect = "SELECT D_Grain, D_FishLow, D_FishMid, D_FishHigh, D_Veg, D_Fat "
    End If
    qrySelect = qrySelect & "FROM tblMealUnitSet WHERE xno=" & intSet
    
    rValue = clsSelect.Query(qrySelect)
    
    
    TotalUnit(0) = rValue(0, 0)  '곡류
    TotalUnit(1) = rValue(1, 0)  '어육류 저지방
    TotalUnit(2) = rValue(2, 0)  '어육류 중지방
    TotalUnit(3) = rValue(4, 0)  '채소
    TotalUnit(4) = rValue(5, 0)  '지방
    
    Erase rValue
        
    '각 끼니 칼로리
    mintBase = Round(TotalUnit(0) * 100 + TotalUnit(1) * 50 + TotalUnit(2) * 75 + TotalUnit(3) * 20 + TotalUnit(4) * 45)
    Set clsSelect = Nothing
    
    Cal_Calory = mintBase
End Function

Private Sub InitialChart()
    With chtSnack
    .Gallery = GANTT
    .Chart3D = False
    .Stacked = CHART_NOSTACKED
    .Border = False
    .RGBBk = &HEFEFEF
    .BorderStyle = BORDER_NONE

    ' Layout Settings
    .LegendBox = False
    .SerLegBox = False
    .ToolBar = False
    .Title(CHART_TOPTIT) = ""
    .PointLabelAlign = LA_LEFT + LA_BOTTOM
    
    .TopGap = 0
    .BottomGap = 0
    .LeftGap = 0
    .RightGap = 0
    
    .Axis(AXIS_Y).Max = 800
    .Axis(AXIS_Y).Min = 0
   
    .AllowDrag = False
    .AllowEdit = False
    .AllowResize = False
    End With
End Sub

Private Sub lblSnack_Click(Index As Integer)
    Dim intRunTime As Integer, sngWeight As Single
    
    sngWeight = WhatWeight(glngCustomerNum)
    intRunTime = (intSnack(9 - Index) / (sngWeight * 0.16))
    lblExTime.Caption = lblSnack(Index).Caption & " = "
    lblExTime.Caption = lblExTime.Caption & intSnack(9 - Index) & "kcal = 달리기 "
    lblExTime.Caption = lblExTime.Caption & intRunTime & "분"
End Sub
