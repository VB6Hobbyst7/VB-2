VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCounsel_1 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   13515
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboChMealCode 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "frmCounsel_1.frx":0000
      Left            =   10365
      List            =   "frmCounsel_1.frx":0007
      Style           =   2  '드롭다운 목록
      TabIndex        =   68
      Top             =   4590
      Width           =   1650
   End
   Begin VB.ComboBox cboChMealTime 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "frmCounsel_1.frx":0017
      Left            =   8865
      List            =   "frmCounsel_1.frx":002D
      Style           =   2  '드롭다운 목록
      TabIndex        =   67
      Top             =   4590
      Width           =   1485
   End
   Begin VB.CheckBox chkChMeal 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "식사처방"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7890
      TabIndex        =   66
      Top             =   4620
      Width           =   960
   End
   Begin VB.CheckBox chkUserCalory 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      Caption         =   "식단열량직접입력하기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7890
      TabIndex        =   60
      Top             =   4365
      Width           =   1965
   End
   Begin VB.TextBox txtUserCalory 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   9870
      MaxLength       =   4
      TabIndex        =   58
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  '평면
      Height          =   2415
      Left            =   7890
      MultiLine       =   -1  'True
      TabIndex        =   57
      Text            =   "frmCounsel_1.frx":0064
      Top             =   5670
      Width           =   3255
   End
   Begin VB.TextBox txtNotice 
      Appearance      =   0  '평면
      Height          =   2415
      Left            =   7890
      MultiLine       =   -1  'True
      TabIndex        =   56
      Text            =   "frmCounsel_1.frx":006A
      Top             =   5670
      Width           =   3255
   End
   Begin VB.ComboBox cmbProgram 
      Height          =   300
      Left            =   9210
      TabIndex        =   21
      Text            =   "cmbProgram"
      ToolTipText     =   "환자에게 처방할 치료프로그램을 선택합니다."
      Top             =   1500
      Width           =   1485
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "신체계측"
      Height          =   315
      Index           =   9
      Left            =   6210
      TabIndex        =   20
      Top             =   7560
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "신체계측"
      Height          =   315
      Index           =   8
      Left            =   6210
      TabIndex        =   19
      Top             =   7260
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "신체계측"
      Height          =   315
      Index           =   7
      Left            =   6210
      TabIndex        =   18
      Top             =   6960
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "신체계측"
      Height          =   315
      Index           =   6
      Left            =   6210
      TabIndex        =   17
      Top             =   6660
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "신체계측"
      Height          =   315
      Index           =   5
      Left            =   6210
      TabIndex        =   16
      Top             =   6360
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "신체계측"
      Height          =   315
      Index           =   4
      Left            =   6210
      TabIndex        =   15
      Top             =   6060
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "R  M  R"
      Height          =   315
      Index           =   3
      Left            =   6210
      TabIndex        =   14
      Top             =   5760
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "W  H  R"
      Height          =   315
      Index           =   2
      Left            =   6210
      TabIndex        =   13
      Top             =   5460
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "허리둘레"
      Height          =   315
      Index           =   1
      Left            =   6210
      TabIndex        =   12
      Top             =   5160
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "체      중"
      Height          =   315
      Index           =   0
      Left            =   6210
      TabIndex        =   11
      Top             =   4860
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdTable 
      Height          =   3345
      Left            =   750
      TabIndex        =   9
      Top             =   4830
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5900
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin ChartfxLibCtl.ChartFX Chart 
      Height          =   3345
      Left            =   750
      TabIndex        =   10
      Top             =   4860
      Width           =   5475
      _cx             =   9657
      _cy             =   5900
      Build           =   20
      TypeMask        =   109576194
      Axis(0).Max     =   90
      Axis(2).Format  =   5
      Axis(2).Format  =   5
      RGB2DBk         =   16777215
      nColors         =   16
      Colors          =   "frmCounsel_1.frx":0070
      nSer            =   1
      NumSer          =   1
      _Data_          =   "frmCounsel_1.frx":0110
   End
   Begin FPSpread.vaSpread sprTreat 
      Height          =   1875
      Left            =   7890
      TabIndex        =   1
      Top             =   6210
      Width           =   1650
      _Version        =   196608
      _ExtentX        =   2910
      _ExtentY        =   3307
      _StockProps     =   64
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmCounsel_1.frx":021D
   End
   Begin FPSpread.vaSpread sprPrint 
      Height          =   1875
      Left            =   9660
      TabIndex        =   0
      Top             =   6210
      Width           =   1455
      _Version        =   196608
      _ExtentX        =   2566
      _ExtentY        =   3307
      _StockProps     =   64
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmCounsel_1.frx":03A5
   End
   Begin VB.Image imgPreTreat 
      Height          =   330
      Left            =   9765
      Picture         =   "frmCounsel_1.frx":052D
      ToolTipText     =   "가장 최근의 이전 처방값을 불러와서 표시해 줍니다."
      Top             =   1815
      Width           =   1185
   End
   Begin VB.Image imgModifyTreat 
      Height          =   330
      Left            =   11010
      Picture         =   "frmCounsel_1.frx":0C2F
      ToolTipText     =   "현재 조회하고 있는 처방을 수정할 수 있도록 합니다."
      Top             =   1815
      Width           =   1185
   End
   Begin VB.Image imgNew 
      Height          =   795
      Left            =   11370
      Picture         =   "frmCounsel_1.frx":12CE
      ToolTipText     =   "현재 작업날짜의 처방을 새로 입력할 수 있게 합니다."
      Top             =   4920
      Width           =   795
   End
   Begin VB.Label lblDispContent 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "조회중"
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
      Left            =   10185
      TabIndex        =   65
      Top             =   1065
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "일자 처방내역"
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
      Left            =   8895
      TabIndex        =   64
      Top             =   1065
      Width           =   1245
   End
   Begin VB.Label lblDispDate 
      BackStyle       =   0  '투명
      Caption         =   "2005-01-01"
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
      Left            =   7785
      TabIndex        =   63
      Top             =   1065
      Width           =   1125
   End
   Begin VB.Image imgValuation 
      Height          =   285
      Left            =   9720
      Picture         =   "frmCounsel_1.frx":199F
      Top             =   4935
      Width           =   1185
   End
   Begin VB.Label lblTreatCal 
      BackStyle       =   0  '투명
      Caption         =   "1,500 kcal"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   9600
      TabIndex        =   62
      Top             =   4155
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "처방한 식단열량 = "
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
      Left            =   8160
      TabIndex        =   61
      Top             =   4155
      Width           =   1395
   End
   Begin VB.Image imgSelectEx 
      Height          =   330
      Left            =   10800
      Picture         =   "frmCounsel_1.frx":1F49
      ToolTipText     =   "운동종목을 선택합니다."
      Top             =   4170
      Width           =   1185
   End
   Begin VB.Image TopImage 
      Height          =   960
      Left            =   -30
      Picture         =   "frmCounsel_1.frx":2A7F
      Top             =   50
      Width           =   13140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "kcal"
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
      Left            =   10350
      TabIndex        =   59
      Top             =   4365
      Width           =   405
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '단색
      Height          =   810
      Left            =   7800
      Shape           =   4  '둥근 사각형
      Top             =   4110
      Width           =   4305
   End
   Begin VB.Image imgTreat 
      Height          =   330
      Left            =   6240
      Picture         =   "frmCounsel_1.frx":4605
      ToolTipText     =   "환자의 치료이력을 달력형태로 조회합니다."
      Top             =   7860
      Width           =   1170
   End
   Begin VB.Image imgSelection 
      Height          =   285
      Left            =   6210
      Top             =   7890
      Width           =   1245
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "31 %"
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
      Left            =   2010
      TabIndex        =   55
      Top             =   3330
      Width           =   600
   End
   Begin VB.Label lblAnaerobic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "100분"
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
      Left            =   10590
      TabIndex        =   54
      Top             =   3870
      Width           =   615
   End
   Begin VB.Label lblAnaerobic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "100분"
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
      Left            =   10590
      TabIndex        =   53
      Top             =   3540
      Width           =   615
   End
   Begin VB.Label lblAnaerobic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "100분"
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
      Left            =   10590
      TabIndex        =   52
      Top             =   3210
      Width           =   615
   End
   Begin VB.Label lblAnaerobic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "100분"
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
      Left            =   10590
      TabIndex        =   51
      Top             =   2910
      Width           =   615
   End
   Begin VB.Label lblAerobic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "47분"
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
      Left            =   9870
      TabIndex        =   50
      Top             =   3870
      Width           =   675
   End
   Begin VB.Label lblAerobic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "47분"
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
      Left            =   9870
      TabIndex        =   49
      Top             =   3540
      Width           =   675
   End
   Begin VB.Label lblAerobic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "47분"
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
      Left            =   9870
      TabIndex        =   48
      Top             =   3210
      Width           =   675
   End
   Begin VB.Label lblAerobic 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "47분"
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
      Left            =   9870
      TabIndex        =   47
      Top             =   2910
      Width           =   675
   End
   Begin VB.Image imgExCal 
      Height          =   300
      Index           =   3
      Left            =   9180
      Picture         =   "frmCounsel_1.frx":524B
      ToolTipText     =   "운동칼로리를 입력합니다."
      Top             =   3795
      Width           =   645
   End
   Begin VB.Image imgExCal 
      Height          =   300
      Index           =   2
      Left            =   9180
      Picture         =   "frmCounsel_1.frx":588B
      ToolTipText     =   "운동칼로리를 입력합니다."
      Top             =   3480
      Width           =   645
   End
   Begin VB.Image imgExCal 
      Height          =   300
      Index           =   1
      Left            =   9180
      Picture         =   "frmCounsel_1.frx":5E98
      ToolTipText     =   "운동칼로리를 입력합니다."
      Top             =   3165
      Width           =   645
   End
   Begin VB.Image imgExCal 
      Height          =   300
      Index           =   0
      Left            =   9180
      Picture         =   "frmCounsel_1.frx":64D5
      ToolTipText     =   "운동칼로리를 입력합니다."
      Top             =   2850
      Width           =   645
   End
   Begin VB.Image imgExDay 
      Height          =   300
      Index           =   3
      Left            =   8730
      Picture         =   "frmCounsel_1.frx":6B07
      ToolTipText     =   "운동일수를 입력합니다."
      Top             =   3795
      Width           =   435
   End
   Begin VB.Image imgExDay 
      Height          =   300
      Index           =   2
      Left            =   8730
      Picture         =   "frmCounsel_1.frx":6FF6
      ToolTipText     =   "운동일수를 입력합니다."
      Top             =   3480
      Width           =   435
   End
   Begin VB.Image imgExDay 
      Height          =   300
      Index           =   1
      Left            =   8730
      Picture         =   "frmCounsel_1.frx":74E1
      ToolTipText     =   "운동일수를 입력합니다."
      Top             =   3165
      Width           =   435
   End
   Begin VB.Image imgExDay 
      Height          =   300
      Index           =   0
      Left            =   8730
      Picture         =   "frmCounsel_1.frx":79D8
      ToolTipText     =   "운동일수를 입력합니다."
      Top             =   2850
      Width           =   435
   End
   Begin VB.Label lblLossWeight 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2.2 kg/월"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   11250
      TabIndex        =   46
      ToolTipText     =   "처방의 효과로 줄어들 감량의 총 양입니다."
      Top             =   3390
      Width           =   825
   End
   Begin VB.Image imgPrint 
      Height          =   780
      Left            =   11370
      Picture         =   "frmCounsel_1.frx":7ECC
      ToolTipText     =   "비만도 평가 결과지를 출력합니다."
      Top             =   7350
      Width           =   750
   End
   Begin VB.Image imgDel 
      Height          =   795
      Left            =   11370
      Picture         =   "frmCounsel_1.frx":8F57
      ToolTipText     =   "현재 조회중인 처방을 삭제합니다."
      Top             =   6540
      Width           =   795
   End
   Begin VB.Image imgSave 
      Height          =   795
      Left            =   11370
      Picture         =   "frmCounsel_1.frx":9EAA
      ToolTipText     =   "현재 입력중인 처방을 저장합니다."
      Top             =   5730
      Width           =   795
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   2
      Left            =   10230
      Picture         =   "frmCounsel_1.frx":AD67
      ToolTipText     =   "현재 처방의 간단한 메모를 조회/입력/수정하도록 합니다."
      Top             =   5250
      Width           =   885
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   1
      Left            =   9360
      Picture         =   "frmCounsel_1.frx":B412
      ToolTipText     =   "환자의 누적된 Notice를 보여줍니다."
      Top             =   5250
      Width           =   885
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   0
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":BB20
      Top             =   5250
      Width           =   1560
   End
   Begin VB.Image imgGoReserve 
      Height          =   330
      Left            =   11010
      Picture         =   "frmCounsel_1.frx":C94B
      ToolTipText     =   "예약을 변경합니다."
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   3
      Left            =   5940
      Top             =   2970
      Width           =   1365
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   2
      Left            =   5940
      Top             =   2550
      Width           =   1365
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   1
      Left            =   5940
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   0
      Left            =   5940
      Top             =   1740
      Width           =   1365
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
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
      Left            =   4620
      TabIndex        =   45
      Top             =   3930
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
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
      Left            =   4620
      TabIndex        =   44
      Top             =   3630
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
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
      Left            =   4620
      TabIndex        =   43
      Top             =   3330
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
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
      Left            =   4620
      TabIndex        =   42
      Top             =   3030
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
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
      Left            =   4620
      TabIndex        =   41
      Top             =   2730
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
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
      Left            =   4620
      TabIndex        =   40
      Top             =   2430
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
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
      Left            =   4620
      TabIndex        =   39
      Top             =   2130
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "59 kg"
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
      Left            =   4620
      TabIndex        =   38
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   37
      Top             =   3930
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   36
      Top             =   3630
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   35
      Top             =   3330
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   34
      Top             =   3030
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   33
      Top             =   2730
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   32
      Top             =   2430
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   31
      Top             =   2130
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   30
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3030
      TabIndex        =   29
      Top             =   3930
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3030
      TabIndex        =   28
      Top             =   3630
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3030
      TabIndex        =   27
      Top             =   3330
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3030
      TabIndex        =   26
      Top             =   3030
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3030
      TabIndex        =   25
      Top             =   2730
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3030
      TabIndex        =   24
      Top             =   2430
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3030
      TabIndex        =   23
      Top             =   2130
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '투명
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3030
      TabIndex        =   22
      Top             =   1800
      Width           =   585
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   7
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":D46F
      Top             =   3975
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   6
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":D7D4
      Top             =   3675
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   5
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":DB39
      Top             =   3375
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   4
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":DE9E
      Top             =   3060
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   3
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":E203
      Top             =   2745
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   2
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":E568
      Top             =   2445
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   1
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":E8CD
      Top             =   2160
      Width           =   150
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   5
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":EC32
      ToolTipText     =   "처방칼로리를 입력합니다."
      Top             =   3795
      Width           =   900
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   4
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":F309
      ToolTipText     =   "처방칼로리를 입력합니다."
      Top             =   3480
      Width           =   900
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   3
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":F98A
      ToolTipText     =   "처방칼로리를 입력합니다."
      Top             =   3165
      Width           =   900
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   2
      Left            =   12450
      Picture         =   "frmCounsel_1.frx":10056
      Top             =   2850
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   1
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":106EC
      ToolTipText     =   "처방칼로리를 입력합니다."
      Top             =   2850
      Width           =   900
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   0
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":10D7E
      Top             =   1830
      Width           =   150
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   0
      Left            =   12450
      Picture         =   "frmCounsel_1.frx":110E3
      Top             =   2490
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgViewHistory 
      Height          =   345
      Index           =   2
      Left            =   3840
      Picture         =   "frmCounsel_1.frx":1176B
      ToolTipText     =   "환자의 치료내역을 리스트로 보여줍니다."
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Image imgViewHistory 
      Height          =   345
      Index           =   1
      Left            =   2250
      Picture         =   "frmCounsel_1.frx":12113
      ToolTipText     =   "환자의 검사결과에 대한 변화도를 그래프로 보여줍니다."
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Image imgViewHistory 
      Height          =   345
      Index           =   0
      Left            =   690
      Picture         =   "frmCounsel_1.frx":12B4A
      ToolTipText     =   "환자의 검사결과에 대한 변화도를 표로 보여줍니다."
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "65 kg"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "87 cm"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   2430
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "0.93"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   2730
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "120 %"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "27"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   3030
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1,280"
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
      Left            =   2010
      TabIndex        =   3
      Top             =   3630
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2,400"
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
      Left            =   2010
      TabIndex        =   2
      Top             =   3930
      Width           =   600
   End
End
Attribute VB_Name = "frmCounsel_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================
' 모 듈 명  : frmCounsel_1
' 작성일시  : 2005-01-28 08:45
' 작 성 자  : 류진선
' 내    용  : 환자의 현재상태(ObesityRecord)와 치료내역/변화도를 확인할수 있고
'             치료의 내용을 조회 수정할 수 있다.
'=======================================================================================

'다이어트 열량 구하기 DietCal (O)
'*** 폼 로드시
'1) 환경설정에서 불러온다. 직접 입력해야 하는 값들(그중에서 필수입력값들)
'*** 저장
'3) 저장한다.(BodyData, BioChem, Treat, TreatData, TreatPrint

Option Explicit
Private Const IMG_ME As String = "\Back\Counsel\01\Me\" '식사칼로리관련 이미지 폴더
Private Const IMG_EX As String = "\Back\Counsel\01\Ex\" '운동처방관련 이미지 폴더
'+---------------------------------------------------------------------------------+
'| 상담 > 처방/비만도상담 > 식사감량칼로리
'+---------------------------------------------------------------------------------+
Private Const IMG_M0Y As String = "\Back\Counsel\01\Me\125-green.jpg"
Private Const IMG_M0N As String = "\Back\Counsel\01\Me\125-gray.jpg"
Private Const IMG_M1Y As String = "\Back\Counsel\01\Me\250-green.jpg"
Private Const IMG_M1N As String = "\Back\Counsel\01\Me\250-gray.jpg"
Private Const IMG_M2Y As String = "\Back\Counsel\01\Me\375-green.jpg"
Private Const IMG_M2N As String = "\Back\Counsel\01\Me\375-gray.jpg"
Private Const IMG_M3Y As String = "\Back\Counsel\01\Me\500-green.jpg"
Private Const IMG_M3N As String = "\Back\Counsel\01\Me\500-gray.jpg"
Private Const IMG_M4Y As String = "\Back\Counsel\01\Me\750-green.jpg"
Private Const IMG_M4N As String = "\Back\Counsel\01\Me\750-gray.jpg"
Private Const IMG_M5Y As String = "\Back\Counsel\01\Me\1000-green.jpg"
Private Const IMG_M5N As String = "\Back\Counsel\01\Me\1000-gray.jpg"
'+---------------------------------------------------------------------------------+
'| 상담 > 처방/비만도상담 > 운동일수
'+---------------------------------------------------------------------------------+
Private Const IMG_EX3DY As String = "\Back\Counsel\01\Ex\3일-green.jpg"
Private Const IMG_EX3DN As String = "\Back\Counsel\01\Ex\3일-gray.jpg"
Private Const IMG_EX4DY As String = "\Back\Counsel\01\Ex\4일-green.jpg"
Private Const IMG_EX4DN As String = "\Back\Counsel\01\Ex\4일-gray.jpg"
Private Const IMG_EX5DY As String = "\Back\Counsel\01\Ex\5일-green.jpg"
Private Const IMG_EX5DN As String = "\Back\Counsel\01\Ex\5일-gray.jpg"
Private Const IMG_EX6DY As String = "\Back\Counsel\01\Ex\6일-green.jpg"
Private Const IMG_EX6DN As String = "\Back\Counsel\01\Ex\6일-gray.jpg"
'+---------------------------------------------------------------------------------+
'| 상담 > 처방/비만도상담 > 운동소모칼로리
'+---------------------------------------------------------------------------------+
Private Const IMG_EX0Y As String = "\Back\Counsel\01\Ex\200-green.jpg"
Private Const IMG_EX0N As String = "\Back\Counsel\01\Ex\200-gray.jpg"
Private Const IMG_EX1Y As String = "\Back\Counsel\01\Ex\300-green.jpg"
Private Const IMG_EX1N As String = "\Back\Counsel\01\Ex\300-gray.jpg"
Private Const IMG_EX2Y As String = "\Back\Counsel\01\Ex\400-green.jpg"
Private Const IMG_EX2N As String = "\Back\Counsel\01\Ex\400-gray.jpg"
Private Const IMG_EX3Y As String = "\Back\Counsel\01\Ex\500-green.jpg"
Private Const IMG_EX3N As String = "\Back\Counsel\01\Ex\500-gray.jpg"
'+---------------------------------------------------------------------------------+
'| 상담 > 처방/비만도상담 > 상하화살표(증권식)
'+---------------------------------------------------------------------------------+
Private Const IMG_UP As String = "\Back\Counsel\01\icon-red.jpg"
Private Const IMG_DOWN As String = "\Back\Counsel\01\icon-blue.jpg"

Private Const IMG_RES_ON As String = "\Back\Counsel\01\예약변경 on.jpg"
Private Const IMG_RES_OFF As String = "\Back\Counsel\01\예약변경 off.jpg"
Private Const IMG_SELEX_ON As String = "\Back\Counsel\01\운동종목선택 on.jpg"
Private Const IMG_SELEX_OFF As String = "\Back\Counsel\01\운동종목선택 off.jpg"
Private Const IMG_TREAT_ON As String = "\Back\Counsel\01\치료이력보기 on.jpg"
Private Const IMG_TREAT_OFF As String = "\Back\Counsel\01\치료이력보기 off.jpg"
'+---------------------------------------------------------------------------------+
'| 상담 > 처방/비만도상담 > 탭, 저장, 삭제, 출력
'+---------------------------------------------------------------------------------+
Private Const PATH01 As String = "\Back\Counsel\01\"
Private Const IMG_LTAB1_ON As String = "변화도표 on.jpg"
Private Const IMG_LTAB1_OFF As String = "변화도표 off.jpg"
Private Const IMG_LTAB2_ON As String = "변화도그래프 on.jpg"
Private Const IMG_LTAB2_OFF As String = "변화도그래프 off.jpg"
Private Const IMG_LTAB3_ON As String = "처방내역조회 on.jpg"
Private Const IMG_LTAB3_OFF As String = "처방내역조회 off.jpg"
Private Const IMG_RTAB1_ON As String = "치료종목 on.jpg"
Private Const IMG_RTAB1_OFF As String = "치료종목 off.jpg"
Private Const IMG_RTAB2_ON As String = "notice on.jpg"
Private Const IMG_RTAB2_OFF As String = "notice off.jpg"
Private Const IMG_RTAB3_ON As String = "memo on.jpg"
Private Const IMG_RTAB3_OFF As String = "memo off.jpg"
Private Const IMG_SAVE_ON As String = "save on.jpg"
Private Const IMG_SAVE_OFF As String = "save off.jpg"
Private Const IMG_DEL_ON As String = "delete on.jpg"
Private Const IMG_DEL_OFF As String = "delete off.jpg"
Private Const IMG_PRINT_ON As String = "비만도평가 on.jpg"
Private Const IMG_PRINT_OFF As String = "비만도평가 off.jpg"

Private Const IMG_NEW_ON As String = "new_on.jpg"
Private Const IMG_NEW_OFF As String = "new_off.jpg"
Private Const IMG_PRETREAT_ON As String = "pretreat_on.jpg"
Private Const IMG_PRETREAT_OFF As String = "pretreat_off.jpg"
Private Const IMG_MODIFYTREAT_ON As String = "modify_on.jpg"
Private Const IMG_MODIFYTREAT_OFF As String = "modify_off.jpg"


'+---------------------------------------------------------------------------------+
'| 평가 > 2Fon
'+---------------------------------------------------------------------------------+
'Private Const IMG_평가_ON As String = "\Img\기타\평가 on.jpg"
'Private Const IMG_평가_OFF As String = "\Img\기타\평가 off.jpg"

Private Const IMG_평가_ON As String = "\Back\Counsel\01\valuation_on.jpg"
Private Const IMG_평가_OFF As String = "\Back\Counsel\01\valuation_off.jpg"

'환자의 구조체타입 모듈변수
Private Type mCustomer
    strCustomName As String
    strJuminNum As String
    strSex As String
    intAge As Integer
    strChPgCode As String       '치료프로그램 코드
End Type

'처방칼로리의 구조체타입 모듈변수
Private Type mTreatCal
    intDietCal As Integer
    intExCal As Integer
    intExDay As Integer
    sngTreatCal As Single
    sngLossWeight As Single
    intUserCal As Single
End Type

'비만도관련 구조체타입 모듈변수
'체중
'비만도
'허리둘레
'WHR
'BMI
'체지방률
'RMR
'TEE
Private Type mObesity
    sngWeight As Single
    sngObesityRate As Single
    sngWaist As Single
    sngWHR As Single
    sngBMI As Single
    sngChFatRate As Single
    sngRMR As Single
    sngTEE As Single
End Type

Private typCustomer As mCustomer
Private typTreatCal As mTreatCal
Private typObesity As mObesity

'------------------------------------------------------------------------------
'|조회수정모드=0,        최초로드시 신규입력모드=1,
'|최초로드시수정모드=2,  무조건 신규입력모드=3
'|최초로드시수정모드(2)는 신체데이터 혹은 처방데이터가 있는 경우임.
'|최초로드시 신규입력모드(1)은 신체데이터와 처방데이터가 둘다 없는 경우임.
'|조회수정모드(0)은 처방리스트 혹은 변화도리스트에서 처방레코드를 클릭했을경우임.
'|무조건 신규입력모드(3)은 어떤상황에서도 새로입력을 하고자 하는 경우임.
'------------------------------------------------------------------------------
Public intMode As Integer
'------------------------------------------------------------------------------

Private glngBodyDataNum As Long       '신체계측 순번
Public glngTreatNum As Long           '진료 순번
Private glngCompDataNum As Long       '계산식 순번

Private gintBottomButton As Integer
Private gintLeftButton As Integer

'운동처방종목 관련
Public gintMain As Integer    '주종목 한개
Public gintSub1 As Integer    '부종목 네개
Public gintSub2 As Integer
Public gintSub3 As Integer
Public gintSub4 As Integer

Public intNowExCalory As Integer   '저장은 아직 안 됐지만 처방된 운동칼로리

'출력 관련
Dim crxApplication As New CRAXDRT.Application
Dim crxReport As CRAXDRT.Report
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxFormula As CRAXDRT.FormulaFieldDefinition
Dim strServer As String, strDBName As String, strUID As String, strPWD As String

'식사감량칼로리, 운동일수, 운동소비칼로리 클릭한 인덱스 기억
Dim idxDietCal As Integer, idxExDay As Integer, idxExCal As Integer

Dim mlngcTreatNum As Long   '현재 조회하고 있는 진료순번
Dim mlngoTreatNum As Long   '현재 조회하고 있는 신체계측정보를 저장한 진료순번

'2005-02-04 류진선 대용식 칼로리 저장
Dim msngChMealCalory As Single
Dim mlngChMealTime As Long
Dim msChMealCode As String

'2005-01-23 류진선 처방컨트롤을 입력가능/불가능으로 만듬.
Private Sub EnabledInput(Optional bEnable As Boolean = True)
Dim ctrl As Control, i As Integer
    If bEnable Then
        lblDispContent = "입력중"
        'cmdInputTreat.Caption = "처방입력"
    Else
        lblDispContent = "조회중"
        'cmdInputTreat.Caption = "처방수정"
    End If
    Select Case intMode
    Case 0
        lblDispContent = "조회중"
    Case 1
        lblDispContent = "입력중"
    Case 2
        lblDispContent = "수정중"
    Case 3
        lblDispContent = "입력중"
    End Select
    
    cmbProgram.Enabled = bEnable
    imgGoReserve.Enabled = bEnable
    imgPreTreat.Enabled = bEnable
    If bEnable Then
        If chkUserCalory.Value = vbChecked Then
            For Each ctrl In imgDietCal
                ctrl.Enabled = False
            Next
            If chkChMeal.Value = vbChecked Then
                If cboChMealTime.ItemData(cboChMealTime.ListIndex) = 10 Then
                    txtUserCalory.Enabled = False
                Else
                    txtUserCalory.Enabled = True
                End If
            Else
                txtUserCalory.Enabled = True
            End If
        Else
            For Each ctrl In imgDietCal
                ctrl.Enabled = True
            Next
            txtUserCalory.Enabled = False
        End If
        chkChMeal.Enabled = bEnable
        cboChMealTime.Enabled = bEnable
        cboChMealCode.Enabled = bEnable
    Else
        For Each ctrl In imgDietCal
            ctrl.Enabled = bEnable
        Next
        txtUserCalory.Enabled = bEnable
        chkChMeal.Enabled = bEnable
        cboChMealTime.Enabled = bEnable
        cboChMealCode.Enabled = bEnable
    End If
    For Each ctrl In imgExDay
        ctrl.Enabled = bEnable
    Next
    For Each ctrl In imgExCal
        ctrl.Enabled = bEnable
    Next
    chkUserCalory.Enabled = bEnable
    'imgSelectEx.Enabled = bEnable
    With sprTreat
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2: .Lock = (Not bEnable)
            .Col = 3: .Lock = (Not bEnable)
        Next
    End With
    With sprPrint
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2: .Lock = (Not bEnable)
        Next
    End With
    txtMemo.Enabled = bEnable
    txtNotice.Enabled = bEnable
    gbIsEnable = bEnable
End Sub

Private Sub cboChMealCode_Click()
    If cboChMealTime.ItemData(cboChMealTime.ListIndex) = 10 Then
        txtUserCalory.Text = cboChMealCode.ItemData(cboChMealCode.ListIndex)
    End If
End Sub

Private Sub cboChMealTime_Click()
Dim rValue As Variant
Dim clsSelect As clsSelect_2
Dim i As Integer
    If cboChMealTime.ItemData(cboChMealTime.ListIndex) = "10" Then  '사용자식단
        txtUserCalory.Enabled = False
        Set clsSelect = New clsSelect_2
        rValue = clsSelect.Query("SELECT UserMenuName, Calory FROM Usermenu")
        cboChMealCode.Clear
        If Not IsNull(rValue) Then
            For i = 0 To UBound(rValue, 2)
                cboChMealCode.AddItem Trim(rValue(0, i)) & "-" & Trim(rValue(1, i)) & " Kcal"
                cboChMealCode.ItemData(i) = Trim(rValue(1, i))
            Next i
            cboChMealCode.ListIndex = 0
        End If
        Set clsSelect = Nothing
    ElseIf cboChMealTime.ItemData(cboChMealTime.ListIndex) = "0" Then
        txtUserCalory.Enabled = True
        cboChMealCode.Clear
    Else
        txtUserCalory.Enabled = True
        Set clsSelect = New clsSelect_2
        rValue = clsSelect.Query("SELECT ChMealName, ChMealCalory FROM ChangeMeal")
        cboChMealCode.Clear
        If Not IsNull(rValue) Then
            For i = 0 To UBound(rValue, 2)
                cboChMealCode.AddItem Trim(rValue(0, i)) & "-" & Trim(rValue(1, i)) & " Kcal"
                cboChMealCode.ItemData(i) = Trim(rValue(1, i))
            Next i
            cboChMealCode.ListIndex = 0
        End If
        Set clsSelect = Nothing
    End If
End Sub

Private Sub chkChMeal_Click()
    AdminYn 11
    If AccessYn = False Then
        MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
        Exit Sub
    End If
    If chkChMeal.Value = vbChecked Then
        cboChMealTime.Visible = True
        cboChMealCode.Visible = True
        cboChMealTime_Click
    Else
        cboChMealTime.Visible = False
        cboChMealCode.Visible = False
        If chkUserCalory.Value = vbChecked Then
            txtUserCalory.Enabled = True
        End If
    End If
End Sub

Private Sub chkUserCalory_Click()
    Dim i As Integer
    Dim sngMinus As Single
    
     AdminYn 11
    If AccessYn = False Then
        MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
        Exit Sub
    End If
    
    With typTreatCal
    If chkUserCalory.Value = 0 Then
        For i = 0 To 5
            imgDietCal(i).Enabled = True
        Next i
        lblTreatCal.Caption = Format(typObesity.sngTEE, "#,###") & " kcal"
        txtUserCalory.Text = ""
        If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
        End If
        .intUserCal = 0
        txtUserCalory.BackColor = FRM_GRAY
        txtUserCalory.Enabled = False
        chkChMeal.Value = vbUnchecked
        chkChMeal.Enabled = False
        
    Else
        For i = 0 To 5
            Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
            imgDietCal(i).Enabled = False
        Next i
        .intDietCal = 0
        If txtUserCalory.Text <> "" And IsNumeric(txtUserCalory.Text) = True Then
            .intUserCal = txtUserCalory.Text
        Else
            .intUserCal = typObesity.sngTEE
        End If
        .sngTreatCal = .intUserCal
        lblTreatCal.Caption = Format(.sngTreatCal, "#,##0") & " kcal"
    '+=========================================
    '+ 감량체중 구하는 식
    '+-----------------------------------------
        sngMinus = typObesity.sngTEE - .intUserCal
        If sngMinus >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((sngMinus * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
        End If
        txtUserCalory.BackColor = vbWhite
        txtUserCalory.Enabled = True
        chkChMeal.Enabled = True
    End If
    End With
End Sub

'Private Sub cmdInputTreat_Click()
'    '만약 frmBottom.dtpUserDay.Value에 해당하는 처방이 없다면
'    '새로 입력한다.
'    '만약 frmBottom.dtpUserDay.Value에 해당하는 처방이 있다면
'    '기존 데이터를 수정한다.
'    '기존 데이터를 수정할때 해당 일자에 여러번의 처방이 있을수 있으므로
'    '해당 날짜의 처방번호를 가지고 데이터를 수정한다.
'
'    If cmdInputTreat.Caption = "처방입력" Then
'        lblDispDate.Caption = Format(gdatUserDay, "YYYY-MM-DD")
'        intMode = 1
'        Call EnabledInput(True)
'    Else
'        intMode = 2
'        Call EnabledInput(True)
'    End If
'    cmdInputTreat.Visible = False
'    Dim qrySelect As String
''    qrySelect = "SELECT TOP 1 * FROM Treat a LEFT JOIN "
''
''    Call EnabledInput(Not cmbProgram.Enabled)
'End Sub

'***************************************************************************************
' 프로시져명    : Form_Load
' 작성일시  : 2005-01-28 10:13
' 작 성 자  : 류진선
' 설    명  : 폼과 폼위의 각종컨트롤을 초기화하고
'             환자정보를 가져오고 칼로리처방리스트를 표시하며
'             현재날짜의 칼로리 처방을 불러온다.
'             glngTreatNum에 현재 처방의 번호를 저장하고
'             로드시 입력/수정 모드를 결정해서 intMode에 저장한다.
'***************************************************************************************
Public Sub Form_Load()
    Dim i As Integer
'1) 신체계측,캘리퍼,생화학수치등을 입력할 수 있도록 스프레드 셋팅
'2) 재진환자의 경우 예약된 사항이 있으면 처치와 출력물에 미리 보여주고, 비만도에 대해서도 보여준다
'3) 이전에 저장된 내역들을 표,그래프,처방을 볼 수 있게 한다.
On Error GoTo ShowErr
    '폼초기화
    Set Me.Picture = LoadPicture(App.Path & "\Back\Counsel\01\상담_처방 back.jpg")
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.Width = FRM_WIDTH
    Me.Height = FRM_HEIGHT
    Me.BackColor = vbWhite
    Set imgGoReserve.Picture = LoadPicture(App.Path & IMG_RES_ON)
    Set imgSelectEx.Picture = LoadPicture(App.Path & IMG_SELEX_ON)
    
    '변수 초기화
    gintBottomButton = 0
    gintLeftButton = 0
    glngBodyDataNum = 0
    glngTreatNum = 0
    glngCompDataNum = 0
    
    gintMain = 0
    gintSub1 = 0: gintSub2 = 0: gintSub3 = 0: gintSub4 = 0
    
    '2005-01-28 류진선
    mlngcTreatNum = 0: mlngoTreatNum = 0
    
    '2005-01-27 류진선 수정
    '모든 컨트롤 초기화
    Call InitialControl2
    '--------------------------------------------------------
    '초기화 끝
    
    '--------------------------------------------------------
    '데이터 로드 시작
    
    '1. 환자정보 로드
    If LoadCustomerInfo = False Then
'        Exit Sub
    End If

    'ObesityRecord(비만정보) 로드
    '=> 왼쪽 상단 정보 표시 및 typObesity 변수 로드
    Call ShowObesityRecord2
    
    'TreatRecord(치료정보) 로드
    Call imgViewHistory_Click(2)
    '오늘 날짜의 신체계측 정보가 있는지 확인
    If ShowTodayBodyData Then
        'TodayTreat(현재날짜의 치료정보) 로드
        If ShowTodayTreat2 > 0 Then
            '최근 칼로리처방번호를 cTreatNum에 저장
            If mlngoTreatNum = mlngcTreatNum Then
               'glngTreatNum에 oTreatNum저장
               glngTreatNum = mlngoTreatNum
            Else
               'glngTreatNum에 cTreatNum저장
               glngTreatNum = mlngcTreatNum
            End If
        End If
        intMode = 2
    Else
        If ShowTodayTreat2 > 0 Then
            'glngTreatNum에 최근칼로리처방번호 저장
            '=> ShowTodayTreat2에서 glngTreatNum에 저장함.
            
            '로드시 수정모드
            glngTreatNum = mlngcTreatNum
            intMode = 2
        Else
            'glngTreatNum에 새로운 TreatNum을 받아서 저장
            '새로입력할때는 저장시에 새 번호를 따온다.
            glngTreatNum = 0
            '로드시 새입력 모드
            intMode = 1
        End If
    End If
    Call EnabledInput
    lblDispDate.Caption = Format(gdatUserDay, "YYYY-MM-DD")

    '가장 최근 처방한 운동값을 가져오기
'    Call GetLatestExItem
    '--------------------------------------------------------
    '데이터 로드 끝
    
    Exit Sub
ShowErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "Form_Load", "frmCounsel_1", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

'오늘날짜의 처방이 있으면 보여주기
Private Sub ShowTodayTreat()
    Dim qrySelect As String, rValue As Variant
    
    qrySelect = "SELECT TreatNum, DietCalory, TreatCalory, UserCalory FROM Treat WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND TreatDay='" & Format(gdatUserDay, "YYYYMMDD") & "'"
    qrySelect = qrySelect & " ORDER BY TreatNum DESC "
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        glngTreatNum = CLng(rValue(0, 0))
        intMode = 2
        If (IsNull(rValue(1, 0)) Or IsNull(rValue(2, 0))) And (IsNull(rValue(2, 0)) Or IsNull(rValue(3, 0))) Then
            imgPreTreat.Visible = True
        Else
            imgPreTreat.Visible = False
        End If
        
        '2004-12-09 류진선 이놈들 왜 이리 왔다리 갔다리 하면서 데이터를 보여주는거지? ㅡㅡ;
        '==========
        '오른쪽 상단(?) 치료내용 보여주기
        Call ShowTreat(glngTreatNum)
        '오른쪽하단 치료종목에 저장된 데이터 보여주기
        Call ShowTreatData(glngTreatNum)
        '오른쪽하단 출력물에 저장된 데이터 보여주기
        Call ShowTreatPrint(glngTreatNum)
        '==========
        
        Call EnabledInput(False)
        lblDispDate.Caption = Format(gdatUserDay, "YYYY-MM-DD")
        lblDispContent.Caption = "조회중"
        imgModifyTreat.Visible = True
    Else
        Call imgNew_MouseUp(0, 0, 0, 0)
    End If
    
End Sub

Private Sub GetLatestExItem()
'가장 최근의 처방된 운동종목(주종목1, 부종목4) 가져오기
    Dim qrySelect As String, rValue As Variant
    
    qrySelect = "SELECT TOP 1 TreatNum, main, sub1, sub2, sub3, sub4 "
    qrySelect = qrySelect & "FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND NOT main IS NULL"
    qrySelect = qrySelect & " ORDER BY TreatDay DESC;"
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        gintMain = Is_Null(rValue(1, 0), 0)
        gintSub1 = Is_Null(rValue(2, 0), 0)
        gintSub2 = Is_Null(rValue(3, 0), 0)
        gintSub3 = Is_Null(rValue(4, 0), 0)
        gintSub4 = Is_Null(rValue(5, 0), 0)
    Else
        gintMain = 0
        gintSub1 = 0
        gintSub2 = 0
        gintSub3 = 0
        gintSub4 = 0
    End If
End Sub

Private Sub cmbProgram_Change()
    Dim qrySelect As String, rValue As Variant

    Set clsSelect = New clsSelect
    AdminYn 10
    If AccessYn = False Then
        Exit Sub
    End If

    qrySelect = "SELECT ChPgCode FROM ChPgUpdate WHERE ChPgName='" & Trim(cmbProgram.Text) & "';"
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        typCustomer.strChPgCode = Trim(rValue(0, 0))
    Else
        typCustomer.strChPgCode = 0
    End If
    Set clsSelect = Nothing
    Debug.Print "cmdProgram_Change()"
    '해당 치료프로그램대로 예약한다.
End Sub

Private Sub cmbProgram_Click()
    AdminYn 10
    If AccessYn = False Then
'        MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
        Exit Sub
    End If
    Call cmbProgram_Change
    Debug.Print "cmdProgram_Click()"

End Sub

'변화도(표),변화도(그래프),처방내역조회 탭의 서브버튼들 클릭이벤트
Private Sub cmdSub_Click(Index As Integer)
    gintLeftButton = Index

    Select Case gintBottomButton
        Case 0  '변화도(표)
            Call Bottom0(gintLeftButton)
        Case 1  '변화도(그래프)
            Call Bottom1(gintLeftButton)
        Case 2  '처방내역조회
            Call Bottom2(gintLeftButton)
    End Select
End Sub

'변화도(표) 탭의 서브버튼들의 클릭이벤트
Private Sub Bottom0(intIndex As Integer)
'0~2
    Dim i As Integer, j As Integer
    Dim qrySelect As String, rValue As Variant
    
On Error GoTo Err
    grdTable.Visible = True
    Chart.Visible = False
    '체성분, 신체둘레, 피부두께, 생화학검사버튼만 보이고 나머지 숨김
    '=============
    For i = 0 To 3
        cmdSub(i).Visible = True
    Next i
    For i = 4 To 9
        cmdSub(i).Visible = False
    Next i
    
    cmdSub(0).Caption = "체   성   분"
    cmdSub(1).Caption = "신 체 둘 레"
    cmdSub(2).Caption = "피 부 두 께"
    cmdSub(3).Caption = "생화학 검사"
    '=============
    
    Select Case intIndex
        Case 0   '신장,체중,VO2,RMR+체성분
            Set clsSelect = New clsSelect
            qrySelect = "SELECT InputName, InputField "
            qrySelect = qrySelect & "FROM C_InputData WHERE InputKind='B' "
            qrySelect = qrySelect & "ORDER BY InputOrder ASC;"
            
            rValue = clsSelect.Query(qrySelect)
            Set clsSelect = Nothing
            
            If Not IsNull(rValue) Then
                With grdTable
                    .Clear
                    .Rows = 3
                    .RowHeight(-1) = 300
                    .RowHeight(1) = 0
                    .Cols = UBound(rValue, 2) + 5 + 4 '+1 진료순번
                    .SelectionMode = flexSelectionByRow
                    
                    .TextMatrix(0, 0) = ""          '순번
                    .ColWidth(0) = 300
                    .TextMatrix(0, 1) = "진료일"
                    .ColWidth(1) = 1000
                    For i = 2 To .Cols - 1
                        .ColWidth(i) = 1000
                    Next i
                    'InputOrder의 순서대로 보여준다.
                    '신체상태, 활동정도는
                    .TextMatrix(0, 2) = "신장"
                    .TextMatrix(1, 2) = "Height"
                    .TextMatrix(0, 3) = "체중"
                    .TextMatrix(1, 3) = "Weight"
                    .TextMatrix(0, 4) = "VO2"
                    .TextMatrix(1, 4) = "VO2"
                    .TextMatrix(0, 5) = "측정RMR"
                    .TextMatrix(1, 5) = "inRMR"
                    .TextMatrix(0, 6) = "추정RMR"
                    .TextMatrix(1, 6) = "RMR"
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 7) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 7) = Trim(rValue(1, i))
                    Next i
                    .TextMatrix(1, i + 7) = "Treat.TreatNum"
                    .ColAlignment(-1) = flexAlignCenterCenter
                    .ColWidth(i + 7) = 0
                End With
                Erase rValue
                '초진이면 그냥 넘어가지만 재진인 경우
                Call ShowTable_Measure
            End If
        Case 1    '신체둘레
            Set clsSelect = New clsSelect
            qrySelect = "SELECT InputName, InputField "
            qrySelect = qrySelect & "FROM C_InputData WHERE InputKind='C' "
            qrySelect = qrySelect & "ORDER BY InputOrder ASC;"
            
            rValue = clsSelect.Query(qrySelect)
            Set clsSelect = Nothing
            
            If Not IsNull(rValue) Then
                With grdTable
                    .Clear
                    .Rows = 3
                    .RowHeight(-1) = 300
                    .RowHeight(1) = 0
                    .Cols = UBound(rValue, 2) + 4                     '+1 진료순번
                    .SelectionMode = flexSelectionByRow
                    .TextMatrix(0, 0) = ""          '순번
                    .ColWidth(0) = 300
                    .TextMatrix(0, 1) = "진료일"
                    .ColWidth(1) = 1000
                    For i = 2 To .Cols - 1
                        .ColWidth(i) = 1290
                    Next i
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 2) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 2) = Trim(rValue(1, i))
                    Next i
                    .TextMatrix(1, i + 2) = "Treat.TreatNum"
                    .ColAlignment(-1) = flexAlignCenterCenter
                    .ColWidth(i + 2) = 0
                End With
                Erase rValue
                Call ShowTable_Circumference
            End If
        Case 2     '피부두께
            Set clsSelect = New clsSelect
            qrySelect = "SELECT InputName, InputField "
            qrySelect = qrySelect & "FROM C_InputData WHERE InputKind='S' "
            qrySelect = qrySelect & "ORDER BY InputOrder ASC;"
            
            rValue = clsSelect.Query(qrySelect)
            Set clsSelect = Nothing
            
            If Not IsNull(rValue) Then
                With grdTable
                    .Clear
                    .Rows = 3
                    .RowHeight(-1) = 300
                    .RowHeight(1) = 0
                    .Cols = UBound(rValue, 2) + 4                     '+1 진료순번
                    .SelectionMode = flexSelectionByRow
                    .TextMatrix(0, 0) = ""          '순번
                    .ColWidth(0) = 300
                    .TextMatrix(0, 1) = "진료일"
                    .ColWidth(1) = 1000
                    For i = 2 To .Cols - 1
                        .ColWidth(i) = 1290
                    Next i
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 2) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 2) = Trim(rValue(1, i))
                    Next i
                    .TextMatrix(1, i + 2) = "Treat.TreatNum"
                    .ColAlignment(-1) = flexAlignCenterCenter
                    .ColWidth(i + 2) = 0
                End With
                Erase rValue
'                Call ShowTable_Caliper
                Call ShowTable_Circumference
            End If
        Case 3     '생화학검사
            Set clsSelect = New clsSelect
            qrySelect = "SELECT InputName, InputField "
            qrySelect = qrySelect & "FROM C_InputData WHERE InputKind='L' "
            qrySelect = qrySelect & "ORDER BY InputOrder ASC;"
            
            rValue = clsSelect.Query(qrySelect)
            Set clsSelect = Nothing
            
            If Not IsNull(rValue) Then
                With grdTable
                    .Clear
                    .Rows = 3
                    .RowHeight(-1) = 300
                    .RowHeight(1) = 0
                    .Cols = UBound(rValue, 2) + 4                     '+1 진료순번
                    .SelectionMode = flexSelectionByRow
                    .TextMatrix(0, 0) = ""          '순번
                    .ColWidth(0) = 300
                    .TextMatrix(0, 1) = "진료일"
                    .ColWidth(1) = 1000
                    For i = 2 To .Cols - 1
                        .ColWidth(i) = 1290
                    Next i
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 2) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 2) = Trim(rValue(1, i))
                    Next i
                    .TextMatrix(1, i + 2) = "Treat.TreatNum"
                    .ColAlignment(-1) = flexAlignCenterCenter
                    .ColWidth(i + 2) = 0
                End With
                Erase rValue
                Call ShowTable_BioChem
            End If
    End Select
    Call SetGridRow
    Exit Sub
Err:
    '2004-12-23 류진선 로그기록
    'WriteLog "Bottom0", "frmCounsel_1", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

'왼쪽 하단의 변화도(표)의 스프래드에 체성분등을 입력한다.
Private Sub ShowTable_Measure()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer

    Set clsSelect = New clsSelect

    With grdTable
    qrySelect = "SELECT TreatDay"
    For i = 1 To .Cols - 2
        qrySelect = qrySelect & "," & Trim(.TextMatrix(1, i + 1))
    Next i
    qrySelect = qrySelect & " FROM BodyData INNER JOIN Treat ON "
    qrySelect = qrySelect & "BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY TreatDay DESC, Treat.TreatNum DESC;"

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        .Rows = .Rows + UBound(rValue, 2)
        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = i + 1
            For j = 0 To .Cols - 2
                If Not IsNull(rValue(j, i)) Then
                    .TextMatrix(i + 2, j + 1) = rValue(j, i)
                Else
                    .TextMatrix(i + 2, j + 1) = "-"
                End If
            Next j
' 2004-12-09 류진선 처방날짜와 아래 작업날짜가 같으면 수정모드로 들어가는 것으로 보이는데 이루틴을 절대로 탈일이 없음.
' 혹시 타야 한다면 그때 다시 한번 ㅡㅡ; 아래 조건문으로 대치

'=========================================================
' 2005-01-26 류진선 여기서 왜 수정모드로 들어가야 하지????
' 수정모드는 수정버튼을 클릭했을경우만...
'오늘 날짜 데이터를 보여주기 위해서하는 것인가?
'=========================================================
            
            If Trim(rValue(0, i)) = Format(gdatUserDay, "yyyy-MM-dd") Then
                
                
                If glngTreatNum = Trim(.TextMatrix(i + 2, .Cols - 1)) Then
                
                'If Trim(rValue(12, i)) = Trim(.TextMatrix(i + 2, .ColS - 1)) Then
    '            If rValue(0, i) = gdatUserDay Then
                    '.Row = i + 2: .Col = 0: .ColSel = .ColS - 1
                    'glngTreatNum = .TextMatrix(i + 2, .ColS - 1)
                    '조회,수정 모드
                   ' intMode = 2
    
                    Call ShowTreat(Trim(.TextMatrix(i + 2, .Cols - 1)))
                    Call ShowTreatData(Trim(.TextMatrix(i + 2, .Cols - 1)))
                    Call ShowTreatPrint(Trim(.TextMatrix(i + 2, .Cols - 1)))
                Else
                    
                End If
            End If
        Next i
    End If
    End With

    Set clsSelect = Nothing
End Sub

Private Sub ShowTable_BioChem()
    Dim qrySelect As String, rValue As Variant
    Dim qrySelect1 As String
    Dim i As Integer, j As Integer, k As Integer
    Dim intCount As Integer
    
    Set clsSelect = New clsSelect
    
'    If grdTable.TextMatrix(1, 1) = "" Then
'        Exit Sub
'    End If






    
    intCount = grdTable.Cols - 3
    
    qrySelect = "SELECT Treat.TreatNum, TreatDay, "
    For i = 1 To intCount - 1
        qrySelect = qrySelect & "a" & i & ","
    Next i
    qrySelect = qrySelect & "a" & i & " FROM Treat INNER JOIN ("
    qrySelect = qrySelect & "SELECT TreatNum, "
    For i = 1 To intCount - 1
        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
    Next i
    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
        
    '서브쿼리
    qrySelect1 = "SELECT Treat.TreatNum, "
    For i = 1 To intCount
        qrySelect1 = qrySelect1 & " CASE BioChemCode WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN BioChemSu ELSE 0 END AS t" & i & ","
    Next i
    qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
    qrySelect = qrySelect & " FROM BodyData INNER JOIN Treat "
    qrySelect = qrySelect & " ON BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & " LEFT JOIN BioChemData "
    qrySelect = qrySelect & " ON Treat.TreatNum=BioChemData.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " GROUP BY Treat.TreatNum, BioChemCode, BioChemSu "
    qrySelect = qrySelect & ") a GROUP BY TreatNum) b ON Treat.TreatNum=b.TreatNum"
    qrySelect = qrySelect & " ORDER BY TreatDay DESC, Treat.TreatNum DESC "
    
    
    
    
'    qrySelect = qrySelect & "SELECT TreatNum, "
'    For i = 1 To intCount - 1
'        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
'    Next i
'    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
'    For i = 1 To intCount
'        qrySelect1 = "SELECT Treat.TreatNum, "
'        If i > 0 Then
'            For j = 0 To i - 1
'                qrySelect1 = qrySelect1 & "0 AS t" & j & ","
'            Next j
'        End If
'        If i <= intCount Then
'            qrySelect1 = qrySelect1 & "CASE BioChemCode "
'            qrySelect1 = qrySelect1 & "WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN BioChemSu ELSE 0 END AS t" & i & ","
'            For k = i + 1 To intCount
'                qrySelect1 = qrySelect1 & "0 AS t" & k & ","
'            Next k
'        End If
'        qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
'        qrySelect = qrySelect & " FROM BodyData INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
'        qrySelect = qrySelect & "LEFT JOIN BioChemData ON Treat.TreatNum=BioChemData.TreatNum "
'        qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
'        qrySelect = qrySelect & " GROUP BY Treat.TreatNum, BioChemCode, BioChemSu "
'        If i < intCount Then
'            qrySelect = qrySelect & "UNION ALL "
'        End If
'    Next i
'    qrySelect = qrySelect & ") a GROUP BY TreatNum) b ON Treat.TreatNum=b.TreatNum"
'    qrySelect = qrySelect & " ORDER BY TreatDay"
    
    rValue = clsSelect.Query(qrySelect)
    
    If Not IsNull(rValue) Then
        With grdTable
            .Rows = UBound(rValue, 2) + 3
            For i = 0 To UBound(rValue, 2)
                .TextMatrix(i + 2, 0) = i + 1       '일련번호
                .TextMatrix(i + 2, 1) = rValue(1, i)
                For j = 2 To intCount + 1
                    If rValue(j, i) = 0 Then
                        .TextMatrix(i + 2, j) = "-"
                    Else
                        .TextMatrix(i + 2, j) = rValue(j, i)
                    End If
                Next j
                .TextMatrix(i + 2, intCount + 2) = rValue(0, i)
            Next i
            .ColAlignment(-1) = flexAlignCenterCenter
        End With
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub ShowTable_TreatCode()
    Dim qrySelect As String, rValue As Variant
    Dim qrySelect1 As String
    Dim i As Integer, j As Integer, k As Integer
    Dim intCount As Integer

    Set clsSelect = New clsSelect
    
    '만약 현재 설정된 처방이 하나도 없는 경우(TreatCode가 없는 경우)에는 그냥 넘어가기
    If grdTable.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    
'   일단 현재 저장되어 있는 TreatCode를 다 불러올린다.
    intCount = grdTable.Cols - 3

    qrySelect = "SELECT TreatNum, TreatDay, "
    For i = 0 To intCount - 1
        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
    Next i
    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
        
        
    '서브쿼리
    qrySelect1 = "SELECT Treat.TreatNum, TreatDay, "
    For i = 0 To intCount
        qrySelect1 = qrySelect1 & " CASE TreatCode WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN 1 ELSE 0 END AS t" & i & ","
    Next i
    'qrySelect = qrySelect & "SELECT Treat.TreatNum, TreatDay, "
    qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
    qrySelect = qrySelect & " FROM Treat LEFT JOIN TreatData "
    qrySelect = qrySelect & "ON Treat.TreatNum=TreatData.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & "AND TreatCalory IS NOT NULL "
    qrySelect = qrySelect & ") a GROUP BY TreatNum, TreatDay"
    qrySelect = qrySelect & " ORDER BY TreatDay DESC "
    
    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
    With grdTable
        .Rows = UBound(rValue, 2) + 3
        .SelectionMode = flexSelectionByRow

        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = rValue(1, i)
            For j = 1 To intCount + 1
                If CInt(rValue(j + 1, i)) = 1 Then
                    .TextMatrix(i + 2, j) = "O"
                Else
                    .TextMatrix(i + 2, j) = "X"
                End If
            Next j
            .TextMatrix(i + 2, intCount + 2) = rValue(0, i)
        Next i
        .ColAlignment(-1) = flexAlignCenterCenter
    End With
    End If

    Set clsSelect = Nothing
End Sub

Private Sub ShowTable_Print()
    Dim qrySelect As String, rValue As Variant
    Dim qrySelect1 As String
    Dim i As Integer, j As Integer, k As Integer
    Dim intCount As Integer

    Set clsSelect = New clsSelect
    '만약 현재 설정된 처방이 하나도 없는 경우에는 그냥 넘어가기
    If grdTable.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    
    intCount = grdTable.Cols - 3

    qrySelect = "SELECT TreatNum, TreatDay, "
    For i = 0 To intCount - 1
        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
    Next i
    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
        
    '서브쿼리
    qrySelect1 = "SELECT Treat.TreatNum, TreatDay, "
    For i = 0 To intCount
        qrySelect1 = qrySelect1 & " CASE PrintoutNum WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN 1 ELSE 0 END AS t" & i & ","
    Next i
    qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
    qrySelect = qrySelect & " FROM Treat LEFT JOIN TreatPrint "
    qrySelect = qrySelect & "ON Treat.TreatNum=TreatPrint.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & "AND TreatCalory IS NOT NULL "
    qrySelect = qrySelect & ") a GROUP BY TreatNum, TreatDay"
    qrySelect = qrySelect & " ORDER BY TreatDay DESC "


    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
    With grdTable
        .Rows = UBound(rValue, 2) + 3
        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = rValue(1, i)
            For j = 1 To intCount + 1
                If CInt(rValue(j + 1, i)) = 1 Then
                    .TextMatrix(i + 2, j) = "O"
                Else
                    .TextMatrix(i + 2, j) = "X"
                End If
            Next j
            .TextMatrix(i + 2, intCount + 2) = rValue(0, i)
        Next i
        .ColAlignment(-1) = flexAlignCenterCenter
    End With
    End If

    Set clsSelect = Nothing
End Sub

'변화도(그래프) 탭의 서브버튼들의 클릭이벤트
Private Sub Bottom1(intIndex As Integer)
    Dim i As Integer
    Dim qrySelect As String, rValue As Variant
    Dim sngMin As Single, sngMax As Single
    Dim strTitle As String

On Error GoTo Err
    grdTable.Visible = False
    For i = 0 To 9
        cmdSub(i).Visible = True
    Next i

    ' *** 변화도 그래프 항목
    ' 체중, 허리둘레, WHR, RMR, 체지방률, 근육량, 칼로리(식사일기)
    cmdSub(0).Caption = "체      중"
    cmdSub(1).Caption = "허리둘레"
    cmdSub(2).Caption = "상완둘레"
    cmdSub(3).Caption = "W  H  R"
    cmdSub(4).Caption = "V  O  2"
    cmdSub(5).Caption = "R  M  R"
    cmdSub(6).Caption = "체지방률"
    cmdSub(7).Caption = "근 육 량"
    cmdSub(8).Caption = "칼로리(식사)"
    cmdSub(9).Caption = "칼로리(운동)"

    Call InitialChart
    ' *** 변화도 그래프 항목
    ' 체중, 허리둘레, WHR, 체지방률, 근육량, 휴식대사량, 칼로리(식사일기), 칼로리(운동일기)
    ' --> 체중, 허리둘레, WHR, RMR, 체지방률, 근육량, 칼로리(식사일기)
    Dim sngStep As Single
    Select Case intIndex
        Case 0   '체중
            qrySelect = "SELECT TreatDay, Weight FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND Weight > 0 ORDER BY TreatDay ASC;"
            sngMin = MinValue("Weight") - 1
            sngMax = MaxValue("Weight") + 1
            sngStep = 2
            strTitle = "체중 변화"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 1   '허리둘레
            qrySelect = "SELECT TreatDay, Waist FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND Waist IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("Waist") - 5
            sngMax = MaxValue("Waist") + 5
            sngStep = 5
            strTitle = "허리둘레 변화"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 2   '상완둘레
            qrySelect = "SELECT TreatDay, UpperArm FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND UpperArm IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("UpperArm") - 3
            sngMax = MaxValue("UpperArm") + 3
            sngStep = 3
            strTitle = "상완둘레 변화"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 3   'WHR
            qrySelect = "SELECT TreatDay, WHR FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND WHR IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("WHR") - 0.1
            sngMax = MaxValue("WHR") + 0.1
            sngStep = 0.1
            strTitle = "WHR 변화"
            Chart.Axis(AXIS_Y).Decimals = 2
        Case 4    'VO2
            qrySelect = "SELECT TreatDay, VO2 FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND VO2 IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("VO2") - 10
            sngMax = MaxValue("VO2") + 10
            sngStep = 10
            strTitle = "VO2 변화"
            Chart.Axis(AXIS_Y).Decimals = 0
        Case 5    'RMR
            qrySelect = "SELECT TreatDay, RMR FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND RMR IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("RMR") - 50
            sngMax = MaxValue("RMR") + 50
            sngStep = 100
            strTitle = "RMR 변화"
            Chart.Axis(AXIS_Y).Decimals = 0
        Case 6   '체지방률
            qrySelect = "SELECT TreatDay, ChFatRate FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND ChFatRate IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("ChFatRate") - 2
            sngMax = MaxValue("ChFatRate") + 2
            sngStep = 2
            strTitle = "체지방률 변화"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 7   '근육량
            qrySelect = "SELECT TreatDay, Muscle FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND Muscle IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("Muscle") - 2
            sngMax = MaxValue("Muscle") + 2
            sngStep = 2
            strTitle = "근육량 변화"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 8   '칼로리(식사일기)
            '식사일기 입력한 날과 당시 총칼로리를 보여줌
            '최대값, 최소값 따로 구한다.
            qrySelect = "SELECT MAX(a), MIN(a) FROM ( "
            qrySelect = qrySelect & "SELECT SUM(MealCalory) AS a FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " GROUP BY MealDate) total;"
            Set clsSelect = New clsSelect
            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue) Then
                If IsNull(rValue(0, 0)) And IsNull(rValue(1, 0)) Then
                    MsgBox "표시할 입력데이터가 없습니다.", vbExclamation, "변화도 그래프"
                    Chart.Visible = False
                    Exit Sub
                End If
                sngMax = CInt(rValue(0, 0)) + 10
                sngMin = CInt(rValue(1, 0)) - 10
            Else
                sngMin = 500
                sngMax = 3000
            End If
            Set clsSelect = Nothing
            
            qrySelect = "SELECT MealDate, SUM(MealCalory) FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " GROUP BY MealDate ORDER BY MealDate ASC;"
            sngStep = 500
            strTitle = "식사량(칼로리) 변화"
             Chart.Axis(AXIS_Y).Decimals = 0
       Case 9   '칼로리(운동일기)
            qrySelect = "SELECT MAX(a), MIN(a) FROM ( "
            qrySelect = qrySelect & "SELECT SUM(BurnCalories) AS a FROM Sportsdiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " GROUP BY PlayDay) total;"
            Set clsSelect = New clsSelect
            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue) Then
                If IsNull(rValue(0, 0)) And IsNull(rValue(1, 0)) Then
                    MsgBox "표시할 입력데이터가 없습니다.", vbExclamation, "변화도 그래프"
                    Chart.Visible = False
                    Exit Sub
                End If
                sngMax = CInt(rValue(0, 0)) + 10
                sngMin = CInt(rValue(1, 0)) - 10
            Else
                sngMin = 100
                sngMax = 1500
            End If
            Set clsSelect = Nothing
            
            qrySelect = "SELECT PlayDay, SUM(BurnCalories) FROM Sportsdiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " GROUP BY PlayDay ORDER BY PlayDay ASC;"
            sngStep = 200
            strTitle = "운동량(칼로리) 변화"
            Chart.Axis(AXIS_Y).Decimals = 0
    End Select
    Set clsSelect = New clsSelect

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        Chart.Visible = True
        Chart.Title(CHART_TOPTIT) = strTitle
        Chart.OpenDataEx COD_VALUES, 0, COD_UNKNOWN
        Chart.Axis(AXIS_Y).Min = sngMin
        Chart.Axis(AXIS_Y).Max = sngMax
        For i = 0 To UBound(rValue, 2)
            Chart.ValueEx(0, i) = rValue(1, i)
            Chart.Axis(AXIS_X).Label(i) = Format(Is_Null(rValue(0, i), ""), "M/D")
        Next i
        Chart.CloseData COD_VALUES
    Else
        MsgBox "표시할 입력데이터가 없습니다.", vbExclamation, "변화도 그래프"
        Chart.Visible = False
    End If

    Set clsSelect = Nothing
    Exit Sub
Err:
    '2004-12-23 류진선 로그기록
    'WriteLog "Bottom1", "frmCounsel_1", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

Private Function MinValue(strField As String) As Single
    Dim qrySelect As String, rMin As Variant

    Set clsSelect = New clsSelect

    If strField = "RMR" Then
        qrySelect = "SELECT MIN(rmr) FROM ( "
        If typCustomer.intAge >= ADULT_AGE Then
            qrySelect = qrySelect & "SELECT CASE AdBasicDsa "
        Else
            qrySelect = qrySelect & "SELECT CASE BaBasicDsa "
        End If
        qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END AS rmr "
        qrySelect = qrySelect & "FROM BodyData INNER JOIN CompData "
        qrySelect = qrySelect & "ON BodyData.CompDataNum=CompData.CompDataNum "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & ") a"
    Else
        qrySelect = "SELECT MIN(" & strField & ") FROM BodyData WHERE CustomerNum=" & glngCustomerNum
    End If
    
    rMin = clsSelect.Query(qrySelect)
    If Not IsNull(rMin(0, 0)) Then
        MinValue = CSng(rMin(0, 0))
        Erase rMin
    Else
        MinValue = 0
    End If
    Set clsSelect = Nothing
End Function

Private Function MaxValue(strField As String) As Single
    Dim qrySelect As String, rMax As Variant

    Set clsSelect = New clsSelect
    
    If strField = "RMR" Then
        qrySelect = "SELECT MAX(rmr) FROM ( "
        If typCustomer.intAge >= ADULT_AGE Then
            qrySelect = qrySelect & "SELECT CASE AdBasicDsa "
        Else
            qrySelect = qrySelect & "SELECT CASE BaBasicDsa "
        End If
        qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END AS rmr "
        qrySelect = qrySelect & "FROM BodyData INNER JOIN CompData "
        qrySelect = qrySelect & "ON BodyData.CompDataNum=CompData.CompDataNum "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & ") a"
    Else
        qrySelect = "SELECT MAX(" & strField & ") FROM BodyData WHERE CustomerNum=" & glngCustomerNum
    End If
    
    rMax = clsSelect.Query(qrySelect)
    If Not IsNull(rMax(0, 0)) Then
        MaxValue = CSng(rMax(0, 0))
        Erase rMax
    Else
        MaxValue = 0
    End If
    
    Set clsSelect = Nothing
End Function

'처방내역조회 탭의 서브버튼들의 클릭이벤트
Private Sub Bottom2(intIndex As Integer)
    Dim i As Integer
    Dim qrySelect As String, rValue As Variant

    grdTable.Visible = True
    Chart.Visible = False
    cmdSub(0).Visible = True
    cmdSub(1).Visible = True
    For i = 2 To 9
        cmdSub(i).Visible = False
    Next i
    cmdSub(0).Caption = "처      치"
    cmdSub(1).Caption = "출  력  물"

    Select Case intIndex
    Case 0  '처치 테이블 처방내역조회 테이블로 초기화
        qrySelect = "SELECT TreatName, TreatCode FROM TreatCode;"
    Case 1  '처치 테이블 출력물내역조회 테이블로 초기화
        qrySelect = "SELECT PrintoutName, PrintoutNum FROM Printout;"
    Case Else
        qrySelect = "SELECT TreatName, TreatCode FROM TreatCode;"
    End Select
    
    With grdTable
        Set clsSelect = New clsSelect
        .Clear
        .Rows = 3
        .RowHeight(1) = 0
        .TextMatrix(0, 0) = "진료일"
        rValue = clsSelect.Query(qrySelect)

        If Not IsNull(rValue) Then
            .Cols = UBound(rValue, 2) + 3
            For i = 0 To UBound(rValue, 2)
                .TextMatrix(0, i + 1) = Trim(rValue(0, i))
                .TextMatrix(1, i + 1) = Trim(rValue(1, i))
            Next i
        Else
            .Cols = 2
            .Rows = 3
        End If
        For i = 0 To .Cols - 2
            .ColWidth(i) = 1000
        Next i
        .ColWidth(.Cols - 1) = 0
        Set clsSelect = Nothing
    End With
    
    '해당 데이터(처방내역조회/출력물내역조회) 표시
    Select Case intIndex
    Case 0  '처치 테이블 처방내역조회
            Call ShowTable_TreatCode
    Case 1  '처치 테이블 출력물내역조회
            Call ShowTable_Print
    Case Else
            Call ShowTable_TreatCode
    End Select
    Call SetGridRow
End Sub

Private Function LoadCustomerInfo() As Boolean
    Dim qrySelect As String, rValue As Variant
On Error GoTo SelErr
    Set clsSelect = New clsSelect

    '환자기본정보
    qrySelect = "SELECT CustomName, JuminNum, Age, Sex, ChPgCode FROM CustomerInfo WHERE CustomerNum=" & glngCustomerNum

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        Me.Enabled = True
        With typCustomer
            .strCustomName = Trim(rValue(0, 0))
            .strJuminNum = Trim(rValue(1, 0))
            .intAge = CInt(rValue(2, 0))
            .strSex = Trim(rValue(3, 0))
            .strChPgCode = Is_Null(rValue(4, 0), 0)
        End With
'        '대용식 칼로리를 가져온다
'        qrySelect = "SELECT ChmealCalory "
'        qrySelect = qrySelect & " FROM Lhmeal A INNER JOIN ChangeMeal B "
'        qrySelect = qrySelect & " ON A.ReMealName = B.ChMealName "
'        qrySelect = qrySelect & " WHERE CustomerNum = " & glngCustomerNum
'        rValue = clsSelect.Query(qrySelect)
'        If Not IsNull(rValue) Then
'            msngChMealCalory = Trim(rValue(0, 0))
'        Else
'            msngChMealCalory = 0
'        End If
        
        '식단정보를 가져온다.
        qrySelect = "SELECT A.EatTime, A.ReMealName, "
        qrySelect = qrySelect & " Case A.EatTime When 10 Then C.Calory Else ChmealCalory End "
        qrySelect = qrySelect & " FROM Lhmeal A LEFT JOIN ChangeMeal B "
        qrySelect = qrySelect & " ON A.ReMealName = B.ChMealName "
        qrySelect = qrySelect & " LEFT JOIN UserMenu C "
        qrySelect = qrySelect & " ON A.ReMealName = C.UserMenuName "
        qrySelect = qrySelect & " WHERE CustomerNum = " & glngCustomerNum
        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            mlngChMealTime = Trim(rValue(0, 0))
            If mlngChMealTime <> 0 Then
                msChMealCode = Trim(rValue(1, 0))
                msngChMealCalory = Trim(rValue(2, 0))
                chkChMeal.Value = vbChecked
                Call setCboChMealTime(mlngChMealTime)
                Call setCboChMealCode(msChMealCode)
            Else
                msChMealCode = ""
                msngChMealCalory = 0
            End If
        Else
            chkChMeal.Value = vbUnchecked
            mlngChMealTime = 0
            msChMealCode = ""
            msngChMealCalory = 0
        End If
    Else
        MsgBox "신체계측/처방을 입력할 환자를 선택하십시오." & vbNewLine & vbNewLine & "처음 방문한 환자이면 '환자등록'을 먼저 하십시오.", vbOKOnly + vbCritical
        LoadCustomerInfo = False
        Me.Enabled = False
        Exit Function
    End If
    
    LoadCustomerInfo = True
    Set clsSelect = Nothing
    Exit Function
SelErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "LoadCustomerInfo", "frmCounsel_1", Err.Number, Err.Description
    LoadCustomerInfo = False
End Function

Private Sub setCboChMealTime(lngChMealTime As Long)
Dim i As Long
If lngChMealTime = 0 Then
    Exit Sub
End If
For i = 0 To cboChMealTime.ListCount - 1
    If cboChMealTime.ItemData(i) = lngChMealTime Then
        cboChMealTime.ListIndex = i
        Exit For
    End If
Next
End Sub

Private Sub setCboChMealCode(sChMealCode As String)
Dim i As Long
If sChMealCode = "" Then
    Exit Sub
End If
For i = 0 To cboChMealCode.ListCount - 1
    If Left(cboChMealCode.List(i), InStr(cboChMealCode.List(i), "-") - 1) = sChMealCode Then
        cboChMealCode.ListIndex = i
        Exit For
    End If
Next
End Sub

'칼로리처방 데이터 초기화
Private Sub InitialTreatCalory()
    Dim i As Integer
    Dim intExTime As Integer
'200 / 300 / 400 / 500
'빨리걷기 : 0.093 / 근력운동 : 0.105
    typTreatCal.intDietCal = 0
    typTreatCal.intExCal = 0
    typTreatCal.intExDay = 0
    typTreatCal.intUserCal = 0
    typTreatCal.sngLossWeight = 0
    If typObesity.sngWeight = 0 Then
        For i = 0 To 3
            lblAerobic(i).Caption = ""
            lblAnaerobic(i).Caption = ""
        Next i
        lblLossWeight.Caption = ""
        Exit Sub
    End If
    '유산소운동과 근력운동 보여주기
    For i = 0 To 3
        intExTime = ((i + 2) * 100) / (typObesity.sngWeight * 0.093)
        lblAerobic(i).Caption = intExTime & "분"
        
        intExTime = ((i + 2) * 100) / (typObesity.sngWeight * 0.105)
        lblAnaerobic(i).Caption = intExTime & "분"
    Next i
    
    For i = 0 To 5
        Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
    Next i
    For i = 0 To 3
        Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "일-gray.jpg")
    Next i
    For i = 0 To 3
        Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
    Next i
    lblLossWeight.Caption = ""
    If typTreatCal.sngTreatCal <> 0 Then    '처방한 식단열량
        lblTreatCal.Caption = Format(typTreatCal.sngTreatCal, "#,###") & " kcal"
    End If
    Call chkUserCalory_Click
End Sub

Private Sub InitialSpread4()
'처방입력하는 스프레드 초기화(치료종목/출력물 화면의 치료종목)
'일단 모든 처방을 불러오고
'예약되어 있는 처방이 있을 시에는 그것에 체크
    Dim qrySelect As String
    Dim rValue As Variant, i As Integer, j As Integer

    With sprTreat
        .EditEnterAction = EditEnterActionDown
        .EditModePermanent = True
        
        .GrayAreaBackColor = vbWhite
        .BackColor = &HCEF7E7
        .GridColor = vbWhite
        .Font.Size = 7
        .RowHeight(-1) = 13
        .MaxCols = 4
        .ColWidth(1) = 9
        .ColWidth(2) = 2
        .ColWidth(3) = 2
        .ColWidth(4) = 0
        .Col = 2: .Row = -1
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .DisplayColHeaders = False
        .DisplayRowHeaders = False
        .ScrollBars = ScrollBarsNone

        Set clsSelect = New clsSelect
        rValue = clsSelect.Query("SELECT TreatName, TreatCode FROM TreatCode;")
        If Not IsNull(rValue) Then
            .MaxRows = UBound(rValue, 2) + 1
            If .MaxRows > 6 Then
                .ScrollBars = ScrollBarsVertical
                .ColWidth(1) = 8.5
                .Width = 1800
            Else
                .ScrollBars = ScrollBarsNone
                .ColWidth(1) = 9
                .Width = 1650
            End If
            .Col = 1
            For i = 0 To UBound(rValue, 2)
                .Row = i + 1
                .Col = 1: .Text = Trim(rValue(0, i)): .Lock = True
                .Col = 2: .CellType = CellTypeCheckBox: .Value = False
                .TypeCheckCenter = True
                .TypeCheckType = TypeCheckTypeNormal
                .Col = 3: .CellType = CellTypeCheckBox: .Value = False
                .Col = 4: .Text = Trim(rValue(1, i))
            Next i
        Else
            .MaxRows = 1
            .Row = 1: .Col = 1: .Text = ""
            .Col = 2: .CellType = CellTypeCheckBox: .Value = False
            .Col = 3: .CellType = CellTypeCheckBox: .Value = False
            .Col = 4: .Text = ""
            Exit Sub
        End If

        '예약되어 있는지 확인하고 있을시에는 해당 치료명은 체크한다
        '셀의 색을 달리한다.
        qrySelect = "SELECT ReserveDetail.ReDetailNum, ReserveTreat.TreatCode "
        qrySelect = qrySelect & "FROM ReserveDetail INNER JOIN "
        qrySelect = qrySelect & "Reserve ON ReserveDetail.ReserveNum = Reserve.ReserveNum "
        qrySelect = qrySelect & "INNER JOIN ReserveTreat ON "
        qrySelect = qrySelect & "ReserveDetail.ReDetailNum = ReserveTreat.ReDetailNum "
        qrySelect = qrySelect & "WHERE ReserveDetail.ReserveDay='" & Format(gdatUserDay, "YYYYMMDD") & "' "
        qrySelect = qrySelect & "AND Reserve.CustomerNum = " & glngCustomerNum

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            For i = 0 To UBound(rValue, 2)
                For j = 1 To .MaxRows
                    .Row = j: .Col = 1
                    If Trim(.Text) = Trim(rValue(1, i)) Then
                        .Col = 2: .CellType = CellTypeCheckBox: .Value = True
                        Exit For
                    End If
                Next j
            Next i
        End If
        Set clsSelect = Nothing
        .Lock = True
    End With
End Sub

Private Sub InitialSpread5()
'처방입력하는 스프레드 초기화(치료종목/출력물 화면의 출력물)
'일단 모든 출력물을 불러오고
'예약되어 있는 출력물이 있을 시에는 그것에 체크
    Dim qrySelect As String
    Dim rValue As Variant, i As Integer, j As Integer

    With sprPrint
        .EditEnterAction = EditEnterActionDown
        .EditModePermanent = True

        .GrayAreaBackColor = vbWhite
        .BackColor = &HCEF7E7
        .GridColor = vbWhite
        .Font.Size = 7
        .RowHeight(-1) = 13
        .MaxCols = 3
        .ColWidth(1) = 10
        .ColWidth(2) = 2
        .ColWidth(3) = 0
        .DisplayColHeaders = False
        .DisplayRowHeaders = False
        .ScrollBars = ScrollBarsNone

        Set clsSelect = New clsSelect
        rValue = clsSelect.Query("SELECT PrintoutName, PrintoutNum FROM PrintOut;")
        If Not IsNull(rValue) Then
            .MaxRows = UBound(rValue, 2) + 1
            If .MaxRows > 6 Then
                .ScrollBars = ScrollBarsVertical
                .ColWidth(1) = 9.5
                .Width = 1650
            Else
                .ScrollBars = ScrollBarsNone
                .ColWidth(1) = 10
                .Width = 1600
            End If
            .Col = 1
            For i = 0 To UBound(rValue, 2)
                .Row = i + 1
                .Col = 1: .Text = Trim(rValue(0, i)): .Lock = True
                .Col = 2: .CellType = CellTypeCheckBox: .Value = False
                .Col = 3: .Text = Trim(rValue(1, i))
            Next i
        Else
            .MaxRows = 1
            .Row = 1: .Col = 1: .Text = ""
            .Col = 2: .CellType = CellTypeCheckBox: .Value = False
            .Col = 3: .Text = ""
            Exit Sub
        End If

        '예약되어 있는 출력물에는 일단 체크하기
        '예약되어 있는지 확인하고 있을시에는 해당 출력물명은 체크한다
        qrySelect = "SELECT ReserveDetail.ReDetailNum, ReservePrint.PrintCode "
        qrySelect = qrySelect & "FROM ReserveDetail INNER JOIN "
        qrySelect = qrySelect & "Reserve ON ReserveDetail.ReserveNum=Reserve.ReserveNum "
        qrySelect = qrySelect & "INNER JOIN ReservePrint ON ReserveDetail.ReDetailNum=ReservePrint.ReDetailNum "
        qrySelect = qrySelect & "WHERE ReserveDetail.ReserveDay='" & Format(gdatUserDay, "YYYYMMDD") & "' "
        qrySelect = qrySelect & "AND Reserve.CustomerNum=" & glngCustomerNum
        
        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            For i = 0 To UBound(rValue, 2)
                For j = 1 To .MaxRows
                    .Row = j: .Col = 1
                    If Trim(.Text) = Trim(rValue(1, i)) Then
                        .Col = 2: .CellType = CellTypeCheckBox: .Value = True
                        Exit For
                    End If
                Next j
            Next i
        End If
        Set clsSelect = Nothing
    End With
End Sub

Private Sub InitialChart()
    With Chart
        .Gallery = LINES
        .Chart3D = False
        .MarkerShape = MK_RECT
        .MarkerSize = 2
        .AxesStyle = CAS_FLATFRAME
        .Axis(0).Grid = True

        ' Color Settings
        .Border = False
        .RGBBk = vbWhite

        ' Layout Settings
        .LegendBox = False
        .SerLegBox = False
        .ToolBar = False
        .PointLabels = True
        .MultipleColors = False
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Function InputValues(sngValue As Single) As String
    Dim strQuery As String

    If sngValue = 0 Then
        strQuery = "NULL"
    Else
        strQuery = CStr(sngValue)
    End If
    InputValues = strQuery
End Function

Private Sub ShowTreatData(lngTreatNum As Long)
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer

    Set clsSelect = New clsSelect
    For j = 0 To sprTreat.MaxRows
        sprTreat.Row = j
        sprTreat.Col = 2
        sprTreat.Value = 0
        sprTreat.Col = 3
        sprTreat.Value = 0
    Next j

    qrySelect = "SELECT TreatCode, Execution FROM TreatData WHERE TreatNum=" & lngTreatNum
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            For j = 0 To sprTreat.MaxRows
                sprTreat.Col = 4
                sprTreat.Row = j
                If Trim(sprTreat.Text) = Trim(rValue(0, i)) Then
                    sprTreat.Col = 2
                    sprTreat.Value = 1
                    sprTreat.Col = 3: sprTreat.Lock = False
                    If Trim(rValue(1, i)) = "Y" Then
                        sprTreat.Col = 3
                        sprTreat.Value = 1
                    End If
                    Exit For
                End If
            Next j
        Next i
    End If

    Set clsSelect = Nothing
End Sub

Private Sub ShowTable_Circumference()
'중간에 사용자가 추가한 입력값이 있으면...?
'BioChemData에서 해당값을 가져온다. TreatNum으로..
'근데 순서를 어떻게 맞추나..중간에 사용자 추가값이 있으면..
    Dim qrySelect As String, rValue As Variant
    Dim rUValue As Variant, intUEnd As Integer, intUCount As Integer
    Dim i As Integer, j As Integer
    
    Set clsSelect = New clsSelect
    
    With grdTable
    intUCount = 0
    qrySelect = "SELECT TreatDay"
    For i = 1 To .Cols - 2
        qrySelect = qrySelect & "," & Trim(.TextMatrix(1, i + 1))
        If IsNumeric(.TextMatrix(1, i + 1)) = True Then
            intUEnd = i + 1
            intUCount = intUCount + 1
        End If
    Next i
    qrySelect = qrySelect & " FROM BodyData INNER JOIN Treat ON "
    qrySelect = qrySelect & "BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY TreatDay ASC;"
    
    rValue = clsSelect.Query(qrySelect)
    
    If Not IsNull(rValue) Then
        .Rows = .Rows + UBound(rValue, 2)
        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = i + 1
            For j = 0 To .Cols - 2
                If Not IsNull(rValue(j, i)) Then
                    .TextMatrix(i + 2, j + 1) = Trim(rValue(j, i))
                Else
                    .TextMatrix(i + 2, j + 1) = "-"
                End If
            Next j
            '사용자 추가 입력값
            If intUCount > 0 Then
                For j = (intUEnd - intUCount + 1) To intUEnd
                    qrySelect = "SELECT BioChemSu FROM BioChemData "
                    qrySelect = qrySelect & "WHERE BioChemCode='" & Trim(.TextMatrix(1, j))
                    qrySelect = qrySelect & "' AND TreatNum=" & Trim(.TextMatrix(i + 2, .Cols - 1))
                    rUValue = clsSelect.Query(qrySelect)
                    If Not IsNull(rUValue) Then
                        .TextMatrix(i + 2, j) = Trim(rUValue(0, 0))
                    Else
                        .TextMatrix(i + 2, j) = "-"
                    End If
                Next j
            End If
            If rValue(0, i) = gdatUserDay Then
                .Row = i + 2: .Col = 0: .ColSel = .Cols - 1
                glngTreatNum = .TextMatrix(i + 2, .Cols - 1)
            End If
        Next i
    End If
    End With
    
    Set clsSelect = Nothing
    
End Sub

Private Sub ShowTreatPrint(lngTreatNum As Long)
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer

    Set clsSelect = New clsSelect
    For j = 0 To sprPrint.MaxRows
        sprPrint.Row = j
        sprPrint.Col = 2
        sprPrint.Value = 0
    Next j

    qrySelect = "SELECT PrintoutNum FROM TreatPrint WHERE TreatNum=" & lngTreatNum
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            For j = 0 To sprPrint.MaxRows
                sprPrint.Col = 3
                sprPrint.Row = j
                If Trim(sprPrint.Text) = Trim(rValue(0, i)) Then
                    sprPrint.Col = 2
                    sprPrint.Value = 1
                    Exit For
                End If
            Next j
        Next i
    End If

    Set clsSelect = Nothing
End Sub

Private Function SaveTreat(Optional datUserDay As Date = "1900-01-01") As Long
    Dim qryInsert As String
    Dim lngTreatNum As Long
    Dim sngChMealCalory As Single
On Error GoTo InsertErr
    If datUserDay = "1900-01-01" Then
        datUserDay = gdatUserDay
    End If
    
    If chkUserCalory.Value = vbChecked Then '식단칼로리 직접입력
        If IsNumeric(txtUserCalory.Text) Then
            '이경우 식단 처방을 따로 했는지 체크
            If chkChMeal.Value = vbChecked Then '식단 처방을 했을 경우
                If cboChMealTime.ItemData(cboChMealTime.ListIndex) = 10 Then    '사용자 식단인경우
                    If cboChMealCode.ListCount = 0 Then
                        MsgBox "사용자 식단이 없습니다. 먼저 사용자 식단을 입력하세요", vbOKOnly + vbExclamation
                        SaveTreat = -1
                        Exit Function
                    Else
                        If CInt(txtUserCalory.Text) < 1000 Or CInt(txtUserCalory.Text) > 3500 Then
                            '직접입력 칼로리가 범위를 벗어남.
                            MsgBox "식단열량은 1000kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
                            txtUserCalory.SelStart = 0
                            txtUserCalory.SelLength = Len(txtUserCalory)
                            txtUserCalory.SetFocus
                            SaveTreat = -1
                            Exit Function
                        End If
                        msngChMealCalory = CSng(cboChMealCode.ItemData(cboChMealCode.ListIndex))
                        msChMealCode = Left(cboChMealCode.List(cboChMealCode.ListIndex), InStr(cboChMealCode.List(cboChMealCode.ListIndex), "-") - 1)
                        mlngChMealTime = 10
                    End If
                ElseIf cboChMealTime.ItemData(cboChMealTime.ListIndex) = 0 Then '구분선을 선택한 경우
                    MsgBox "식단처방을 선택하세요.", vbOKOnly + vbExclamation
                    cboChMealTime.SetFocus
                    SaveTreat = -1
                    Exit Function
                Else    '대용식을 선택한 경우
                    If cboChMealCode.ListCount = 0 Then
                        MsgBox "대용식이 없습니다. 먼저 대용식을 입력하세요", vbOKOnly + vbExclamation
                        SaveTreat = -1
                        Exit Function
                    Else
                        sngChMealCalory = CSng(cboChMealCode.ItemData(cboChMealCode.ListIndex))
                        If CInt(txtUserCalory.Text) - sngChMealCalory < 1000 Or CInt(txtUserCalory.Text) + sngChMealCalory > 3500 Then
                            '직접입력 칼로리가 범위를 벗어남.
                            MsgBox "식단열량은 " & Format(1000 + sngChMealCalory, "0") & "kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
                            txtUserCalory.SelStart = 0
                            txtUserCalory.SelLength = Len(txtUserCalory)
                            txtUserCalory.SetFocus
                            SaveTreat = -1
                            Exit Function
                        End If
                        msngChMealCalory = sngChMealCalory
                        msChMealCode = Left(cboChMealCode.List(cboChMealCode.ListIndex), InStr(cboChMealCode.List(cboChMealCode.ListIndex), "-") - 1)
                        mlngChMealTime = cboChMealTime.ItemData(cboChMealTime.ListIndex)
                    End If
                End If
            Else                            '식단 처방을 하지 않았을 경우
                If CInt(txtUserCalory.Text) < 1000 Or CInt(txtUserCalory.Text) > 3500 Then
                    '직접입력 칼로리가 범위를 벗어남.
                    MsgBox "식단열량은 1000kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
                    txtUserCalory.SelStart = 0
                    txtUserCalory.SelLength = Len(txtUserCalory)
                    txtUserCalory.SetFocus
                    SaveTreat = -1
                    Exit Function
                Else
                    msChMealCode = "NULL"
                    mlngChMealTime = 0
                End If
            End If
            typTreatCal.intUserCal = CInt(txtUserCalory.Text)
        Else
            MsgBox "식단열량은 " & Format(1000 + msngChMealCalory, "0") & "kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
            txtUserCalory.SelStart = 0
            txtUserCalory.SelLength = Len(txtUserCalory)
            txtUserCalory.SetFocus
            SaveTreat = -1
            Exit Function
        End If
    Else        '식단칼로리 직접입력 아님.
        typTreatCal.intUserCal = 0
        msChMealCode = "NULL"
        mlngChMealTime = 0
    End If
    
'    If txtUserCalory.Text = "" Then
'        If chkUserCalory.Value = vbChecked Then
'            MsgBox "식단열량은 " & Format(1000 + msngChMealCalory, "0") & "kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
'            txtUserCalory.SelStart = 0
'            txtUserCalory.SelLength = Len(txtUserCalory)
'            txtUserCalory.SetFocus
'            SaveTreat = -1
'            Exit Function
'        Else
'            typTreatCal.intUserCal = 0
'        End If
'    Else
'        If IsNumeric(txtUserCalory.Text) = True Then
'            If CInt(txtUserCalory.Text) - msngChMealCalory < 1000 Or CInt(txtUserCalory.Text) + msngChMealCalory > 3500 Then
'                MsgBox "식단열량은 " & Format(1000 + msngChMealCalory, "0") & "kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
'                txtUserCalory.SelStart = 0
'                txtUserCalory.SelLength = Len(txtUserCalory)
'                txtUserCalory.SetFocus
'                SaveTreat = -1
'                'SaveTreat = 0
'                Exit Function
'            Else
'                typTreatCal.intUserCal = CInt(txtUserCalory.Text)
'            End If
'        Else
'            typTreatCal.intUserCal = 0
'        End If
'    End If
    If typTreatCal.sngTreatCal < typObesity.sngRMR Then
        If MsgBox("처방한 " & Format(typTreatCal.sngTreatCal, "0") & _
            " kcal는 기초대사량 " & Format(typObesity.sngRMR, "0") & _
            " kcal보다 적습니다." & vbNewLine & vbNewLine & _
            "기초대사량보다 적은 열량의 무리한 절식은 건강에 해롭습니다." & vbNewLine & vbNewLine & _
            "계속 진행하시겠습니까?", vbYesNo + vbExclamation, "처방열량") = vbNo Then
            SaveTreat = -1
            Exit Function
        End If
    End If
    lngTreatNum = WhatisCode("Treat", "TreatNum")

    qryInsert = "INSERT INTO Treat("
    qryInsert = qryInsert & "TreatNum,"
    qryInsert = qryInsert & "CustomerNum,"
    qryInsert = qryInsert & "TreatDay, "
    qryInsert = qryInsert & "LossWeight,"
    qryInsert = qryInsert & "ExDay,"
    qryInsert = qryInsert & "ExCalory,"
    qryInsert = qryInsert & "DietCalory, "
    qryInsert = qryInsert & "TreatCalory,"
    qryInsert = qryInsert & "Notice,"
    qryInsert = qryInsert & "Memo,"
    qryInsert = qryInsert & "main,"
    qryInsert = qryInsert & "sub1,"
    qryInsert = qryInsert & "sub2,"
    qryInsert = qryInsert & "sub3,"
    qryInsert = qryInsert & "sub4,"
    qryInsert = qryInsert & "UserCalory,"
    qryInsert = qryInsert & "ChMealTime, "
    qryInsert = qryInsert & "ChMealCode "
    qryInsert = qryInsert & ") "
    qryInsert = qryInsert & "VALUES(" & lngTreatNum
    qryInsert = qryInsert & "," & glngCustomerNum
    qryInsert = qryInsert & ",'" & Format(datUserDay, "YYYYMMDD") & "'"
    qryInsert = qryInsert & "," & InputValues(typTreatCal.sngLossWeight)
    qryInsert = qryInsert & "," & InputValues(CSng(typTreatCal.intExDay))
    qryInsert = qryInsert & "," & InputValues(CSng(typTreatCal.intExCal))
    qryInsert = qryInsert & "," & InputValues(CSng(typTreatCal.intDietCal))
    qryInsert = qryInsert & "," & InputValues(typTreatCal.sngTreatCal)
    qryInsert = qryInsert & ",'" & Trim(txtNotice.Text) & "'"
    qryInsert = qryInsert & ",'" & Trim(txtMemo.Text) & "'"
    qryInsert = qryInsert & "," & InputValues(CSng(gintMain))
    qryInsert = qryInsert & "," & InputValues(CSng(gintSub1))
    qryInsert = qryInsert & "," & InputValues(CSng(gintSub2))
    qryInsert = qryInsert & "," & InputValues(CSng(gintSub3))
    qryInsert = qryInsert & "," & InputValues(CSng(gintSub4))
    qryInsert = qryInsert & "," & InputValues(CInt(typTreatCal.intUserCal))
    qryInsert = qryInsert & "," & CLng(mlngChMealTime)
    If msChMealCode = "NULL" Then
        qryInsert = qryInsert & ",NULL"
    Else
        qryInsert = qryInsert & ",'" & Trim(msChMealCode) & "'"
    End If
    qryInsert = qryInsert & ");"
    modSql.AdoExcuteSql qryInsert

    SaveTreat = lngTreatNum
    Exit Function
InsertErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "SaveTreat", "frmCounsel_1", Err.Number, Err.Description
    SaveTreat = 0
End Function

'리턴값 :   -1  => 입력한 열량 오류
'           0   => 저장실패
'           1   => 저장 성공
Private Function UpdateTreat(lngTreatNum As Long) As Long
    Dim qryUpdate As String

On Error GoTo UpdateErr
    'If txtUserCalory.Text = "" Then
    If chkUserCalory.Value = vbUnchecked Then
        typTreatCal.intUserCal = 0
    Else
        If IsNumeric(txtUserCalory.Text) Then
            If chkChMeal.Value = vbChecked Then
                If CLng(cboChMealTime.ItemData(cboChMealTime.ListIndex)) = 10 Then
                    If CInt(txtUserCalory) < 1000 Or CInt(txtUserCalory.Text) > 3500 Then
                        MsgBox "식단열량은 " & Format(1000 + msngChMealCalory, "0") & "kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
                        txtUserCalory.SelStart = 0
                        txtUserCalory.SelLength = Len(txtUserCalory)
                        txtUserCalory.SetFocus
                        UpdateTreat = -1
                        Exit Function
                    Else
                        typTreatCal.intUserCal = CInt(txtUserCalory.Text)
                    End If
                
                ElseIf CLng(cboChMealTime.ItemData(cboChMealTime.ListIndex)) = 0 Then
                    MsgBox "식단처방을 선택하세요.", vbOKOnly + vbExclamation
                    cboChMealTime.SetFocus
                    UpdateTreat = -1
                    Exit Function
                Else
                    If CInt(txtUserCalory.Text) - cboChMealCode.ItemData(cboChMealCode.ListIndex) < 1000 Or CInt(txtUserCalory.Text) > 3500 Then
                        MsgBox "식단열량은 " & Format(1000 + msngChMealCalory, "0") & "kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
                        txtUserCalory.SelStart = 0
                        txtUserCalory.SelLength = Len(txtUserCalory)
                        txtUserCalory.SetFocus
                        UpdateTreat = -1
                        Exit Function
                    Else
                        typTreatCal.intUserCal = CInt(txtUserCalory.Text)
                    End If
                End If
            Else
                If CInt(txtUserCalory.Text) < 1000 Or CInt(txtUserCalory.Text) > 3500 Then
                    MsgBox "식단열량은 1000kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
                    txtUserCalory.SelStart = 0
                    txtUserCalory.SelLength = Len(txtUserCalory)
                    txtUserCalory.SetFocus
                    UpdateTreat = -1
                    Exit Function
                Else
                    typTreatCal.intUserCal = CInt(txtUserCalory.Text)
                End If
            End If
        Else
            typTreatCal.intUserCal = 0
        End If
    End If

    If typTreatCal.sngTreatCal < typObesity.sngRMR Then
        If MsgBox("처방한 " & Format(typTreatCal.sngTreatCal, "0") & _
            " kcal는 기초대사량 " & Format(typObesity.sngRMR, "0") & _
            " kcal보다 적습니다." & vbNewLine & vbNewLine & _
            "기초대사량보다 적은 열량의 무리한 절식은 건강에 해롭습니다." & vbNewLine & vbNewLine & _
            "계속 진행하시겠습니까?", vbYesNo + vbExclamation, "처방열량") = vbNo Then
            UpdateTreat = -1
            Exit Function
        End If
    End If
    
    qryUpdate = "UPDATE Treat SET "
    qryUpdate = qryUpdate & "LossWeight=" & InputValues(typTreatCal.sngLossWeight)
    qryUpdate = qryUpdate & ", ExDay=" & InputValues(CSng(typTreatCal.intExDay))
    qryUpdate = qryUpdate & ", ExCalory=" & InputValues(CSng(typTreatCal.intExCal))
    qryUpdate = qryUpdate & ", DietCalory=" & InputValues(CSng(typTreatCal.intDietCal))
    qryUpdate = qryUpdate & ", TreatCalory=" & InputValues(typTreatCal.sngTreatCal)
    qryUpdate = qryUpdate & ", Notice='" & Trim(txtNotice.Text) & "'"
    qryUpdate = qryUpdate & ", Memo='" & Trim(txtMemo.Text) & "'"
    qryUpdate = qryUpdate & ", main=" & InputValues(CSng(gintMain))
    qryUpdate = qryUpdate & ", sub1=" & InputValues(CSng(gintSub1))
    qryUpdate = qryUpdate & ", sub2=" & InputValues(CSng(gintSub2))
    qryUpdate = qryUpdate & ", sub3=" & InputValues(CSng(gintSub3))
    qryUpdate = qryUpdate & ", sub4=" & InputValues(CSng(gintSub4))
    
    If chkUserCalory.Value = 1 Then
        qryUpdate = qryUpdate & ", UserCalory=" & InputValues(typTreatCal.intUserCal)
    Else
        qryUpdate = qryUpdate & ", UserCalory=Null"
    End If
    
    If chkChMeal.Value = vbChecked Then '대용식 혹은 사용자정의 식단을 처방한경우
        qryUpdate = qryUpdate & ", ChMealTime = " & InputValues(CSng(cboChMealTime.ItemData(cboChMealTime.ListIndex))) & " "
        qryUpdate = qryUpdate & ", ChMealCode = '" & Left(cboChMealCode.List(cboChMealCode.ListIndex), InStr(cboChMealCode.List(cboChMealCode.ListIndex), "-") - 1) & "' "
    Else                                '대용식 혹은 사용자정의 식단을 처방하지 않은 경우
        qryUpdate = qryUpdate & ", ChMealTime = 0 "
        qryUpdate = qryUpdate & ", ChMealCode = NULL "
    End If
    
    qryUpdate = qryUpdate & " WHERE TreatNum=" & lngTreatNum & ";"

    modSql.AdoExcuteSql (qryUpdate)
    UpdateTreat = 1
    Exit Function
UpdateErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "UpdateTreat", "frmCounsel_1", Err.Number, Err.Description
    UpdateTreat = 0
End Function

Private Function SaveTreatData(Optional lngTreatNum As Long = 0) As Boolean
    Dim qryInsert As String, qryDelete As String
    Dim lngTreatDataNum As Long, i As Integer

On Error GoTo InsertErr
    
    If lngTreatNum = 0 Then
        lngTreatNum = glngTreatNum
    End If
    qryDelete = "DELETE FROM TreatData WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)

    lngTreatDataNum = WhatisCode("TreatData", "TreatDataNum")
    With sprTreat
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2
            If .Value <> 0 Then
                qryInsert = "INSERT INTO TreatData(TreatDataNum, ReDetailNum, TreatNum, TreatCode, Execution) "
                qryInsert = qryInsert & "VALUES(" & lngTreatDataNum
                qryInsert = qryInsert & ",0"    '******* 손볼것
                qryInsert = qryInsert & "," & lngTreatNum
                .Col = 4
                qryInsert = qryInsert & "," & Trim(.Text)
                .Col = 3
                If .Value = 1 Then
                    qryInsert = qryInsert & ",'Y');"
                Else
                    qryInsert = qryInsert & ",'N');"
                End If

                modSql.AdoExcuteSql (qryInsert)
                lngTreatDataNum = lngTreatDataNum + 1
            End If
        Next i
    End With
    SaveTreatData = True
    Exit Function
InsertErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "SaveTreatData", "frmCounsel_1", Err.Number, Err.Description
    SaveTreatData = False

End Function

Private Function SaveTreatPrint(Optional lngTreatNum As Long = 0) As Boolean
    Dim qryInsert As String, qryDelete As String
    Dim lngTreatPrintNum As Long, i As Integer
On Error GoTo InsertErr
    If lngTreatNum = 0 Then
        lngTreatNum = glngTreatNum
    End If
    qryDelete = "DELETE FROM TreatPrint WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)

    lngTreatPrintNum = WhatisCode("TreatPrint", "TreatPrintNum")
    With sprPrint
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2
            If .Value <> 0 Then
                qryInsert = "INSERT INTO TreatPrint(TreatPrintNum, ReDetailNum, TreatNum, PrintoutNum) "
                qryInsert = qryInsert & "VALUES(" & lngTreatPrintNum
                qryInsert = qryInsert & ",0"      '******* 같이손볼곳
                qryInsert = qryInsert & "," & glngTreatNum
                .Col = 3
                qryInsert = qryInsert & "," & Trim(.Text) & ");"

                modSql.AdoExcuteSql (qryInsert)
                lngTreatPrintNum = lngTreatPrintNum + 1
            End If
        Next i
    End With
    SaveTreatPrint = True
    Exit Function
InsertErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "SaveTreatPrint", "frmCounsel_1", Err.Number, Err.Description
    SaveTreatPrint = False

End Function

Private Function DeleteTreat(lngTreatNum As Long) As Boolean
    Dim qryDelete As String
    
On Error GoTo DelErr
    'TreatData  삭제
    qryDelete = "DELETE FROM TreatData WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)
    
    'TreatPrint 삭제
    qryDelete = "DELETE FROM TreatPrint WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)
    
    'BodyData 삭제
    qryDelete = "DELETE FROM BodyData WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)

    'Treat 삭제
    qryDelete = "DELETE FROM Treat WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)
    
    DeleteTreat = True
    Exit Function
DelErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "DeleteTreat", "frmCounsel_1", Err.Number, Err.Description
    DeleteTreat = False
End Function

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub grdTable_Click()
    If grdTable.Row <= 1 Then
        Exit Sub
    End If

    If grdTable.TextMatrix(grdTable.Row, grdTable.Cols - 1) = "" Or IsNumeric(grdTable.TextMatrix(grdTable.Row, grdTable.Cols - 1)) = False Then
        'glngTreatNum = 0
        Exit Sub
    End If

    mlngcTreatNum = grdTable.TextMatrix(grdTable.Row, grdTable.Cols - 1)
    If mlngcTreatNum <> 0 Then
        Call ShowTreat(mlngcTreatNum)
        Call ShowTreatData(mlngcTreatNum)
        Call ShowTreatPrint(mlngcTreatNum)
        If glngTreatNum = mlngcTreatNum Then
            intMode = 2
            Call EnabledInput(True)
            lblDispDate.Caption = Format(grdTable.TextMatrix(grdTable.Row, IIf(gintBottomButton = 0, 1, 0)), "YYYY-MM-DD")
            imgSave.Enabled = True
            imgModifyTreat.Visible = False
        Else
            intMode = 0
            Call EnabledInput(False)
            lblDispDate.Caption = Format(grdTable.TextMatrix(grdTable.Row, IIf(gintBottomButton = 0, 1, 0)), "YYYY-MM-DD")
            imgSave.Enabled = False
            imgModifyTreat.Visible = True
        End If

    End If
End Sub

Private Sub ShowTreat(lngTreatNum As Long)
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, Index As Integer
    Dim sngWeight As Single, intExTime As Integer
    Dim lngChMealTime As Long, sChMealCode As String
    Set clsSelect = New clsSelect

    qrySelect = "SELECT LossWeight, ExCalory, DietCalory, TreatCalory, ExDay, Memo "
    qrySelect = qrySelect & ",main, sub1, sub2, sub3, sub4, UserCalory "
    qrySelect = qrySelect & ",Weight, ChMealTime, ChMealCode "
    qrySelect = qrySelect & "FROM Treat LEFT JOIN BodyData "
    qrySelect = qrySelect & "ON Treat.TreatNum=BodyData.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.TreatNum=" & lngTreatNum
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        If Not IsNull(rValue(0, 0)) Then
            typTreatCal.sngLossWeight = Is_Null(rValue(0, 0), 0)
            typTreatCal.intExCal = Is_Null(rValue(1, 0), 0)
            typTreatCal.intDietCal = Is_Null(rValue(2, 0), 0)
            typTreatCal.sngTreatCal = Is_Null(rValue(3, 0), typObesity.sngTEE)
            typTreatCal.intExDay = Is_Null(rValue(4, 0), 0)
            typTreatCal.intUserCal = Is_Null(rValue(11, 0), 0)
            
            txtMemo.Text = Is_Null(rValue(5, 0), "")
            
            gintMain = Is_Null(rValue(6, 0), 0)
            gintSub1 = Is_Null(rValue(7, 0), 0)
            gintSub2 = Is_Null(rValue(8, 0), 0)
            gintSub3 = Is_Null(rValue(9, 0), 0)
            gintSub4 = Is_Null(rValue(10, 0), 0)
            
            '대용식처방을 보여준다.
            lngChMealTime = Is_Null(rValue(13, 0), 0)
            sChMealCode = Trim(Is_Null(rValue(14, 0), ""))
        Else
            typTreatCal.sngLossWeight = 0
            typTreatCal.intExCal = 0
            typTreatCal.intDietCal = 0
            typTreatCal.sngTreatCal = 0
            typTreatCal.intExDay = 0
            typTreatCal.intUserCal = 0
            
            txtMemo.Text = ""
            
            gintMain = 0
            gintSub1 = 0
            gintSub2 = 0
            gintSub3 = 0
            gintSub4 = 0
            lngChMealTime = 0
            sChMealCode = ""
        End If
        
    Else
        typTreatCal.sngLossWeight = 0
        typTreatCal.intExCal = 0
        typTreatCal.intDietCal = 0
        typTreatCal.sngTreatCal = 0
        typTreatCal.intExDay = 0
        typTreatCal.intUserCal = 0
        
        txtMemo.Text = ""
        
        gintMain = 0
        gintSub1 = 0
        gintSub2 = 0
        gintSub3 = 0
        gintSub4 = 0
        lngChMealTime = 0
        sChMealCode = ""
        
    End If
    
    If Not IsNull(rValue) Then
        If Not IsNull(rValue(12, 0)) Then
            sngWeight = rValue(12, 0)
            Call ShowEx(sngWeight)
        Else
            Call ShowRecentEx
    '        '현재 날짜에 과거의 가장최근 몸무게를 가져온다.
    '        qrySelect = " SELECT Top 1 Weight FROM Treat a inner join BodyData b "
    '        qrySelect = qrySelect & " ON a.TreatNum = b.TreatNum "
    '        qrySelect = qrySelect & " WHERE a.CustomerNum = '" & glngCustomerNum & "' "
    '        qrySelect = qrySelect & " AND Weight <> '' AND Weight > 0 "
    '        qrySelect = qrySelect & " AND TreatDay < '" & Format(gdatUserDay, "YYYY-MM-DD") & "' "
    '        qrySelect = qrySelect & " ORDER BY a.TreatDay DESC, a.TreatNum DESC"
    '        rValue = clsSelect.Query(qrySelect)
    '        If Not IsNull(rValue(0, 0)) Then
    '            sngWeight = rValue(0, 0)
    '            For i = 0 To 3
    '                intExTime = ((i + 2) * 100) / (sngWeight * 0.093)
    '                lblAerobic(i).Caption = intExTime & "분"
    '
    '                intExTime = ((i + 2) * 100) / (sngWeight * 0.105)
    '                lblAnaerobic(i).Caption = intExTime & "분"
    '            Next i
    '        End If
        End If
    Else
        Call ShowRecentEx
    End If
    
    qrySelect = "SELECT Notice "
    qrySelect = qrySelect & "FROM Treat WHERE TreatNum=" & lngTreatNum
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        If Not IsNull(rValue(0, 0)) Then
            txtNotice.Text = rValue(0, 0)
        Else
            txtNotice.Text = ""
        End If
    Else
        txtNotice.Text = ""
    End If
    
    Set clsSelect = Nothing
    
    '식사칼로리
    With typTreatCal
    
        For i = 0 To 5
            Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
        Next i
        Select Case .intDietCal
            Case 125: Index = 0
            Case 250: Index = 1
            Case 375: Index = 2
            Case 500: Index = 3
            Case 750: Index = 4
            Case 1000: Index = 5
            Case Else: Index = 6
        End Select
        If Index <> 6 Then
            Set imgDietCal(Index).Picture = LoadPicture(App.Path & IMG_ME & Index & "green.jpg")
        End If
        
        '일일필요열량에 따라 감량안되는 칼로리가 있음..그런 칼로리는 아예 안보이도록 2004.04.01
        Select Case typObesity.sngTEE
            Case 1000 To 1249:      '1000~1200
                For i = 0 To 5
                    imgDietCal(i).Visible = False
                Next i
            Case 1250 To 1449:      '1300~1400
                imgDietCal(0).Visible = False
                imgDietCal(1).Visible = True
                For i = 2 To 5
                    imgDietCal(i).Visible = False
                Next i
            Case 1450 To 1749:      '1500~1700
                imgDietCal(0).Visible = False
                imgDietCal(1).Visible = True
                imgDietCal(2).Visible = False
                imgDietCal(3).Visible = True
                imgDietCal(4).Visible = False
                imgDietCal(5).Visible = False
            Case 1750 To 1949:      '1800~1900
                imgDietCal(0).Visible = False
                imgDietCal(1).Visible = True
                imgDietCal(2).Visible = False
                imgDietCal(3).Visible = True
                imgDietCal(4).Visible = True
                imgDietCal(5).Visible = False
            Case 1950 To 3500:      '2000~3500
                imgDietCal(0).Visible = False
                imgDietCal(1).Visible = True
                imgDietCal(2).Visible = False
                For i = 3 To 5
                    imgDietCal(i).Visible = True
                Next i
            Case Else
                For i = 0 To 5
                    imgDietCal(i).Visible = False
                Next i
        End Select
        
        '운동일수
        For i = 0 To 3
            Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "일-gray.jpg")
        Next i
        If .intExDay >= 3 And .intExDay <= 6 Then
            Set imgExDay(.intExDay - 3).Picture = LoadPicture(App.Path & IMG_EX & .intExDay & "일-green.jpg")
        End If
        
        '운동칼로리
        For i = 0 To 3
            Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
        Next i
        If .intExCal >= 200 And .intExCal <= 500 Then
            Set imgExCal((.intExCal / 100) - 2).Picture = LoadPicture(App.Path & IMG_EX & (.intExCal / 100) & "00-green.jpg")
        End If
        
        '감량체중
        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
        If .intUserCal = 0 Then
            txtUserCalory.Text = ""
            txtUserCalory.BackColor = FRM_GRAY
            txtUserCalory.Enabled = False
            chkUserCalory.Value = 0
        Else
            For i = 0 To 5
                Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
                imgDietCal(i).Enabled = False
            Next i
            txtUserCalory.Text = .intUserCal
            txtUserCalory.BackColor = vbWhite
            txtUserCalory.Enabled = True
            chkUserCalory.Value = 1
            .sngTreatCal = .intUserCal
        End If
        If .sngTreatCal = 0 Then
            .sngTreatCal = typObesity.sngTEE
        End If
        lblTreatCal.Caption = Format(.sngTreatCal, "#,##0") & " kcal"
        
    End With
    Call ShowChMealTreat(lngChMealTime, sChMealCode)

End Sub

Private Sub ShowChMealTreat(lngChMealTime As Long, sChMealCode As String)

    Select Case lngChMealTime
    Case 0
        
        chkChMeal.Value = vbUnchecked
        If chkUserCalory.Value = vbChecked Then
            chkChMeal.Enabled = True
        Else
            chkChMeal.Enabled = False
        End If
    Case 10
        chkChMeal.Value = vbChecked
        chkChMeal.Enabled = True
        Call setCboChMealTime(lngChMealTime)
        Call setCboChMealCode(sChMealCode)
        txtUserCalory.Enabled = False
    Case Else
        chkChMeal.Value = vbChecked
        chkChMeal.Enabled = True
        Call setCboChMealTime(lngChMealTime)
        Call setCboChMealCode(sChMealCode)
    End Select
End Sub

Private Sub ShowRecentEx()
Dim qrySelect As String
Dim i As Integer
Dim sngWeight As Single
Dim rValue As Variant
    Set clsSelect = New clsSelect
    '현재 날짜에 과거의 가장최근 몸무게를 가져온다.
    qrySelect = " SELECT Top 1 Weight FROM Treat a inner join BodyData b "
    qrySelect = qrySelect & " ON a.TreatNum = b.TreatNum "
    qrySelect = qrySelect & " WHERE a.CustomerNum = '" & glngCustomerNum & "' "
    qrySelect = qrySelect & " AND Weight <> '' AND Weight > 0 "
    qrySelect = qrySelect & " AND TreatDay < '" & Format(gdatUserDay, "YYYY-MM-DD") & "' "
    qrySelect = qrySelect & " ORDER BY a.TreatDay DESC, a.TreatNum DESC"
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue(0, 0)) Then
        sngWeight = rValue(0, 0)
        Call ShowEx(sngWeight)
    End If
End Sub

Private Sub ShowEx(sngWeight As Single)
Dim i As Integer
Dim intExTime As Integer
    If sngWeight > 0 Then
        For i = 0 To 3
            intExTime = ((i + 2) * 100) / (sngWeight * 0.093)
            lblAerobic(i).Caption = intExTime & "분"
            
            intExTime = ((i + 2) * 100) / (sngWeight * 0.105)
            lblAnaerobic(i).Caption = intExTime & "분"
        Next i
    End If
End Sub

Private Sub InitialDietCalory()
    Dim i As Integer
    
    Select Case typObesity.sngTEE
        Case 1000 To 1249:      '1000~1200
            For i = 0 To 5
                imgDietCal(i).Visible = False
            Next i
        Case 1250 To 1449:      '1300~1400
            imgDietCal(0).Visible = False
            imgDietCal(1).Visible = True
            For i = 2 To 5
                imgDietCal(i).Visible = False
            Next i
        Case 1450 To 1749:      '1500~1700
            imgDietCal(0).Visible = False
            imgDietCal(1).Visible = True
            imgDietCal(2).Visible = False
            imgDietCal(3).Visible = True
            imgDietCal(4).Visible = False
            imgDietCal(5).Visible = False
        Case 1750 To 1949:      '1800~1900
            imgDietCal(0).Visible = False
            imgDietCal(1).Visible = True
            imgDietCal(2).Visible = False
            imgDietCal(3).Visible = True
            imgDietCal(4).Visible = True
            imgDietCal(5).Visible = False
        Case 1950 To 3500:      '2000~3500
            imgDietCal(0).Visible = False
            imgDietCal(1).Visible = True
            imgDietCal(2).Visible = False
            For i = 3 To 5
                imgDietCal(i).Visible = True
            Next i
        Case Else
            For i = 0 To 5
                imgDietCal(i).Visible = False
            Next i
    End Select
End Sub

'비만관련 데이터 보여주기
'체중/비만도/허리둘레/WHR/BMI/체지방률/RMR/TEE
' 휴식대사량 8 : inRMR / 9 : etcRMR / 그밖에 : RMR
Private Sub ShowObesityRecord()
    Dim qrySelect As String, rValue As Variant
    Dim sngGab As Single
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT TOP 2 Weight, ObesityRate, Waist, b.WHR, BMI, ChFatRate,"
    If typCustomer.intAge >= ADULT_AGE Then
        qrySelect = qrySelect & "CASE AdBasicDsa "
    Else
        qrySelect = qrySelect & "CASE BaBasicDsa "
    End If
    qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END RMR, TEE "
    qrySelect = qrySelect & "FROM Treat RIGHT JOIN BodyData AS b ON b.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "INNER JOIN CompData AS c ON b.CompDataNum=c.CompDataNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY Treat.TreatDay DESC;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With typObesity
            .sngWeight = Is_Null(rValue(0, 0), 0)
            .sngObesityRate = Is_Null(rValue(1, 0), 0)
            .sngWaist = Is_Null(rValue(2, 0), 0)
            .sngWHR = Is_Null(rValue(3, 0), 0)
            .sngBMI = Is_Null(rValue(4, 0), 0)
            .sngChFatRate = Is_Null(rValue(5, 0), 0)
            .sngRMR = Is_Null(rValue(6, 0), 0)
            .sngTEE = Is_Null(rValue(7, 0), 0)
        End With
        typTreatCal.sngTreatCal = typObesity.sngTEE
        '1) 현재
        For i = 0 To 7
            lblObInfo(i).Caption = Is_Null(rValue(i, 0), "-")
        Next i
        lblObInfo(0).Caption = lblObInfo(0).Caption & "kg"
        lblObInfo(1).Caption = CInt(lblObInfo(1).Caption) & "%"
        If lblObInfo(2).Caption <> "-" Then
            lblObInfo(2).Caption = lblObInfo(2).Caption & "cm"
        End If
        If lblObInfo(3).Caption <> "-" Then
            lblObInfo(3).Caption = Format(lblObInfo(3).Caption, "0.00")
        End If
        If lblObInfo(4).Caption <> "-" Then
            lblObInfo(4).Caption = Format(lblObInfo(4).Caption, "#.0")
        End If
        If lblObInfo(5).Caption <> "-" Then
            lblObInfo(5).Caption = Format(lblObInfo(5).Caption, "#.0") & "%"
        End If
        lblObInfo(6).Caption = Format(lblObInfo(6).Caption, "#,###")
        lblObInfo(7).Caption = Format(lblObInfo(7).Caption, "#,###")
        If UBound(rValue, 2) > 0 Then
            '2) 전회대비
            If Not IsNull(rValue(0, 1)) Then    '몸무게
                sngGab = typObesity.sngWeight - rValue(0, 1)
                Call DrawUpDown(sngGab, 0, "kg")
            Else
                Set imgUpDown(0).Picture = LoadPicture("")
                lblUpDown(0).Caption = "-"
            End If
            If Not IsNull(rValue(1, 1)) Then    '비만도
                sngGab = CInt(typObesity.sngObesityRate - rValue(1, 1))
                Call DrawUpDown(sngGab, 1, "%")
            Else
                Set imgUpDown(1).Picture = LoadPicture("")
                lblUpDown(1).Caption = "-"
            End If
            If Not IsNull(rValue(2, 1)) Then    '허리
                sngGab = typObesity.sngWaist - rValue(2, 1)
                Call DrawUpDown(sngGab, 2, "cm")
            Else
                Set imgUpDown(2).Picture = LoadPicture("")
                lblUpDown(2).Caption = "-"
            End If
            If Not IsNull(rValue(3, 1)) Then    'WHR
                sngGab = typObesity.sngWHR - rValue(3, 1)
                Call DrawUpDown(sngGab, 3, "", "0.00")
            Else
                Set imgUpDown(3).Picture = LoadPicture("")
                lblUpDown(3).Caption = "-"
            End If
            If Not IsNull(rValue(4, 1)) Then    'BMI
                sngGab = typObesity.sngBMI - rValue(4, 1)
                Call DrawUpDown(sngGab, 4, "")
                lblUpDown(4).Caption = Format(lblUpDown(4).Caption, "0.0")
            Else
                Set imgUpDown(4).Picture = LoadPicture("")
                lblUpDown(4).Caption = "-"
            End If
            If Not IsNull(rValue(5, 1)) Then    '체지방율
                sngGab = typObesity.sngChFatRate - rValue(5, 1)
                
                Call DrawUpDown(Format(sngGab, "0.0"), 5, "%")
            Else
                Set imgUpDown(5).Picture = LoadPicture("")
                lblUpDown(5).Caption = "-"
            End If
            If Not IsNull(rValue(6, 1)) Then    'RMR(휴식대사량)
                sngGab = typObesity.sngRMR - rValue(6, 1)
                Call DrawUpDown(sngGab, 6, "")
            Else
                Set imgUpDown(6).Picture = LoadPicture("")
                lblUpDown(6).Caption = "-"
            End If
            If Not IsNull(rValue(7, 1)) Then    'TEE (?? 뭐였더라?)
                sngGab = typObesity.sngTEE - rValue(7, 1)
                Call DrawUpDown(sngGab, 7, "")
            Else
                Set imgUpDown(7).Picture = LoadPicture("")
                lblUpDown(7).Caption = "-"
            End If
            '3) 최고/최저
            '체중
            lblMax(0).Caption = MaxValue("Weight") & "kg"
            lblMin(0).Caption = MinValue("Weight") & "kg"
            '비만도
            lblMax(1).Caption = CInt(MaxValue("ObesityRate")) & "%"
            lblMin(1).Caption = CInt(MinValue("ObesityRate")) & "%"
            '허리둘레
            lblMax(2).Caption = MaxValue("Waist") & "cm"
            lblMin(2).Caption = MinValue("Waist") & "cm"
            'WHR
            lblMax(3).Caption = Format(MaxValue("WHR"), "0.00")
            lblMin(3).Caption = Format(MinValue("WHR"), "0.00")
            'BMI
            lblMax(4).Caption = Format(MaxValue("BMI"), "0.0")
            lblMin(4).Caption = Format(MinValue("BMI"), "0.0")
            '체지방률
            lblMax(5).Caption = Format(MaxValue("ChFatRate"), "0.0") & "%"
            lblMin(5).Caption = Format(MinValue("ChFatRate"), "0.0") & "%"
            'RMR
            lblMax(6).Caption = Format(MaxValue("RMR"), "#,###")
            lblMin(6).Caption = Format(MinValue("RMR"), "#,###")
            'TEE
            lblMax(7).Caption = Format(MaxValue("TEE"), "#,###")
            lblMin(7).Caption = Format(MinValue("TEE"), "#,###")
        End If
    Else
        With typObesity
            .sngWeight = 0
            .sngObesityRate = 0
            .sngWaist = 0
            .sngWHR = 0
            .sngBMI = 0
            .sngChFatRate = 0
            .sngRMR = 0
            .sngTEE = 0
        End With
        typTreatCal.sngTreatCal = 0

        For i = 0 To 7
            lblObInfo(i).Caption = ""
            Set imgUpDown(i).Picture = LoadPicture("")
            lblUpDown(i).Caption = ""
            lblMax(i).Caption = ""
            lblMin(i).Caption = ""
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub DrawUpDown(sngGab As Single, i As Integer, strUnit As String, Optional strFormat As String)
    If sngGab < 0 Then       '하향 화살표 파란색 글씨
        Set imgUpDown(i).Picture = LoadPicture(App.Path & IMG_DOWN)
        If strFormat <> "" Then
            lblUpDown(i).Caption = Format(sngGab, strFormat) & strUnit
        Else
            lblUpDown(i).Caption = sngGab & strUnit
        End If
    ElseIf sngGab > 0 Then   '상향 화살표 빨간색 글씨
        Set imgUpDown(i).Picture = LoadPicture(App.Path & IMG_UP)
        If strFormat <> "" Then
            lblUpDown(i).Caption = Format(sngGab, strFormat) & strUnit
        Else
            lblUpDown(i).Caption = sngGab & strUnit
        End If
    Else                     '이전과 같음..변동없음
        Set imgUpDown(i).Picture = LoadPicture("")
        lblUpDown(i).Caption = "---"
    End If
End Sub

Private Sub imgAppend_Click(Index As Integer)
    frmPop_Additional.intCounselNo = Index
    Call frmPop_Additional.Show
End Sub

Private Sub imgDel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgDel.Picture = LoadPicture(App.Path & PATH01 & IMG_DEL_ON)
End Sub

Private Sub imgDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgDel.Picture = LoadPicture(App.Path & PATH01 & IMG_DEL_OFF)
    
    If glngTreatNum = 0 Then
        MsgBox "삭제할 진료내역을 선택하십시오.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    '신체계측기록이 한개뿐인 경우 경고멘트 날림
    If grdTable.Rows <= 3 Then
        If MsgBox("선택한 내역을 삭제하시면 자료가 하나도 남지 않습니다." & vbNewLine & "그래도 삭제하시겠습니까?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    Else
        If gintBottomButton = 0 Then
            If MsgBox(grdTable.TextMatrix(grdTable.Row, 1) & " 일자 진료내역을 삭제하시겠습니까?" & vbNewLine & vbNewLine & "* 주의 : 해당일의 검사항목 및 Notice 모든 데이터가 삭제됩니다 !", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        ElseIf gintBottomButton = 2 Then
            If MsgBox(grdTable.TextMatrix(grdTable.Row, 0) & " 일자 진료내역을 모두 삭제하시겠습니까?" & vbNewLine & vbNewLine & "* 주의 : 해당일의 검사항목 및 Notice 모든 데이터가 삭제됩니다 !", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    'Treat,BodyData,TreatData,TreatPrint 삭제
    If DeleteTreat(glngTreatNum) = True Then
        Call Form_Load
        MsgBox "삭제되었습니다.", vbOKOnly + vbInformation
    Else
        MsgBox "삭제에 실패했습니다.", vbOKOnly + vbCritical
    End If
End Sub

Private Sub imgDietCal_Click(Index As Integer)
    Dim i As Integer
    
    AdminYn 11
    If AccessYn = False Then
        MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
        Exit Sub
    End If
    
    For i = 0 To 5
        Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
    Next i
    If idxDietCal = Index Then
        With typTreatCal
            If .intDietCal = 0 Then
                Set imgDietCal(Index).Picture = LoadPicture(App.Path & IMG_ME & Index & "green.jpg")
                idxDietCal = Index
            Else
                .intDietCal = 0
                .sngTreatCal = typObesity.sngTEE - .intDietCal
                lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
                If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                    .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                    lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
                End If
                Exit Sub
            End If
        End With
    Else
        Set imgDietCal(Index).Picture = LoadPicture(App.Path & IMG_ME & Index & "green.jpg")
        idxDietCal = Index
    End If
    
    With typTreatCal
        Select Case Index
            Case 0: .intDietCal = 125
            Case 1: .intDietCal = 250
            Case 2: .intDietCal = 375
            Case 3: .intDietCal = 500
            Case 4: .intDietCal = 750
            Case 5: .intDietCal = 1000
            Case Else: .intDietCal = 0
        End Select
    
'마리 says:
'감량체중구하는식
'마리 says:
'(식사열량*7 + 운동열량*운동일수)/7700 * 4
    
        If .intDietCal >= 0 Then
            .sngTreatCal = typObesity.sngTEE - .intDietCal
            lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
        End If
    '+=========================================
    '+ 감량체중 구하는 식
    '+-----------------------------------------
        If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
        End If
    End With
End Sub

Private Sub imgExCal_Click(Index As Integer)
    Dim i As Integer
    
    AdminYn 11
    If AccessYn = False Then
        MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
        Exit Sub
    End If
    
    For i = 0 To 3
        Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
    Next i
    If idxExCal = Index Then
        With typTreatCal
            If .intExCal = 0 Then
                Set imgExCal(Index).Picture = LoadPicture(App.Path & IMG_EX & Index + 2 & "00-green.jpg")
                .intExCal = (Index + 2) * 100
                typExProgram.sngExCalory = (Index + 2) * 100
                intNowExCalory = .intExCal
                idxExCal = Index
            Else
                .intExCal = 0
                If chkUserCalory.Value = 0 Then
                    If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                        .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
                    End If
                Else
                    If .intUserCal >= 0 And .intExCal >= 0 And .intExCal >= 0 Then
                        .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
                    End If
                End If
                Exit Sub
            End If
        End With
    Else
        Set imgExCal(Index).Picture = LoadPicture(App.Path & IMG_EX & Index + 2 & "00-green.jpg")
        typTreatCal.intExCal = (Index + 2) * 100
        typExProgram.sngExCalory = (Index + 2) * 100
        intNowExCalory = typTreatCal.intExCal
        idxExCal = Index
    End If
    
    With typTreatCal
        .intExCal = (Index + 2) * 100
        intNowExCalory = .intExCal
        
        If chkUserCalory.Value = 0 Then
            If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
            End If
        Else
            If .intUserCal >= 0 And .intExCal >= 0 And .intExCal >= 0 Then
                .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
            End If
        End If
    End With
End Sub

Private Sub imgExDay_Click(Index As Integer)
    Dim i As Integer
    
    AdminYn 11
    If AccessYn = False Then
        MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
        Exit Sub
    End If
    
    For i = 0 To 3
        Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "일-gray.jpg")
    Next i
    If idxExDay = Index Then
        With typTreatCal
            If .intExDay = 0 Then
                Set imgExDay(Index).Picture = LoadPicture(App.Path & IMG_EX & Index + 3 & "일-green.jpg")
                .intExDay = Index + 3
                idxExDay = Index
            Else
                .intExDay = 0
                If chkUserCalory.Value = 0 Then
                    If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                        .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
                    End If
                Else
                    If .intUserCal >= 0 And .intExCal >= 0 And .intExCal >= 0 Then
                        .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
                    End If
                End If
                Exit Sub
            End If
        End With
    Else
        Set imgExDay(Index).Picture = LoadPicture(App.Path & IMG_EX & Index + 3 & "일-green.jpg")
        typTreatCal.intExDay = Index + 3
        idxExDay = Index
    End If
    
    With typTreatCal
        .intExDay = Index + 3
        
        If chkUserCalory.Value = 0 Then
            If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
            End If
        Else
            If .intUserCal >= 0 And .intExCal >= 0 And .intExCal >= 0 Then
                .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
            End If
        End If
    End With
End Sub

Private Sub imgGoReserve_Click()
    AdminYn 10
    If AccessYn = False Then
        MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
        Exit Sub
    End If
   frm_ReserveInsert.Show vbModal
    Call InitialSpread4
    Call InitialSpread5
'    Call InitialSpread42
'    Call InitialSpread52
'   초기화와 예약 내역을 보여줘야 함.
    
End Sub

Private Sub imgGoReserve_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgGoReserve.Picture = LoadPicture(App.Path & IMG_RES_OFF)
End Sub

Private Sub imgGoReserve_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgGoReserve.Picture = LoadPicture(App.Path & IMG_RES_ON)
End Sub

'+--------------------------------------------------
'+ 비만도평가
'+--------------------------------------------------
Private Sub Print_Obesity()
    Dim strfilename As String
    Dim qrySelect As String, rValue As Variant
            
On Error GoTo PrintErr
    qrySelect = "SELECT BodyDataNum FROM Treat RIGHT JOIN BodyData "
    qrySelect = qrySelect & "ON Treat.TreatNum=BodyData.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.TreatNum=" & glngTreatNum
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If IsNull(rValue) Then
        MsgBox "비만도평가를 출력할 내역이 없습니다. 다른 진료기록을 선택하십시오.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    Call Prepare_OBMent(typCustomer.strSex)
    If typCustomer.intAge >= ADULT_AGE Then
        strfilename = "\Report\비만도.rpt"
    Else
        If typCustomer.strSex = "M" Then
            strfilename = "\Report\비만도MB.rpt"
        ElseIf typCustomer.strSex = "F" Then
            strfilename = "\Report\비만도FB.rpt"
        End If
    End If
    Set crxReport = crxApplication.OpenReport(App.Path & strfilename)
    crxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    crxReport.PaperOrientation = crPortrait
    crxReport.RecordSelectionFormula = "{Treat.TreatNum}=" & glngTreatNum
        
    crxReport.Database.Tables(1).SetLogOnInfo strServer, strDBName, strUID, strPWD
    crxReport.PrintOut
   
    MsgBox "출력이 완료되었습니다.", vbOKOnly + vbInformation, "출력"
    Exit Sub
PrintErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "Print_Obesity", "frmCounsel_1", Err.Number, Err.Description
    MsgBox "출력에 실패했습니다." & vbNewLine & Err.Number & Err.Description, vbOKOnly + vbCritical
End Sub

'비만도평가의 멘트를 준비하는 함수
Private Sub Prepare_OBMent(strSex As String)
    Dim qrySelect As String
    Dim rValue As Variant
    Dim intSex As Integer

    If strSex = "M" Then
        intSex = 1
    Else
        intSex = 2
    End If
    Set clsSelect = New clsSelect
    
    '비만도 멘트를 준비한다.
    qrySelect = "SELECT ment FROM BodyData a, ObesityMent b "
    qrySelect = qrySelect & "WHERE a.ObesityRate >= b.ObesityRate1 "
    qrySelect = qrySelect & "AND a.ObesityRate < b.ObesityRate2 "
    qrySelect = qrySelect & "AND a.ChFatRate >= b.ChFatRate1 "
    qrySelect = qrySelect & "AND a.ChFatRate < b.ChFatRate2 "
    qrySelect = qrySelect & "AND ISNULL(a.WHR, 0) >= b.WHR1 "
    qrySelect = qrySelect & "AND ISNULL(a.WHR, 0) < b.WHR2 "
    qrySelect = qrySelect & "AND Sex=" & intSex

    qrySelect = qrySelect & " AND a.TreatNum=" & glngTreatNum

    rValue = clsSelect.Query(qrySelect)
        
    If Not IsNull(rValue) Then
        qrySelect = "DELETE FROM RPT_ObMent WHERE CustomerNum=" & glngCustomerNum
        Call modSql.AdoExcuteSql(qrySelect)
        
        qrySelect = "INSERT INTO RPT_ObMent (CustomerNum, Ment) "
        qrySelect = qrySelect & "VALUES(" & glngCustomerNum
        qrySelect = qrySelect & ",'" & rValue(0, 0) & "')"
        Call modSql.AdoExcuteSql(qrySelect)
        
        Erase rValue
    End If
End Sub

Private Sub imgModifyTreat_Click()
    '만약 frmBottom.dtpUserDay.Value에 해당하는 처방이 없다면
    '새로 입력한다.
    '만약 frmBottom.dtpUserDay.Value에 해당하는 처방이 있다면
    '기존 데이터를 수정한다.
    '기존 데이터를 수정할때 해당 일자에 여러번의 처방이 있을수 있으므로
    '해당 날짜의 처방번호를 가지고 데이터를 수정한다.
    
'    If cmdInputTreat.Caption = "처방입력" Then
'        lblDispDate.Caption = Format(gdatUserDay, "YYYY-MM-DD")
'        intMode = 1
'        Call EnabledInput(True)
'    Else
'        intMode = 2
'        Call EnabledInput(True)
'    End If
    imgModifyTreat.Visible = False
    Dim qrySelect As String
'    qrySelect = "SELECT TOP 1 * FROM Treat a LEFT JOIN "
'
'    Call EnabledInput(Not cmbProgram.Enabled)


Dim lngcTreatNum
With grdTable
    If .TextMatrix(.Row, .Cols - 1) = "" Or IsNumeric(.TextMatrix(.Row, .Cols - 1)) = False Then
        Exit Sub
    End If
    
    lngcTreatNum = Format(.TextMatrix(.Row, .Cols - 1), "0")
    
    If lngcTreatNum > 0 Then
        mlngcTreatNum = lngcTreatNum
        intMode = 0
        Call EnabledInput(True)
        lblDispDate.Caption = Format(grdTable.TextMatrix(grdTable.Row, IIf(gintBottomButton = 0, 1, 0)), "YYYY-MM-DD")
        lblDispContent.Caption = "수정중"
        imgSave.Enabled = True
    Else
        MsgBox "수정할 처방을 입력하세요", vbInformation, "처방수정"
    End If
End With

End Sub

Private Sub imgModifyTreat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgModifyTreat.Picture = LoadPicture(App.Path & PATH01 & IMG_MODIFYTREAT_OFF)
End Sub

Private Sub imgModifyTreat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgModifyTreat.Picture = LoadPicture(App.Path & PATH01 & IMG_MODIFYTREAT_ON)
End Sub

Private Sub imgNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgNew.Picture = LoadPicture(App.Path & PATH01 & IMG_NEW_OFF)
End Sub

Private Sub imgNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgNew.Picture = LoadPicture(App.Path & PATH01 & IMG_NEW_ON)
    mlngcTreatNum = 0
    intMode = 3
    EnabledInput (True)
    imgSave.Enabled = True
    lblDispDate.Caption = Format(gdatUserDay, "YYYY-MM-DD")
    grdTable.ColSel = 0
    imgPreTreat.Visible = True
    imgModifyTreat.Visible = False
End Sub

Private Sub imgPreTreat_Click()
'가장 최근 처방한 식이열량, 운동열량, 운동일수 가져와서 체크
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, Index As Integer
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT TOP 1 TreatNum, DietCalory, ExDay, ExCalory, TreatCalory, UserCalory "
    qrySelect = qrySelect & "FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND NOT TreatCalory IS NULL"
    qrySelect = qrySelect & " AND (NOT DietCalory IS NULL OR NOT ExDay IS NULL OR NOT ExCalory IS NULL "
    qrySelect = qrySelect & " OR NOT UserCalory IS NULL)"
    qrySelect = qrySelect & " ORDER BY TreatDay DESC;"
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With typTreatCal
            .intDietCal = Is_Null(rValue(1, 0), 0)
            .intExDay = Is_Null(rValue(2, 0), 0)
            .intExCal = Is_Null(rValue(3, 0), 0)
            .sngTreatCal = Is_Null(rValue(4, 0), 0)
            lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
            '다이어트 열량
            If Not IsNull(rValue(5, 0)) Then
                chkUserCalory.Value = 1
                .intUserCal = rValue(5, 0)
                txtUserCalory.Text = .intUserCal
                For i = 0 To 5
                    Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
                    imgDietCal(i).Enabled = False
                Next i
            Else
                chkUserCalory.Value = 0
                .intUserCal = 0
                txtUserCalory.Text = ""
                For i = 0 To 5
                    Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
                    imgDietCal(i).Enabled = True
                Next i
                Select Case .intDietCal
                    Case 125: Index = 0
                    Case 250: Index = 1
                    Case 375: Index = 2
                    Case 500: Index = 3
                    Case 750: Index = 4
                    Case 1000: Index = 5
                    Case Else: Index = 6
                End Select
                If Index <> 6 Then
                    Set imgDietCal(Index).Picture = LoadPicture(App.Path & IMG_ME & Index & "green.jpg")
                End If
            End If
            '운동일수
            For i = 0 To 3
                Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "일-gray.jpg")
            Next i
            If .intExDay >= 3 And .intExDay <= 6 Then
                Set imgExDay(.intExDay - 3).Picture = LoadPicture(App.Path & IMG_EX & .intExDay & "일-green.jpg")
            End If
            '운동칼로리
            For i = 0 To 3
                Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
            Next i
            If .intExCal >= 200 And .intExCal <= 500 Then
                Set imgExCal((.intExCal / 100) - 2).Picture = LoadPicture(App.Path & IMG_EX & (.intExCal / 100) & "00-green.jpg")
            End If
            '감량체중
            ' 감량체중은 현재 체중을 고려해서 다시 계산한다.
'            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
            If .intUserCal = 0 Then
                .sngTreatCal = typObesity.sngTEE - .intDietCal
                If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                    .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                    lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
                End If
            Else
                .sngTreatCal = typObesity.sngTEE - .intUserCal
                If .intUserCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                    .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                    lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
                End If
            End If
        End With
    Else
        MsgBox "이전에 처방한 내역이 없습니다.", vbOKOnly + vbExclamation
    End If
    Set clsSelect = Nothing
End Sub

Private Sub imgPreTreat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPreTreat.Picture = LoadPicture(App.Path & PATH01 & IMG_PRETREAT_OFF)
End Sub

Private Sub imgPreTreat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPreTreat.Picture = LoadPicture(App.Path & PATH01 & IMG_PRETREAT_ON)
End Sub

Private Sub imgPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH01 & IMG_PRINT_ON)
End Sub

Private Sub imgPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH01 & IMG_PRINT_OFF)

    If glngTreatNum = 0 Then
        MsgBox "출력할 진료 기록을 선택하세요.", vbOKOnly + vbExclamation
        Exit Sub
    End If

    strServer = ServerName
'2005-01-18 류진선 DB정보수정
    strDBName = DBinfo.DBName
    strUID = DBinfo.DBID
    strPWD = DBinfo.DBPWD
'    strDBName = "Body"
'    strUID = "sa"
'    strPWD = "1111"

''미리보기
'    CrystalReport1.Destination = crptToWindow
''프린트출력
'    CrystalReport1.Destination = crptToPrinter
    Call Print_Obesity
End Sub

Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgSave.Picture = LoadPicture(App.Path & PATH01 & IMG_SAVE_ON)
End Sub

Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' * 저장 : 신체계측 or 신체계측+처치내용 or 처치내용
'입력되어 있으면 저장해야하나?

    Dim qryUpdate As String
    Dim lRet As Long
On Error GoTo SaveErr
'트랜잭션 걸어야겠지
    Set imgSave.Picture = LoadPicture(App.Path & PATH01 & IMG_SAVE_OFF)
    If glngCustomerNum = 0 Then
        MsgBox "저장할 환자를 먼저 선택하세요", vbInformation
        Exit Sub
    End If
    If typTreatCal.sngTreatCal = 0 Then
        MsgBox "칼로리처방을 하신후 다시 저장하십시오.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    If intMode = 0 Then '조회하다가 수정버튼을 클릭한경우 mlngcTreatNum에 있는 데이터를 수정한다.
        If mlngcTreatNum = 0 Then
            MsgBox "수정할 진료내역을 선택하십시오.", vbOKOnly + vbInformation
            Exit Sub
        Else
            lRet = UpdateTreat(mlngcTreatNum)
            If lRet = 1 Then
                If SaveTreatData(mlngcTreatNum) = True Then
                    If SaveTreatPrint(mlngcTreatNum) = True Then
                        '치료프로그램 저장
                        qryUpdate = "UPDATE CustomerInfo SET ChPgCode='" & typCustomer.strChPgCode
                        qryUpdate = qryUpdate & "' WHERE CustomerNum=" & glngCustomerNum
                        modSql.AdoExcuteSql (qryUpdate)
            
                        Call imgViewHistory_Click(2)
                        MsgBox "저장되었습니다.", vbOKOnly + vbInformation
                    Else
                        MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                        Exit Sub
                    End If
                Else
                    MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                    Exit Sub
                End If
            ElseIf lRet = 0 Then
                MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                Exit Sub
            End If
        End If
        
    ElseIf intMode = 1 Then    '신규입력모드
                                '최초로드시 신체Data와 처방칼로리 Data가 없어서 자동으로 입력이 되는 경우
                                '새 glngTreatNum을 따서 만듦.
        Dim oldTreatNum  As Long
        oldTreatNum = glngTreatNum
        glngTreatNum = SaveTreat
        If glngTreatNum = 0 Then
            MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
            Exit Sub
        ElseIf glngTreatNum = -1 Then
            glngTreatNum = oldTreatNum
        Else
            If SaveTreatData = True Then
                If SaveTreatPrint = True Then
                    '하단의 표와 그래프 업데이트 한 후 처치내역으로 보여준다.
            
                    '치료프로그램 저장
                    qryUpdate = "UPDATE CustomerInfo SET ChPgCode='" & typCustomer.strChPgCode
                    qryUpdate = qryUpdate & "' WHERE CustomerNum=" & glngCustomerNum
                    modSql.AdoExcuteSql (qryUpdate)
            
                    '하단의 표와 그래프 업데이트
                    Call imgViewHistory_Click(2)
                    MsgBox "저장되었습니다.", vbOKOnly + vbInformation
                Else
                    MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                Exit Sub
            End If
        End If
    ElseIf intMode = 2 Then    '수정모드
                                '최초로드시 다양한 이유로 자동으로 수정이 되는 경우
                                'glngTreatNum에 해당하는 데이터를 수정함.
        If glngTreatNum = 0 Then
            MsgBox "수정할 진료내역을 선택하십시오.", vbOKOnly + vbInformation
            Exit Sub
        Else
            lRet = UpdateTreat(glngTreatNum)
            If lRet = 1 Then
                If SaveTreatData = True Then
                    If SaveTreatPrint = True Then
                        '치료프로그램 저장
                        qryUpdate = "UPDATE CustomerInfo SET ChPgCode='" & typCustomer.strChPgCode
                        qryUpdate = qryUpdate & "' WHERE CustomerNum=" & glngCustomerNum
                        modSql.AdoExcuteSql (qryUpdate)
            
                        Call imgViewHistory_Click(2)
                        MsgBox "저장되었습니다.", vbOKOnly + vbInformation
                    Else
                        MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                        Exit Sub
                    End If
                Else
                    MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                    Exit Sub
                End If
            ElseIf lRet = 0 Then
                MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                Exit Sub
            End If
        End If
    ElseIf intMode = 3 Then ' 무조건 새로 삽입하고 싶어함. 즉, New버튼을 클릭했을 경우
        mlngcTreatNum = SaveTreat(Trim(lblDispDate))
        If mlngcTreatNum = 0 Then
            MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
            Exit Sub
        Else
            If SaveTreatData(mlngcTreatNum) = True Then
                If SaveTreatPrint(mlngcTreatNum) = True Then
                    '하단의 표와 그래프 업데이트 한 후 처치내역으로 보여준다.
            
                    '치료프로그램 저장
                    qryUpdate = "UPDATE CustomerInfo SET ChPgCode='" & typCustomer.strChPgCode
                    qryUpdate = qryUpdate & "' WHERE CustomerNum=" & glngCustomerNum
                    modSql.AdoExcuteSql (qryUpdate)
            
                    '하단의 표와 그래프 업데이트
                    Call imgViewHistory_Click(2)
                    MsgBox "저장되었습니다.", vbOKOnly + vbInformation
                Else
                    MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "저장에 실패했습니다.", vbOKOnly + vbCritical
                Exit Sub
            End If
        End If
        
    End If

    Exit Sub
SaveErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "imgSave_MouseUp", "frmCounsel_1", Err.Number, Err.Description
    MsgBox "저장에 실패했습니다." & vbNewLine & vbNewLine & Err.Description, vbOKOnly + vbCritical
End Sub

Private Sub imgSelectEx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgSelectEx.Picture = LoadPicture(App.Path & IMG_SELEX_OFF)
End Sub

Private Sub imgSelectEx_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intProgram As Integer
    
    Call LoadExProgram(glngCustomerNum)
    intProgram = WhatExProgram(typExProgram.intAge, typExProgram.strSex, typExProgram.intObesity, typExProgram.intComplication, typExProgram.strBodyStatus)
    
    Set imgSelectEx.Picture = LoadPicture(App.Path & IMG_SELEX_ON)
    If WhatExConfig = True Then
        If intProgram = 9 Or intProgram = 23 Or intProgram = 37 Then
            MsgBox "운동종목을 선택하실 수 없습니다.", vbOKOnly + vbExclamation
            Exit Sub
        Else
            frmPop_ExAerobic.Show vbModal
        End If
    Else
        frmPop_ExAerobic2.Show vbModal
    End If
End Sub

Private Sub imgTreat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgTreat.Picture = LoadPicture(App.Path & IMG_TREAT_OFF)
End Sub

Private Sub imgTreat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgTreat.Picture = LoadPicture(App.Path & IMG_TREAT_ON)
    '치료이력 창 띄우기
    frm_TreatDetail.Show vbModal
End Sub

'변화도(표),변화도(그래프),처방내역조회 탭을 클릭함.
Private Sub imgViewHistory_Click(Index As Integer)
    Dim i As Integer
On Error GoTo Err
    gintBottomButton = Index
    Select Case Index
        Case 0  '변화도(표)
            Set imgViewHistory(0).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB1_ON)
            Set imgViewHistory(1).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB2_OFF)
            Set imgViewHistory(2).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB3_OFF)
            
            Call Bottom0(0)
            imgTreat.Visible = False
        Case 1  '변화도(그래프)
            Set imgViewHistory(0).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB1_OFF)
            Set imgViewHistory(1).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB2_ON)
            Set imgViewHistory(2).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB3_OFF)
            
            Call Bottom1(0)
            imgTreat.Visible = False
        Case 2  '처방내역조회
            Set imgViewHistory(0).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB1_OFF)
            Set imgViewHistory(1).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB2_OFF)
            Set imgViewHistory(2).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB3_ON)
            
            Call Bottom2(0)
            imgTreat.Visible = True
    End Select
    Exit Sub
Err:
    '2004-12-23 류진선 로그기록
    'WriteLog "imgViewHistory_Click", "frmCounsel_1", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

Private Sub imgTab_Click(Index As Integer)
    Select Case Index
        Case 0: '치료종목/출력물 탭
            Set imgTab(0).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB1_ON)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB2_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB3_OFF)
            
            sprTreat.Visible = True
            sprPrint.Visible = True
            txtNotice.Visible = False
            txtMemo.Visible = False
        Case 1: 'Notice 탭
            AdminYn 14
            If AccessYn = False Then
                MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
                Exit Sub
            End If

            Set imgTab(0).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB1_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB2_ON)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB3_OFF)
            
            txtNotice.Visible = True
            sprTreat.Visible = False
            sprPrint.Visible = False
            txtMemo.Visible = False
            
            If txtNotice.Enabled Then txtNotice.SetFocus
        Case 2: 'Memo탭
            AdminYn 13
            If AccessYn = False Then
                MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
                Exit Sub
            End If
            
            Set imgTab(0).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB1_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB2_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB3_ON)
            
            txtMemo.Visible = True
            sprTreat.Visible = False
            sprPrint.Visible = False
            txtNotice.Visible = False
            If txtMemo.Enabled Then txtMemo.SetFocus
    End Select
End Sub

Private Sub imgTab_DblClick(Index As Integer)
    If Index = 2 Then
        AdminYn 13
        If AccessYn = False Then
            MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
            Exit Sub
        End If
        frmCounsel_1Pop.Show vbModal
    End If
End Sub

Private Sub sprTreat_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Col = 3 Then
        AdminYn 12
        If AccessYn = False Then
            MsgBox "접근권한이 없습니다. 관리자에게 문의하십시요.", vbOKOnly + vbInformation, "접근권한 없슴"
            Exit Sub
        End If
    End If
End Sub

Private Sub txtUserCalory_Change()
    Dim sngMinus As Single
    
    If txtUserCalory.Text = "" Then
        'lblTreatCal.Caption = Format(msngChMealCalory, "#,###") & " kcal"
        If chkUserCalory.Value = vbChecked Then
            lblTreatCal.Caption = "0 kcal"
        End If
        Exit Sub
    End If
    If IsNumeric(txtUserCalory.Text) = False Then
        Exit Sub
    End If

    If CInt(txtUserCalory.Text) > 3500 Then
        
        MsgBox "식단열량은 " & CStr(1000 + msngChMealCalory) & "kcal이상 3500kcal이하로 입력하십시오.", vbOKOnly + vbExclamation
        txtUserCalory.SelStart = 0
        txtUserCalory.SelLength = Len(txtUserCalory)
        txtUserCalory.SetFocus
        Exit Sub
    End If
    
    With typTreatCal

        .intUserCal = CInt(txtUserCalory.Text)
        If .intUserCal >= 0 Then

            .sngTreatCal = .intUserCal
            lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
        End If
    '+=========================================
    '+ 감량체중 구하는 식
    '+-----------------------------------------
        sngMinus = typObesity.sngTEE - .intUserCal
        If sngMinus >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((sngMinus * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
'        Else
'            MsgBox "..."
        End If
    End With
End Sub

Private Sub txtUserCalory_LostFocus()
    Dim sngMinus As Single
    
    If txtUserCalory.Text = "" Then
        Exit Sub
    End If
    If IsNumeric(txtUserCalory.Text) = False Then
        Exit Sub
    End If
    With typTreatCal
        .intUserCal = CInt(txtUserCalory.Text)
        If .intUserCal >= 0 Then
            .sngTreatCal = .intUserCal
            lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
        End If
        '+=========================================
        '+ 감량체중 구하는 식
        '+-----------------------------------------
        sngMinus = typObesity.sngTEE - .intUserCal
        If sngMinus >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((sngMinus * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/월"
        End If
    End With
End Sub

'************ 2Fon 관련 추가개발(평가등록 팝업)
Private Sub imgValuation_Click()
    frmPop_Valuation.datValuation = Format(lblDispDate, "yyyy-mm-dd")
    frmPop_Valuation.Show vbModal
End Sub

Private Sub imgValuation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgValuation.Picture = LoadPicture(App.Path & IMG_평가_ON)
End Sub

Private Sub imgValuation_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgValuation.Picture = LoadPicture(App.Path & IMG_평가_OFF)
End Sub




'================================================================================================
'================================================================================================
'================================================================================================
'================================================================================================
'2005-01-27 류진선 수정
Private Sub InitialControl2()
    Dim rValue As Variant
    Dim i As Integer

    Set clsSelect = New clsSelect

    '비만관련 항목 초기화(왼쪽상단)
    For i = 0 To 7
        lblObInfo(i).Caption = ""                   '현재
        Set imgUpDown(i).Picture = LoadPicture("")  '전회대비 이미지
        lblUpDown(i).Caption = ""                   '전회대비 값
        lblMax(i).Caption = ""                      '최대치
        lblMin(i).Caption = ""                      '최소치
    Next i
    '치료내역리스트(왼쪽하단)
    Set imgViewHistory(0).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB1_OFF)
    Set imgViewHistory(1).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB2_OFF)
    Set imgViewHistory(2).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB3_ON)
    imgTreat.Visible = True
    
    '치료내역리스트를 보여줌.
    gintBottomButton = 2
    grdTable.Visible = True
    Chart.Visible = False
    cmdSub(0).Visible = True
    cmdSub(1).Visible = True
    For i = 2 To 9
        cmdSub(i).Visible = False
    Next i
    cmdSub(0).Caption = "처      치"
    cmdSub(1).Caption = "출  력  물"

    '치료프로그램 초기화(오른쪽상단)
    rValue = clsSelect.Query("SELECT ChPgName, ChPgCode FROM ChPgUpdate;")
    If Not IsNull(rValue) Then
        cmbProgram.Clear
        For i = 0 To UBound(rValue, 2)
            cmbProgram.AddItem Trim(rValue(0, i))
            If Trim(rValue(1, i)) = Trim(typCustomer.strChPgCode) Then
                cmbProgram.ListIndex = i
            End If
        Next i
    End If
    Call InitialTreatCalory2
    
    
    '식단열량 직접입력하기(오른쪽상단일부)
    txtUserCalory.Text = ""
    txtUserCalory.BackColor = FRM_GRAY
    txtUserCalory.Enabled = False
    chkUserCalory.Value = vbUnchecked
    
    '대용식 처방
    chkChMeal.Value = vbUnchecked
    cboChMealTime.Visible = False
    cboChMealTime.Clear
    cboChMealTime.AddItem "대용식 - 아침": cboChMealTime.ItemData(0) = 1
    cboChMealTime.AddItem "대용식 - 점심": cboChMealTime.ItemData(1) = 2
    cboChMealTime.AddItem "대용식 - 저녁": cboChMealTime.ItemData(2) = 3
    cboChMealTime.AddItem "대용식 - 아침+간식": cboChMealTime.ItemData(3) = 4
    cboChMealTime.AddItem "대용식 - 점심+간식": cboChMealTime.ItemData(4) = 5
    cboChMealTime.AddItem "대용식 - 저녁+간식": cboChMealTime.ItemData(5) = 6
    cboChMealTime.AddItem "==================": cboChMealTime.ItemData(6) = 0
    cboChMealTime.AddItem "사용자 식단": cboChMealTime.ItemData(7) = 10
    cboChMealTime.ListIndex = 0
    cboChMealCode.Visible = False
    
    
    '치료종목/출력물/Memo/Notice(오른쪽하단)
    Call InitialSpread42 '치료종목
    Call InitialSpread52 '출력물
    txtNotice.Text = ""
    txtMemo.Text = ""
    txtNotice.Visible = False
    txtMemo.Visible = False
    


    Set clsSelect = Nothing
End Sub

'2005-01-27 류진선 수정
Private Sub InitialSpread42()
'처방입력하는 스프레드 초기화(치료종목/출력물 화면의 치료종목)
'모든 치료종목을 불러옴
    Dim qrySelect As String
    Dim rValue As Variant, i As Integer, j As Integer

    With sprTreat
        .EditEnterAction = EditEnterActionDown
        .EditModePermanent = True
        
        .GrayAreaBackColor = vbWhite
        .BackColor = &HCEF7E7
        .GridColor = vbWhite
        .Font.Size = 7
        .RowHeight(-1) = 13
        .MaxCols = 4
        .ColWidth(1) = 9
        .ColWidth(2) = 2
        .ColWidth(3) = 2
        .ColWidth(4) = 0
        .Col = 2: .Row = -1
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .DisplayColHeaders = False
        .DisplayRowHeaders = False
        .ScrollBars = ScrollBarsNone

        Set clsSelect = New clsSelect
        rValue = clsSelect.Query("SELECT TreatName, TreatCode FROM TreatCode;")
        If Not IsNull(rValue) Then
            .MaxRows = UBound(rValue, 2) + 1
            If .MaxRows > 6 Then
                .ScrollBars = ScrollBarsVertical
                .ColWidth(1) = 8.5
                .Width = 1800
            Else
                .ScrollBars = ScrollBarsNone
                .ColWidth(1) = 9
                .Width = 1650
            End If
            .Col = 1
            For i = 0 To UBound(rValue, 2)
                .Row = i + 1
                .Col = 1: .Text = Trim(rValue(0, i)): .Lock = True
                .Col = 2: .CellType = CellTypeCheckBox: .Value = False
                .TypeCheckCenter = True
                .TypeCheckType = TypeCheckTypeNormal
                .Col = 3: .CellType = CellTypeCheckBox: .Value = False
                .Col = 4: .Text = Trim(rValue(1, i))
            Next i
        Else
            .MaxRows = 1
            .Row = 1: .Col = 1: .Text = ""
            .Col = 2: .CellType = CellTypeCheckBox: .Value = False
            .Col = 3: .CellType = CellTypeCheckBox: .Value = False
            .Col = 4: .Text = ""
            Exit Sub
        End If

'        '예약되어 있는지 확인하고 있을시에는 해당 치료명은 체크한다
'        '셀의 색을 달리한다.
'        qrySelect = "SELECT ReserveDetail.ReDetailNum, ReserveTreat.TreatCode "
'        qrySelect = qrySelect & "FROM ReserveDetail INNER JOIN "
'        qrySelect = qrySelect & "Reserve ON ReserveDetail.ReserveNum = Reserve.ReserveNum "
'        qrySelect = qrySelect & "INNER JOIN ReserveTreat ON "
'        qrySelect = qrySelect & "ReserveDetail.ReDetailNum = ReserveTreat.ReDetailNum "
'        qrySelect = qrySelect & "WHERE ReserveDetail.ReserveDay='" & Format(gdatUserDay, "YYYYMMDD") & "' "
'        qrySelect = qrySelect & "AND Reserve.CustomerNum = " & glngCustomerNum
'
'        rValue = clsSelect.Query(qrySelect)
'        If Not IsNull(rValue) Then
'            For i = 0 To UBound(rValue, 2)
'                For j = 1 To .MaxRows
'                    .Row = j: .Col = 1
'                    If Trim(.Text) = Trim(rValue(1, i)) Then
'                        .Col = 2: .CellType = CellTypeCheckBox: .Value = True
'                        Exit For
'                    End If
'                Next j
'            Next i
'        End If
        Set clsSelect = Nothing
        .Lock = True
    End With
End Sub


'2005-01-27 류진선 수정
Private Sub InitialSpread52()
'처방입력하는 스프레드 초기화(치료종목/출력물 화면의 출력물)
'모든 출력물을 불러옴.
    Dim qrySelect As String
    Dim rValue As Variant, i As Integer, j As Integer

    With sprPrint
        .EditEnterAction = EditEnterActionDown
        .EditModePermanent = True

        .GrayAreaBackColor = vbWhite
        .BackColor = &HCEF7E7
        .GridColor = vbWhite
        .Font.Size = 7
        .RowHeight(-1) = 13
        .MaxCols = 3
        .ColWidth(1) = 10
        .ColWidth(2) = 2
        .ColWidth(3) = 0
        .DisplayColHeaders = False
        .DisplayRowHeaders = False
        .ScrollBars = ScrollBarsNone

        Set clsSelect = New clsSelect
        rValue = clsSelect.Query("SELECT PrintoutName, PrintoutNum FROM PrintOut;")
        If Not IsNull(rValue) Then
            .MaxRows = UBound(rValue, 2) + 1
            If .MaxRows > 6 Then
                .ScrollBars = ScrollBarsVertical
                .ColWidth(1) = 9.5
                .Width = 1650
            Else
                .ScrollBars = ScrollBarsNone
                .ColWidth(1) = 10
                .Width = 1600
            End If
            .Col = 1
            For i = 0 To UBound(rValue, 2)
                .Row = i + 1
                .Col = 1: .Text = Trim(rValue(0, i)): .Lock = True
                .Col = 2: .CellType = CellTypeCheckBox: .Value = False
                .Col = 3: .Text = Trim(rValue(1, i))
            Next i
        Else
            .MaxRows = 1
            .Row = 1: .Col = 1: .Text = ""
            .Col = 2: .CellType = CellTypeCheckBox: .Value = False
            .Col = 3: .Text = ""
            Exit Sub
        End If

'        '예약되어 있는 출력물에는 일단 체크하기
'        '예약되어 있는지 확인하고 있을시에는 해당 출력물명은 체크한다
'        qrySelect = "SELECT ReserveDetail.ReDetailNum, ReservePrint.PrintCode "
'        qrySelect = qrySelect & "FROM ReserveDetail INNER JOIN "
'        qrySelect = qrySelect & "Reserve ON ReserveDetail.ReserveNum=Reserve.ReserveNum "
'        qrySelect = qrySelect & "INNER JOIN ReservePrint ON ReserveDetail.ReDetailNum=ReservePrint.ReDetailNum "
'        qrySelect = qrySelect & "WHERE ReserveDetail.ReserveDay='" & Format(gdatUserDay, "YYYYMMDD") & "' "
'        qrySelect = qrySelect & "AND Reserve.CustomerNum=" & glngCustomerNum
'
'        rValue = clsSelect.Query(qrySelect)
'        If Not IsNull(rValue) Then
'            For i = 0 To UBound(rValue, 2)
'                For j = 1 To .MaxRows
'                    .Row = j: .Col = 1
'                    If Trim(.Text) = Trim(rValue(1, i)) Then
'                        .Col = 2: .CellType = CellTypeCheckBox: .Value = True
'                        Exit For
'                    End If
'                Next j
'            Next i
'        End If
        Set clsSelect = Nothing
    End With
End Sub

'2005-01-27 류진선 수정
'칼로리처방 데이터 초기화
Private Sub InitialTreatCalory2()
    Dim i As Integer
    Dim intExTime As Integer
'200 / 300 / 400 / 500
'빨리걷기 : 0.093 / 근력운동 : 0.105
    typTreatCal.sngTreatCal = 0
    typTreatCal.intDietCal = 0
    typTreatCal.intExCal = 0
    typTreatCal.intExDay = 0
    typTreatCal.intUserCal = 0
    typTreatCal.sngLossWeight = 0
    For i = 0 To 3
        lblAerobic(i).Caption = ""
        lblAnaerobic(i).Caption = ""
    Next i
    '유산소운동과 근력운동 보여주기
    For i = 0 To 3
        lblAerobic(i).Caption = "0 분"
        lblAnaerobic(i).Caption = "0 분"
    Next i
    
    For i = 0 To 5
        Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
    Next i
    For i = 0 To 3
        Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "일-gray.jpg")
    Next i
    For i = 0 To 3
        Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
    Next i
    lblLossWeight.Caption = ""
    lblTreatCal.Caption = Format(typTreatCal.sngTreatCal, "#,###") & "kcal"
End Sub

'2005-01-27 류진선 수정
'비만관련 데이터 보여주기
'체중/비만도/허리둘레/WHR/BMI/체지방률/RMR/TEE
' 휴식대사량 8 : inRMR / 9 : etcRMR / 그밖에 : RMR
Private Sub ShowObesityRecord2()
    Dim qrySelect As String, rValue As Variant
    Dim sngGab As Single
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT TOP 2 Weight, ObesityRate, Waist, b.WHR, BMI, ChFatRate,"
    If typCustomer.intAge >= ADULT_AGE Then
        qrySelect = qrySelect & "CASE AdBasicDsa "
    Else
        qrySelect = qrySelect & "CASE BaBasicDsa "
    End If
    qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END RMR, TEE "
    qrySelect = qrySelect & "FROM Treat RIGHT JOIN BodyData AS b ON b.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "INNER JOIN CompData AS c ON b.CompDataNum=c.CompDataNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY Treat.TreatDay DESC;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With typObesity
            .sngWeight = Is_Null(rValue(0, 0), 0)
            .sngObesityRate = Is_Null(rValue(1, 0), 0)
            .sngWaist = Is_Null(rValue(2, 0), 0)
            .sngWHR = Is_Null(rValue(3, 0), 0)
            .sngBMI = Is_Null(rValue(4, 0), 0)
            .sngChFatRate = Is_Null(rValue(5, 0), 0)
            .sngRMR = Is_Null(rValue(6, 0), 0)
            .sngTEE = Is_Null(rValue(7, 0), 0)
        End With
        typTreatCal.sngTreatCal = typObesity.sngTEE
        '1) 현재
        For i = 0 To 7
            lblObInfo(i).Caption = Is_Null(rValue(i, 0), "-")
        Next i
        lblObInfo(0).Caption = lblObInfo(0).Caption & "kg"
        lblObInfo(1).Caption = CInt(lblObInfo(1).Caption) & "%"
        If lblObInfo(2).Caption <> "-" Then
            lblObInfo(2).Caption = lblObInfo(2).Caption & "cm"
        End If
        If lblObInfo(3).Caption <> "-" Then
            lblObInfo(3).Caption = Format(lblObInfo(3).Caption, "0.00")
        End If
        If lblObInfo(4).Caption <> "-" Then
            lblObInfo(4).Caption = Format(lblObInfo(4).Caption, "#.0")
        End If
        If lblObInfo(5).Caption <> "-" Then
            lblObInfo(5).Caption = Format(lblObInfo(5).Caption, "#.0") & "%"
        End If
        lblObInfo(6).Caption = Format(lblObInfo(6).Caption, "#,###")
        lblObInfo(7).Caption = Format(lblObInfo(7).Caption, "#,###")
        If UBound(rValue, 2) > 0 Then
            '2) 전회대비
            If Not IsNull(rValue(0, 1)) Then    '몸무게
                sngGab = typObesity.sngWeight - rValue(0, 1)
                Call DrawUpDown(sngGab, 0, "kg")
            Else
                Set imgUpDown(0).Picture = LoadPicture("")
                lblUpDown(0).Caption = "-"
            End If
            If Not IsNull(rValue(1, 1)) Then    '비만도
                sngGab = CInt(typObesity.sngObesityRate - rValue(1, 1))
                Call DrawUpDown(sngGab, 1, "%")
            Else
                Set imgUpDown(1).Picture = LoadPicture("")
                lblUpDown(1).Caption = "-"
            End If
            If Not IsNull(rValue(2, 1)) Then    '허리
                sngGab = typObesity.sngWaist - rValue(2, 1)
                Call DrawUpDown(sngGab, 2, "cm")
            Else
                Set imgUpDown(2).Picture = LoadPicture("")
                lblUpDown(2).Caption = "-"
            End If
            If Not IsNull(rValue(3, 1)) Then    'WHR
                sngGab = typObesity.sngWHR - rValue(3, 1)
                Call DrawUpDown(sngGab, 3, "", "0.00")
            Else
                Set imgUpDown(3).Picture = LoadPicture("")
                lblUpDown(3).Caption = "-"
            End If
            If Not IsNull(rValue(4, 1)) Then    'BMI
                sngGab = typObesity.sngBMI - rValue(4, 1)
                Call DrawUpDown(sngGab, 4, "")
                lblUpDown(4).Caption = Format(lblUpDown(4).Caption, "0.0")
            Else
                Set imgUpDown(4).Picture = LoadPicture("")
                lblUpDown(4).Caption = "-"
            End If
            If Not IsNull(rValue(5, 1)) Then    '체지방율
                sngGab = typObesity.sngChFatRate - rValue(5, 1)
                
                Call DrawUpDown(Format(sngGab, "0.0"), 5, "%")
            Else
                Set imgUpDown(5).Picture = LoadPicture("")
                lblUpDown(5).Caption = "-"
            End If
            If Not IsNull(rValue(6, 1)) Then    'RMR(휴식대사량)
                sngGab = typObesity.sngRMR - rValue(6, 1)
                Call DrawUpDown(sngGab, 6, "")
            Else
                Set imgUpDown(6).Picture = LoadPicture("")
                lblUpDown(6).Caption = "-"
            End If
            If Not IsNull(rValue(7, 1)) Then    'TEE (?? 뭐였더라?)
                sngGab = typObesity.sngTEE - rValue(7, 1)
                Call DrawUpDown(sngGab, 7, "")
            Else
                Set imgUpDown(7).Picture = LoadPicture("")
                lblUpDown(7).Caption = "-"
            End If
            '3) 최고/최저
            '체중
            lblMax(0).Caption = MaxValue("Weight") & "kg"
            lblMin(0).Caption = MinValue("Weight") & "kg"
            '비만도
            lblMax(1).Caption = CInt(MaxValue("ObesityRate")) & "%"
            lblMin(1).Caption = CInt(MinValue("ObesityRate")) & "%"
            '허리둘레
            lblMax(2).Caption = MaxValue("Waist") & "cm"
            lblMin(2).Caption = MinValue("Waist") & "cm"
            'WHR
            lblMax(3).Caption = Format(MaxValue("WHR"), "0.00")
            lblMin(3).Caption = Format(MinValue("WHR"), "0.00")
            'BMI
            lblMax(4).Caption = Format(MaxValue("BMI"), "0.0")
            lblMin(4).Caption = Format(MinValue("BMI"), "0.0")
            '체지방률
            lblMax(5).Caption = Format(MaxValue("ChFatRate"), "0.0") & "%"
            lblMin(5).Caption = Format(MinValue("ChFatRate"), "0.0") & "%"
            'RMR
            lblMax(6).Caption = Format(MaxValue("RMR"), "#,###")
            lblMin(6).Caption = Format(MinValue("RMR"), "#,###")
            'TEE
            lblMax(7).Caption = Format(MaxValue("TEE"), "#,###")
            lblMin(7).Caption = Format(MinValue("TEE"), "#,###")
        End If
    Else
        With typObesity
            .sngWeight = 0
            .sngObesityRate = 0
            .sngWaist = 0
            .sngWHR = 0
            .sngBMI = 0
            .sngChFatRate = 0
            .sngRMR = 0
            .sngTEE = 0
        End With
        typTreatCal.sngTreatCal = 0

        For i = 0 To 7
            lblObInfo(i).Caption = ""
            Set imgUpDown(i).Picture = LoadPicture("")
            lblUpDown(i).Caption = ""
            lblMax(i).Caption = ""
            lblMin(i).Caption = ""
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

'2005-01-27 류진선 수정
'오늘날짜의 처방이 있으면 보여주기
Private Function ShowTodayTreat2() As Long
    Dim qrySelect As String, rValue As Variant
    
    qrySelect = "SELECT TOP 2 A.TreatNum, A.DietCalory, A.TreatCalory, A.UserCalory, "
    qrySelect = qrySelect & " A.ChMealTime, A.ChMealCode, "
    qrySelect = qrySelect & " Case A.ChMealTime When 10 Then C.Calory "
    qrySelect = qrySelect & " When 0 Then 0 Else B.ChMealCalory End "
    qrySelect = qrySelect & " FROM Treat A "
    qrySelect = qrySelect & " LEFT JOIN ChangeMeal B "
    qrySelect = qrySelect & " ON A.ChMealCode = B.ChMealName "
    qrySelect = qrySelect & " LEFT JOIN UserMenu C "
    qrySelect = qrySelect & " ON A.ChMealCode = C.UserMenuName "
    qrySelect = qrySelect & " WHERE A.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND A.TreatDay='" & Format(gdatUserDay, "YYYYMMDD") & "'"
    qrySelect = qrySelect & " ORDER BY A.TreatNum DESC "
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        mlngcTreatNum = CLng(rValue(0, 0))
        If (IsNull(rValue(1, 0)) Or IsNull(rValue(2, 0))) And (IsNull(rValue(2, 0)) Or IsNull(rValue(3, 0))) Then
            imgPreTreat.Visible = True
        Else
            imgPreTreat.Visible = False
        End If
        
        '2004-12-09 류진선 이놈들 왜 이리 왔다리 갔다리 하면서 데이터를 보여주는거지? ㅡㅡ;
        '==========
        '오른쪽 상단(?) 치료내용 보여주기
        Call ShowTreat(mlngcTreatNum)
'        '식단처방내용 보여주기
'        If Not IsNull(rValue(4, 0)) Then
'            If Trim(rValue(4, 0)) <> 0 Then
'                mlngChMealTime = Trim(rValue(4, 0))
'                msChMealCode = Trim(rValue(5, 0))
'                msngChMealCalory = Trim(rValue(6, 0))
'                chkChMeal.Value = vbChecked
'                Call setCboChMealTime(mlngChMealTime)
'                Call setCboChMealCode(msChMealCode)
'            Else
'                chkChMeal.Value = vbUnchecked
'            End If
'        End If
        
        '오른쪽하단 치료종목에 저장된 데이터 보여주기
        Call ShowTreatData(mlngcTreatNum)
        '오른쪽하단 출력물에 저장된 데이터 보여주기
        Call ShowTreatPrint(mlngcTreatNum)
        '==========
        
        Call SetGridRow
        ShowTodayTreat2 = UBound(rValue, 2) + 1
    Else
        Call ShowTreat(0)
        'Call ShowRecentEx
        Call imgNew_MouseUp(0, 0, 0, 0)
        ShowTodayTreat2 = 0
    End If
    
End Function

Private Sub SetGridRow()
Dim i As Integer
Dim selRow As Integer
selRow = -1
With grdTable
    .Col = 0
    .ColSel = 0
    
    '먼저 선택될 행을 찾는다
    .Row = 0    'TreatNum이 있는 컬럼
    For i = 1 To .Rows - 1
        If Trim(.TextMatrix(i, .Cols - 1)) = mlngcTreatNum Then
            selRow = i
            Exit For
        End If
    Next
    If selRow > -1 Then
        .Row = selRow
        .Col = 0
        .ColSel = .Cols - 1
        .TopRow = selRow
    End If
End With
End Sub

'2005-01-28 류진선 추가
Private Function ShowTodayBodyData() As Boolean
Dim qrySelect As String
Dim rValue
    qrySelect = "SELECT TOP 1 a.TreatNum "
    qrySelect = qrySelect & " FROM Treat a INNER JOIN BodyData b "
    qrySelect = qrySelect & " ON a.TreatNUM = b.TreatNUM "
    qrySelect = qrySelect & " WHERE a.TreatDay = '" & Format(gdatUserDay, "YYYY-MM-DD") & "' "
    qrySelect = qrySelect & " ORDER BY a.TreatDay DESC, a.TreatNum DESC "
    
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        mlngoTreatNum = Is_Null(rValue(0, 0), "0")
        ShowTodayBodyData = True
    Else
        mlngoTreatNum = 0
        ShowTodayBodyData = False
    End If

End Function
