VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCounsel_5 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDaily 
      Height          =   300
      Left            =   10500
      Style           =   2  '드롭다운 목록
      TabIndex        =   37
      Top             =   2220
      Width           =   1515
   End
   Begin ChartfxLibCtl.ChartFX chtWeek 
      Height          =   3045
      Left            =   750
      TabIndex        =   32
      Top             =   5190
      Width           =   7815
      _cx             =   13785
      _cy             =   5371
      Build           =   20
      TypeMask        =   44564482
      Volume          =   30
      AxesStyle       =   3
      Axis(0).Max     =   80
      Axis(0).TickMark=   -32767
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      RGBBk           =   16777215
      nColors         =   16
      Colors          =   "frmCounsel_5.frx":0000
      nPts            =   7
      nSer            =   1
      NumPoint        =   7
      NumSer          =   1
      BorderS         =   13
      _Data_          =   "frmCounsel_5.frx":00A0
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   300
      Left            =   10530
      Style           =   2  '드롭다운 목록
      TabIndex        =   7
      Top             =   1740
      Width           =   1455
   End
   Begin ChartfxLibCtl.ChartFX ChartFX1 
      Height          =   3045
      Left            =   750
      TabIndex        =   0
      Top             =   5190
      Width           =   7815
      _cx             =   13785
      _cy             =   5371
      Build           =   20
      TypeMask        =   111673345
      LeftGap         =   64
      BottomGap       =   43
      MarkerShape     =   0
      AxesStyle       =   3
      Axis(0).Max     =   3
      Axis(0).Decimals=   0
      Axis(0).TickMark=   -32767
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).TickMark=   -32767
      Axis(2).Format  =   5
      Axis(2).Format  =   5
      RGBBk           =   16777251
      nColors         =   16
      Colors          =   "frmCounsel_5.frx":01AD
      nSer            =   1
      NumSer          =   1
      Title(2)        =   "식사  횟수"
      BorderS         =   8
      _Data_          =   "frmCounsel_5.frx":024D
   End
   Begin ChartfxLibCtl.ChartFX chtTime 
      Height          =   3045
      Left            =   750
      TabIndex        =   1
      Top             =   5190
      Width           =   7815
      _cx             =   13785
      _cy             =   5371
      Build           =   20
      TypeMask        =   1183318017
      LeftGap         =   80
      MarkerShape     =   0
      AxesStyle       =   3
      Axis(0).Min     =   29221
      Axis(0).Max     =   29221
      Axis(0).Style   =   10280
      Axis(0).TickMark=   -32767
      Axis(0).Format  =   1
      Axis(0).Format  =   1
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).Format  =   5
      Axis(2).Format  =   5
      RGBBk           =   16777251
      nColors         =   16
      Colors          =   "frmCounsel_5.frx":035A
      nSer            =   3
      NumSer          =   3
      Title(2)        =   "식사시간"
      BorderS         =   8
      _Data_          =   "frmCounsel_5.frx":03FA
   End
   Begin ChartfxLibCtl.ChartFX chtSpeed 
      Height          =   3045
      Left            =   750
      TabIndex        =   2
      Top             =   5190
      Width           =   7815
      _cx             =   13785
      _cy             =   5371
      Build           =   20
      Axis(0).Decimals=   0
      Axis(2).Format  =   5
      Axis(2).Format  =   5
      RGBBk           =   16777251
      nColors         =   16
      Colors          =   "frmCounsel_5.frx":04F4
      nSer            =   3
      NumSer          =   3
      BorderS         =   8
      _Data_          =   "frmCounsel_5.frx":0594
   End
   Begin ChartfxLibCtl.ChartFX chtEatingOut 
      Height          =   3045
      Left            =   750
      TabIndex        =   3
      Top             =   5190
      Width           =   7815
      _cx             =   13785
      _cy             =   5371
      Build           =   20
      TypeMask        =   111673345
      LeftGap         =   64
      BottomGap       =   43
      MarkerShape     =   0
      AxesStyle       =   3
      Axis(0).Max     =   3
      Axis(0).Decimals=   0
      Axis(0).TickMark=   -32767
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).TickMark=   -32767
      Axis(2).Format  =   5
      Axis(2).Format  =   5
      RGBBk           =   16777251
      nColors         =   16
      Colors          =   "frmCounsel_5.frx":05EC
      nSer            =   1
      NumSer          =   1
      BorderS         =   8
      _Data_          =   "frmCounsel_5.frx":068C
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   285
      Left            =   10650
      TabIndex        =   4
      Top             =   2070
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      Format          =   23789569
      CurrentDate     =   37818
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   285
      Left            =   10650
      TabIndex        =   5
      Top             =   2370
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      Format          =   23789569
      CurrentDate     =   37818
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   9
      Left            =   10500
      Top             =   6750
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   8
      Left            =   10500
      Top             =   6330
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   7
      Left            =   10500
      Top             =   5940
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   6
      Left            =   10500
      Top             =   5550
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   5
      Left            =   10500
      Top             =   5160
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   4
      Left            =   10500
      Top             =   4740
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   3
      Left            =   10500
      Top             =   4350
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   2
      Left            =   10500
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   1
      Left            =   10500
      Top             =   3570
      Width           =   1485
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   0
      Left            =   10500
      Top             =   3150
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   8220
      Picture         =   "frmCounsel_5.frx":0799
      Top             =   1380
      Width           =   915
   End
   Begin VB.Image imgStart 
      Height          =   300
      Left            =   9300
      Picture         =   "frmCounsel_5.frx":0EBF
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label lblFeel 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "좋음"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   6660
      TabIndex        =   36
      Top             =   8820
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblFeel 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "좋음"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   7920
      TabIndex        =   35
      Top             =   4410
      Width           =   765
   End
   Begin VB.Label lblFeel 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "좋음"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   7920
      TabIndex        =   34
      Top             =   3510
      Width           =   765
   End
   Begin VB.Label lblFeel 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "좋음"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   7920
      TabIndex        =   33
      Top             =   2610
      Width           =   765
   End
   Begin VB.Label lblCalories 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "700 kcal (25%) 지방 10.9g(11%)"
      Height          =   345
      Index           =   3
      Left            =   7470
      TabIndex        =   31
      Top             =   8820
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblCalories 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "700 kcal (25%) 지방 10.9g(11%)"
      Height          =   345
      Index           =   2
      Left            =   8730
      TabIndex        =   30
      Top             =   4410
      Width           =   1305
   End
   Begin VB.Label lblCalories 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "700 kcal (25%) 지방 10.9g(11%)"
      Height          =   345
      Index           =   1
      Left            =   8730
      TabIndex        =   29
      Top             =   3510
      Width           =   1305
   End
   Begin VB.Label lblCalories 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "700 kcal (25%) 지방 10.9g(11%)"
      Height          =   345
      Index           =   0
      Left            =   8730
      TabIndex        =   28
      Top             =   2610
      Width           =   1305
   End
   Begin VB.Image imgFeeling 
      Height          =   555
      Index           =   3
      Left            =   6780
      Top             =   8610
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgFeeling 
      Height          =   555
      Index           =   2
      Left            =   8040
      Top             =   4260
      Width           =   555
   End
   Begin VB.Image imgFeeling 
      Height          =   555
      Index           =   1
      Left            =   8040
      Top             =   3300
      Width           =   555
   End
   Begin VB.Image imgFeeling 
      Height          =   555
      Index           =   0
      Left            =   8040
      Top             =   2430
      Width           =   555
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "15분최대 : 25분최소 : 20분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   5790
      TabIndex        =   27
      Top             =   8640
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "15분최대 : 25분최소 : 20분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   7050
      TabIndex        =   26
      Top             =   4260
      Width           =   795
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "15분최대 : 25분최소 : 20분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   7050
      TabIndex        =   25
      Top             =   3330
      Width           =   795
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "15분최대 : 25분최소 : 20분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   7050
      TabIndex        =   24
      Top             =   2430
      Width           =   795
   End
   Begin VB.Label lblFeeling 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "약간배고픈정도"
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
      Left            =   4650
      TabIndex        =   23
      Top             =   8820
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblFeeling 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "약간배고픈정도"
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
      Left            =   5910
      TabIndex        =   22
      Top             =   4410
      Width           =   975
   End
   Begin VB.Label lblFeeling 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "약간배고픈정도"
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
      Left            =   5910
      TabIndex        =   21
      Top             =   3510
      Width           =   975
   End
   Begin VB.Label lblFeeling 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "아주배부름"
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
      Left            =   5910
      TabIndex        =   20
      Top             =   2610
      Width           =   975
   End
   Begin VB.Label lblPlace 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사무실"
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
      Left            =   2790
      TabIndex        =   19
      Top             =   8790
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblPlace 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사무실"
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
      Left            =   4050
      TabIndex        =   18
      Top             =   4410
      Width           =   585
   End
   Begin VB.Label lblPlace 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사무실"
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
      Left            =   4050
      TabIndex        =   17
      Top             =   3510
      Width           =   585
   End
   Begin VB.Label lblPlace 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사무실"
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
      Left            =   4050
      TabIndex        =   16
      Top             =   2610
      Width           =   585
   End
   Begin VB.Label lblTime 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "오전 5:30 최대 5:30 최소 5:30"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   1800
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblTime 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "오전 5:30 최대 5:30 최소 5:30"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   3060
      TabIndex        =   14
      Top             =   4260
      Width           =   885
   End
   Begin VB.Label lblTime 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "오전 5:30 최대 5:30 최소 5:30"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   3060
      TabIndex        =   13
      Top             =   3330
      Width           =   885
   End
   Begin VB.Label lblTime 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "오전 5:30 최대 5:30 최소 5:30"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   3060
      TabIndex        =   12
      Top             =   2430
      Width           =   885
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2일 / 4일   (50%)"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   11
      Top             =   8730
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2일 / 4일   (50%)"
      Height          =   375
      Index           =   2
      Left            =   1650
      TabIndex        =   10
      Top             =   4350
      Width           =   1365
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2일 / 4일   (50%)"
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   9
      Top             =   3420
      Width           =   1365
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "2일 / 4일 (100%)"
      Height          =   375
      Index           =   0
      Left            =   1620
      TabIndex        =   8
      Top             =   2520
      Width           =   1365
   End
   Begin VB.Image imgSub 
      Height          =   375
      Index           =   4
      Left            =   8820
      Top             =   7230
      Width           =   1005
   End
   Begin VB.Image imgSub 
      Height          =   375
      Index           =   3
      Left            =   8820
      Top             =   6750
      Width           =   1005
   End
   Begin VB.Image imgSub 
      Height          =   375
      Index           =   2
      Left            =   8820
      Top             =   6300
      Width           =   1005
   End
   Begin VB.Image imgSub 
      Height          =   375
      Index           =   1
      Left            =   8820
      Picture         =   "frmCounsel_5.frx":17E6
      Top             =   5820
      Width           =   1005
   End
   Begin VB.Image imgSub 
      Height          =   375
      Index           =   0
      Left            =   8820
      Picture         =   "frmCounsel_5.frx":20A2
      Top             =   5370
      Width           =   1005
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "~"
      Height          =   195
      Index           =   1
      Left            =   10440
      TabIndex        =   6
      Top             =   2430
      Width           =   165
   End
   Begin VB.Image imgPrint 
      Height          =   975
      Left            =   10680
      Top             =   7350
      Width           =   1155
   End
End
Attribute VB_Name = "frmCounsel_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 아직 덜 한 부분===============================================(2004.2.25)
' 1. 상단 표에서 식사시간 부분 아직 못함 - 완료(2004.3.8)
' 2. 감정의 이미지 아직 덜 받음. -> 이미지 말고 그냥 글자로 보여주기로 함 2004.3.8
' 3. 기간선택에서 dtpicker의 lostfocus에서 평가 보여주는 부분 수정할 것
' 4. 평균을 단순 평균값이 아닌 갯수가 가장 많은 것으로 보여줌(식사장소, 식후배고픔정도, 느낌) 2004.3.8
'======================================================================
Option Explicit
Private Type mCustomerInfo
    intState As Integer
    intAge As Integer
    strSex As String
    sngDietCal As Single
End Type
Private typCustomerInfo As mCustomerInfo

Dim crxApplication As New CRAXDRT.Application
Public crxReport As CRAXDRT.Report
Public crxReport2 As CRAXDRT.Report
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxFormula As CRAXDRT.FormulaFieldDefinition
Dim strServer As String, strDBName As String, strUID As String, strPWD As String
'+---------------------------------------------------------------------------------+
'| 상담 > 식사 > 식습관평가
'+---------------------------------------------------------------------------------+
Private Const PATH05 As String = "\Back\Counsel\05\"
Private Const IMG_FEEL As String = "\Back\Counsel\"

Private Const IMG_SUB1_ON As String = "식사횟수 on.jpg"
Private Const IMG_SUB1_OFF As String = "식사횟수 off.jpg"
Private Const IMG_SUB2_ON As String = "식사시간 on.jpg"
Private Const IMG_SUB2_OFF As String = "식사시간 off.jpg"
Private Const IMG_SUB3_ON As String = "식사속도 on.jpg"
Private Const IMG_SUB3_OFF As String = "식사속도 off.jpg"
Private Const IMG_SUB4_ON As String = "외식횟수 on.jpg"
Private Const IMG_SUB4_OFF As String = "외식횟수 off.jpg"
Private Const IMG_SUB5_ON As String = "요일분석 on.jpg"
Private Const IMG_SUB5_OFF As String = "요일분석 off.jpg"

Private Sub cmbPeriod_Click()
    Select Case cmbPeriod.ListIndex
        Case 0:   '특정일
            dtpBegin.Visible = False
            dtpEnd.Visible = False
            Label5(1).Visible = False
            cmbDaily.Visible = True
            Call ShowVal
        Case 1:   '특정기간
            cmbDaily.Visible = False
            dtpBegin.Visible = True
            dtpEnd.Visible = True
            Label5(1).Visible = True
        Case 2:   '전체기간
            cmbDaily.Visible = False
            dtpBegin.Visible = False
            dtpEnd.Visible = False
            Label5(1).Visible = False
            Call ShowVal
    End Select
End Sub

Private Sub cmbPeriod_Change()
    Call cmbPeriod_Click
End Sub

Private Sub dtpBegin_LostFocus()
    If dtpBegin.Value <= dtpEnd.Value Then
    Else
        MsgBox "시작일은 종료일 보다 앞서야 합니다.", vbOKOnly + vbExclamation
    End If
End Sub

Private Sub dtpEnd_LostFocus()
    If dtpBegin.Value <= dtpEnd.Value Then
    Else
        MsgBox "시작일은 종료일 보다 앞서야 합니다.", vbOKOnly + vbExclamation
    End If
End Sub

Public Sub Form_Load()
    Dim i As Integer
    
    Set Me.Picture = LoadPicture(App.Path & "\Back\Counsel\05\05.jpg")
    
    Me.Width = FRM_WIDTH
    Me.Height = FRM_HEIGHT
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.BackColor = vbWhite

'입력한 일기 중에 가장 최종일 영양평가를 보여줌
'만약 현재 선택된 일기가 있다면 그것을 보여줌- 현재 몇일치의 영양평가를 보여주고 있는지 보여줄 것
    If ExistDiary = False Then
        MsgBox "입력한 식사일기가 없습니다. ", vbOKOnly + vbExclamation
        Call InitialLabel
        Set imgSub(0).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB1_ON)
        Set imgSub(1).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB2_OFF)
        Set imgSub(2).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB3_OFF)
        Set imgSub(3).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB4_OFF)
        Set imgSub(4).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB5_OFF)
        For i = 0 To 4
            imgSub(i).Enabled = False
        Next i
        
       ChartFX1.Visible = False
        chtTime.Visible = False
        chtSpeed.Visible = False
        chtWeek.Visible = False
        chtWeek.Visible = False
        
        cmbDaily.Visible = False
        dtpBegin.Visible = False
        dtpEnd.Visible = False
        cmbPeriod.Enabled = False
        Exit Sub
    End If
    
    Call InitialLabel
    Call InitialDailyCombo
   
    cmbPeriod.Enabled = True
    cmbPeriod.Clear
    cmbPeriod.AddItem "특정일"
    cmbPeriod.AddItem "특정기간"
    cmbPeriod.AddItem "전체"
    cmbPeriod.ListIndex = 0
    
    dtpBegin.Value = Now()
    dtpEnd.Value = Now()
    
    Call imgSub_Click(0)
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub imgAppend_Click(Index As Integer)
    frmPop_Additional1.mintNumber = Index + 1
    frmPop_Additional1.Show vbModal
End Sub

Private Sub imgPrint_Click()
    Call PrintData
End Sub

Private Sub ShowVal()
    Dim qrySelect As String, rValue As Variant
    Set clsSelect = New clsSelect
    '기간내 입력된 일기가 있는지 먼저 체크할 것
    qrySelect = "SELECT DISTINCT(MealDate) FROM DietDiary WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    rValue = clsSelect.Query(qrySelect)
    If IsNull(rValue) Then
       MsgBox "기간내에 입력된 식사일기가 없습니다.", vbOKOnly + vbExclamation
       Exit Sub
    End If
    Set clsSelect = Nothing
    Call TopCount
    Call TopTime
    Call TopPlace2
    Call TopAfterHungry2
    Call TopSpeed
    Call TopFeeling2
    Call TopCalories

    '식습관평가
    Call InitialCountHaveMeal
    Call CountHaveMeal

    Call InitialTimeHaveMeal
    Call TimeHaveMeal

    Call InitialSpeedHaveMeal
    Call SpeedHaveMeal

    Call InitialEatingOut
    Call EatingOut
    
    '요일별 평가
    Call InitialWeek
    Call WeekCalories

    '일별평가
'    Call LoadMealCalory

    '고객정보에서 다이어트열량, 신체상태, 나이, 성별등을 불러옴
    Call LoadCustomerInfo(glngCustomerNum)
    If typCustomerInfo.sngDietCal = 0 Then
        Exit Sub
    End If

    With typCustomerInfo
        If Calculate_Nut(.sngDietCal, .intState, .intAge, .strSex) = True Then
            '끼니별평가
            Call MealSectionRate

        End If
    End With
End Sub

Private Sub InitialLabel()
    Dim i As Integer
    
    For i = 0 To 3
        lblCount(i).Caption = ""
        lblTime(i).Caption = ""
        lblPlace(i).Caption = ""
        lblFeeling(i).Caption = ""
        lblSpeed(i).Caption = ""
        lblFeel(i).Caption = ""
        Set imgFeeling(i).Picture = LoadPicture("")
        lblCalories(i).Caption = ""
    Next i
End Sub

Private Sub InitialDailyCombo()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT DISTINCT(MealDate) FROM DietDiary WHERE CustomerNum=" & glngCustomerNum
    
    rValue = clsSelect.Query(qrySelect)
    cmbDaily.Clear
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            cmbDaily.AddItem Trim(rValue(0, i))
        Next i
        cmbDaily.ListIndex = UBound(rValue, 2)
    End If
    Set clsSelect = Nothing
End Sub

Private Sub TopCount()
    Dim qrySelect As String, rValue As Variant
    Dim intB As Integer, intL As Integer, intD As Integer, intT As Integer
    Dim intAve As Integer
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT SUM(a), SUM(b), SUM(c), SUM(total) FROM("
    qrySelect = qrySelect & "SELECT COUNT(DietDiaryNum) AS a, 0 AS b, 0 AS c, 0 AS total FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=1 AND NOT MealCalory IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, COUNT(DietDiaryNum) AS b, 0 AS c, 0 AS total FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=2 AND NOT MealCalory IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, COUNT(DietDiaryNum) AS c, 0 AS total FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=3 AND NOT MealCalory IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, 0 AS c, COUNT(DISTINCT(MealDate)) AS total FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection<>4 AND NOT MealCalory IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & ") i"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        intB = rValue(0, 0)
        intL = rValue(1, 0)
        intD = rValue(2, 0)
        intT = rValue(3, 0)
        intAve = (intB + intL + intD) / 3
        
        lblCount(0).Caption = intB & "일 / " & intT & "일" & vbNewLine & "(" & CInt((intB / intT) * 100) & "%)"
        lblCount(1).Caption = intL & "일 / " & intT & "일" & vbNewLine & "(" & CInt((intL / intT) * 100) & "%)"
        lblCount(2).Caption = intD & "일 / " & intT & "일" & vbNewLine & "(" & CInt((intD / intT) * 100) & "%)"
        lblCount(3).Caption = intAve & "일 / " & intT & "일" & vbNewLine & "(" & CInt((intAve / intT) * 100) & "%)"
    Else
        lblCount(0).Caption = ""
        lblCount(1).Caption = ""
        lblCount(2).Caption = ""
        lblCount(3).Caption = ""
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub TopTime()
'각 시간을 분으로 환산해서 평균값을 구한다.

'2005-01-31 류진선 Int형 최대값(32767) 초과 수정
    Dim lngCount As Long, lngSum As Long
    Dim lngHour As Long, lngMin As Long, lngTotal As Long
    
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer
    
    For j = 0 To 2
        qrySelect = "SELECT MealTime FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & " AND NOT MealTime IS NULL "
        qrySelect = qrySelect & "AND MealSection=" & j + 1
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        
        Set clsSelect = New clsSelect
    
        rValue = clsSelect.Query(qrySelect)
        
        Set clsSelect = Nothing
'2005-01-31 류진선 Int형 최대값(32767) 초과 수정
        lngHour = 0
        lngMin = 0
        lngTotal = 0
        lngSum = 0
        If Not IsNull(rValue) Then
'2005-01-31 류진선 Int형 최대값(32767) 초과 수정
            lngCount = UBound(rValue, 2)
            For i = 0 To lngCount
                lngHour = CInt(Left(rValue(0, i), 2))
                lngMin = CInt(Right(rValue(0, i), 2))
                lngTotal = (lngHour * 60) + lngMin
                lngSum = lngSum + lngTotal
            Next i
            lngSum = lngSum / (lngCount + 1)
            lngHour = lngSum / 60
            lngMin = lngSum Mod 60
            lblTime(j).Caption = Format(lngHour, "00") & ":" & Format(lngMin, "00")
            lblTime(j).Caption = lblTime(j).Caption & vbNewLine & MaxMinTime(j + 1)
        Else
            lblTime(j).Caption = ""
        End If
    Next j
End Sub

Private Function MaxMinTime(intMealSection As Integer) As String
    Dim qrySelect As String, rValue As Variant
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT MAX(MealTime), MIN(MealTime) FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND NOT MealTime IS NULL "
    qrySelect = qrySelect & "AND MealSection=" & intMealSection
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        MaxMinTime = "최대 " & Format(Is_Null(rValue(0, 0), "0000"), "00:00")
        MaxMinTime = MaxMinTime & vbNewLine & "최소 " & Format(Is_Null(rValue(1, 0), "0000"), "00:00")
    Else
        MaxMinTime = ""
    End If
    
    Set clsSelect = Nothing
End Function

Private Sub TopPlace()
    Dim qrySelect  As String, rValue As Variant
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT SUM(a), SUM(b), SUM(c), SUM(t) FROM("
    qrySelect = qrySelect & "SELECT AVG(MealPlace) AS a, 0 AS b, 0 AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=1 AND NOT MealPlace IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, AVG(MealPlace) AS b, 0 AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=2 AND NOT MealPlace IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, AVG(MealPlace) AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=3 AND NOT MealPlace IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, 0 AS c, AVG(MealPlace) AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection<>4 AND NOT MealPlace IS NULL) i"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To 3
            Select Case rValue(i, 0)
                Case 1: lblPlace(i).Caption = "집"
                Case 2: lblPlace(i).Caption = "사무실"
                Case 3: lblPlace(i).Caption = "외식"
                Case Else: lblPlace(i).Caption = ""
            End Select
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub TopPlace2()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer
    Dim intMax As Integer, intPlace As Integer
    
    For i = 0 To 3
        qrySelect = "SELECT MealPlace, COUNT(MealPlace) "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & " AND NOT MealPlace IS NULL "
        If i <> 3 Then
            qrySelect = qrySelect & " AND MealSection=" & i + 1
        Else
            qrySelect = qrySelect & " AND MealSection<>4"
        End If
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & " GROUP BY MealPlace;"
        Set clsSelect = New clsSelect
        rValue = clsSelect.Query(qrySelect)
        Set clsSelect = Nothing
        '가장 갯수가 많은 MealPlace 찾기
        If Not IsNull(rValue) Then
            intMax = 0
            intPlace = 0
            For j = 0 To UBound(rValue, 2)
                If j = 0 Then
                    intMax = rValue(1, 0)
                    intPlace = rValue(0, 0)
                ElseIf intMax < rValue(1, j) Then
                    intMax = rValue(1, j)
                    intPlace = rValue(0, j)
                End If
            Next j
            Select Case intPlace
                Case 1: lblPlace(i).Caption = "집"
                Case 2: lblPlace(i).Caption = "사무실"
                Case 3: lblPlace(i).Caption = "외식"
                Case Else: lblPlace(i).Caption = ""
            End Select
        Else
            lblPlace(i).Caption = ""
        End If
    Next i

End Sub

Private Sub TopAfterHungry()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT SUM(a), SUM(b), SUM(c), SUM(t) FROM("
    qrySelect = qrySelect & "SELECT AVG(AfterMealHungry) AS a, 0 AS b, 0 AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=1 AND NOT AfterMealHungry IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, AVG(AfterMealHungry) AS b, 0 AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=2 AND NOT AfterMealHungry IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, AVG(AfterMealHungry) AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=3 AND NOT AfterMealHungry IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, 0 AS c, AVG(AfterMealHungry) AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection<>4 AND NOT AfterMealHungry IS NULL) i"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To 3
            Select Case rValue(i, 0)
                Case 1: lblFeeling(i).Caption = "아주배고픔"
                Case 2: lblFeeling(i).Caption = "조금배고픔"
                Case 3: lblFeeling(i).Caption = "보통"
                Case 4: lblFeeling(i).Caption = "조금배부름"
                Case 5: lblFeeling(i).Caption = "아주배부름"
                Case Else: lblFeeling(i).Caption = ""
            End Select
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub TopAfterHungry2()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer
    Dim intMax As Integer, intAfter As Integer
    
    For i = 0 To 3
        qrySelect = "SELECT AfterMealHungry, COUNT(AfterMealHungry) "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & " AND NOT AfterMealHungry IS NULL "
        If i <> 3 Then
            qrySelect = qrySelect & " AND MealSection=" & i + 1
        Else
            qrySelect = qrySelect & " AND MealSection<>4"
        End If
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & " GROUP BY AfterMealHungry;"
        Set clsSelect = New clsSelect
        rValue = clsSelect.Query(qrySelect)
        Set clsSelect = Nothing
        '가장 갯수가 많은 AfterMealHungry 찾기
        If Not IsNull(rValue) Then
            intMax = 0
            intAfter = 0
            For j = 0 To UBound(rValue, 2)
                If j = 0 Then
                    intMax = rValue(1, 0)
                    intAfter = rValue(0, 0)
                ElseIf intMax < rValue(1, j) Then
                    intMax = rValue(1, j)
                    intAfter = rValue(0, j)
                End If
            Next j
            Select Case intAfter
                Case 1: lblFeeling(i).Caption = "아주배고픔"
                Case 2: lblFeeling(i).Caption = "조금배고픔"
                Case 3: lblFeeling(i).Caption = "보통"
                Case 4: lblFeeling(i).Caption = "조금배부름"
                Case 5: lblFeeling(i).Caption = "아주배부름"
                Case Else: lblFeeling(i).Caption = ""
            End Select
        Else
            lblFeeling(i).Caption = ""
        End If
    Next i

End Sub

Private Sub TopSpeed()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT SUM(a), SUM(xa), SUM(na), SUM(b), SUM(xb), SUM(nb), "
    qrySelect = qrySelect & "SUM(c), SUM(xc), SUM(nc), SUM(t), SUM(xt), SUM(nt) FROM("
    qrySelect = qrySelect & "SELECT AVG(MealNeedTime) AS a, 0 AS b, 0 AS c, 0 AS t, "
    qrySelect = qrySelect & "MAX(MealNeedTime) AS xa, MIN(MealNeedTime) AS na, 0 AS xb, 0 AS nb, "
    qrySelect = qrySelect & "0 AS xc, 0 AS nc, 0 AS xt, 0 AS nt "
    qrySelect = qrySelect & "FROM DietDiary WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=1 AND NOT MealNeedTime IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, AVG(MealNeedTime) AS b, 0 AS c, 0 AS t, "
    qrySelect = qrySelect & "0 AS xa, 0 AS na, MAX(MealNeedTime) AS xb, MIN(MealNeedTime) AS nb, "
    qrySelect = qrySelect & "0 AS xc, 0 AS nc, 0 AS xt, 0 AS nt "
    qrySelect = qrySelect & "FROM DietDiary WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=2 AND NOT MealNeedTime IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, AVG(MealNeedTime) AS c, 0 AS t, "
    qrySelect = qrySelect & "0 AS xa, 0 AS na, 0 AS xb, 0 AS nb, "
    qrySelect = qrySelect & "MAX(MealNeedTime) AS xc, MIN(MealNeedTime) AS nc, 0 AS xt, 0 AS nt "
    qrySelect = qrySelect & "FROM DietDiary WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=3 AND NOT MealNeedTime IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, 0 AS c, AVG(MealNeedTime) AS t, "
    qrySelect = qrySelect & "0 AS xa, 0 AS na, 0 AS xb, 0 AS nb, "
    qrySelect = qrySelect & "0 AS xc, 0 AS nc, MAX(MealNeedTime) AS xt, MIN(MealNeedTime) AS nt "
    qrySelect = qrySelect & "FROM DietDiary WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection<>4 AND NOT MealNeedTime IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & ") i"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To 3
        lblSpeed(i).Caption = rValue(i * 3, 0) & "분" & vbNewLine & "최대 " & rValue(i * 3 + 1, 0) & "분" & vbNewLine & "최소 " & rValue(i * 3 + 2, 0) & "분"
        Next i
    Else
        For i = 0 To 3
        lblSpeed(i).Caption = ""
        Next i
    End If
   
    Set clsSelect = Nothing
End Sub

Private Sub TopFeeling()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT SUM(a), SUM(b), SUM(c), SUM(t) FROM("
    qrySelect = qrySelect & "SELECT AVG(MealFeeling) AS a, 0 AS b, 0 AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=1 AND NOT MealFeeling IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, AVG(MealFeeling) AS b, 0 AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=2 AND NOT MealFeeling IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, AVG(MealFeeling) AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=3 AND NOT MealFeeling IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, 0 AS c, AVG(MealFeeling) AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection<>4 AND NOT MealFeeling IS NULL) i"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To 3
           Select Case rValue(i, 0)
                Case 1: lblFeel(i).Caption = "좋음"
                Case 2: lblFeel(i).Caption = "짜증"
                Case 3: lblFeel(i).Caption = "우울"
                Case 4: lblFeel(i).Caption = "피곤"
                Case 5: lblFeel(i).Caption = "긴장"
                Case 6: lblFeel(i).Caption = "바쁨"
                Case 7: lblFeel(i).Caption = "보통"
                Case 8: lblFeel(i).Caption = "기타"
                Case Else: lblFeel(i).Caption = ""
            End Select
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub TopFeeling2()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer
    Dim intMax As Integer, intFeeling As Integer
    
    For i = 0 To 3
        qrySelect = "SELECT MealFeeling, COUNT(MealFeeling) "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & " AND NOT MealFeeling IS NULL "
        If i <> 3 Then
            qrySelect = qrySelect & " AND MealSection=" & i + 1
        Else
            qrySelect = qrySelect & " AND MealSection<>4"
        End If
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & " GROUP BY MealFeeling;"
        Set clsSelect = New clsSelect
        rValue = clsSelect.Query(qrySelect)
        Set clsSelect = Nothing
        '가장 갯수가 많은 MealFeeling 찾기
        If Not IsNull(rValue) Then
            intMax = 0
            intFeeling = 0
            For j = 0 To UBound(rValue, 2)
                If j = 0 Then
                    intMax = rValue(1, 0)
                    intFeeling = rValue(0, 0)
                ElseIf intMax < rValue(1, j) Then
                    intMax = rValue(1, j)
                    intFeeling = rValue(0, j)
                End If
            Next j
            Select Case intFeeling
                Case 1: lblFeel(i).Caption = "좋음"
                Case 2: lblFeel(i).Caption = "짜증"
                Case 3: lblFeel(i).Caption = "우울"
                Case 4: lblFeel(i).Caption = "피곤"
                Case 5: lblFeel(i).Caption = "긴장"
                Case 6: lblFeel(i).Caption = "바쁨"
                Case 7: lblFeel(i).Caption = "보통"
                Case 8: lblFeel(i).Caption = "기타"
                Case Else: lblFeel(i).Caption = ""
            End Select
        Else
            lblFeel(i).Caption = ""
        End If
    Next i

End Sub

Private Sub TopCalories()
    Dim qrySelect  As String, rValue As Variant
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT SUM(a), SUM(b), SUM(c), SUM(t) FROM("
    qrySelect = qrySelect & "SELECT AVG(MealCalory) AS a, 0 AS b, 0 AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=1 AND NOT MealCalory IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, AVG(MealCalory) AS b, 0 AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=2 AND NOT MealCalory IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, AVG(MealCalory) AS c, 0 AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection=3 AND NOT MealCalory IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " UNION ALL "
    qrySelect = qrySelect & "SELECT 0 AS a, 0 AS b, 0 AS c, AVG(MealCalory) AS t FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection<>4 AND NOT MealCalory IS NULL"
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & ") i"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To 3
            lblCalories(i).Caption = Format(rValue(i, 0), "#,##0 kcal")
        Next i
    Else
        For i = 0 To 3
            lblCalories(i).Caption = "0 kcal"
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub LoadMealCalory()
'현재 해당환자의 저장된 식사일기들을 선택된 기간내에 해당하는 것을 불러서 보여준다
    Dim qrySelect As String
    Dim rValue As Variant
    Dim i As Integer

    Set clsSelect = New clsSelect

    qrySelect = "SELECT DISTINCT MealDate "
    qrySelect = qrySelect & "FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If

    rValue = clsSelect.Query(qrySelect)

    If IsNull(rValue) Then
        MsgBox "기간내에 입력된 식사일기가 없습니다.", vbInformation
        Call InitialCountHaveMeal
        Call InitialTimeHaveMeal
        Call InitialSpeedHaveMeal
        Call InitialEatingOut
        Exit Sub
    End If

    Set clsSelect = Nothing
End Sub

'영양평가를 위한 테이블
Private Function Calculate_Nut(sngDietCal As Single, intState As Integer, intAge As Integer, strSex As String) As Boolean
    Dim rValue As Variant
    Dim rValue2 As Variant
    Dim qrySelect As String

    '   s(영양요소, 끼니코드 1=아침   4=간식, 0=합)
    Dim s(1 To 36, 0 To 4) As Single, temp As Single
    '   각 끼니를 해당기간내에 먹은 횟수 1:아침~4:간식,0은 기간내 입력한 일기수
    '   해당기간내에 평균을 내기 위함..
    Dim intSectionCnt(0 To 4) As Integer

    ' 각 영양소, 끼니별 count
    Dim c(0 To 19, 0 To 4) As Integer
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim strEelem(1 To 4) As String
    Dim strTime(1 To 6) As String

    For j = 0 To 4
        For i = 1 To 36
            s(i, j) = 0
        Next i
        For i = 0 To 19
            c(i, j) = 0
        Next i
    Next j
    strTime(1) = "아침"
    strTime(2) = "점심"
    strTime(3) = "저녁"
    strTime(4) = "간식"
    strTime(5) = "1일 합계"
    strTime(6) = "권장량대비%"

    strEelem(1) = "열량"
    strEelem(2) = "단백질"
    strEelem(3) = "지방"
    strEelem(4) = "탄수화물"

    Set clsSelect = New clsSelect

    qrySelect = "SELECT Count(a.MealDate) FROM"
    qrySelect = qrySelect & "(SELECT MealDate FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY MealDate) a"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        intSectionCnt(0) = CInt(rValue(0, 0))
    End If

    For i = 1 To 4
        qrySelect = "SELECT COUNT(DietDiaryNum) FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & " AND MealSection=" & i
        qrySelect = qrySelect & " AND MealCalory IS NOT NULL;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            intSectionCnt(i) = CInt(rValue(0, 0))
        End If
    Next i
    
    qrySelect = "SELECT MealCode,"                                                       '0
    qrySelect = qrySelect & "Energy,Protein,Fat,Carbohy,Fiber,"                          '1-5
    qrySelect = qrySelect & "Ash,Ca,P,Fe,Na,"                                            '6-10
    qrySelect = qrySelect & "K,Zn,Vitamine_A,Retinol,Carotene,"                          '11-15
    qrySelect = qrySelect & "Vitamine_B1,Vitamine_B2,Vitamine_B6,Niacin,Vitamine_C,"     '16-20
    qrySelect = qrySelect & "Folic,Vitamine_E,Cholesterol,Waste,DietFiberDry,"           '21-25
    qrySelect = qrySelect & "DietFiberWet,Vitamine_B12,Vitamine_D,MealSection,FoodCode," '26-30
    qrySelect = qrySelect & "FoodWeight,FK_PartID,NatureID "                             '31,32,33
    qrySelect = qrySelect & "FROM DietDiary INNER JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "INNER JOIN DietFood ON DietMeal.DietMealNum=DietFood.DietMealNum "
    qrySelect = qrySelect & "INNER JOIN tblFood ON DietFood.FoodCode=tblFood.FoodID "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            'k is 끼니
            k = rValue(29, i)
            For j = 1 To 28
                If Not IsNull(rValue(j, i)) Then
                    '모든 영양소는 100g에 대한 값임
                    temp = rValue(j, i) * rValue(31, i) / 100
                    s(j, k) = s(j, k) + temp
                    s(j, 0) = s(j, 0) + temp
                Else
                    temp = 0
                End If

                '동물성/식물성 섭취비율
                If j = 2 Or j = 3 Or j = 7 Or j = 9 Then
                    Select Case j
                       Case 2        '단백질
                            l = 29
                       Case 3        '지방
                            l = 31
                       Case 7        '칼슘
                            l = 33
                       Case 9        '철분
                            l = 35
                    End Select

                    If rValue(33, i) = "1" Then      '식물성
                        If Not IsNull(rValue(j, i)) Then
                            s(l, k) = s(l, k) + temp
                            s(l, 0) = s(l, 0) + temp
                        End If
                    ElseIf rValue(33, i) = "2" Then  '동물성
                        If Not IsNull(rValue(j, i)) Then
                            s(l + 1, k) = s(l + 1, k) + temp
                            s(l + 1, 0) = s(l + 1, 0) + temp
                        End If
                    End If
                End If
            Next j
            c(rValue(32, i), k) = c(rValue(32, i), k) + 1
            c(rValue(32, i), 0) = c(rValue(32, i), 0) + 1
        Next i
        Erase rValue
        '평균을 내자 : 동,식물성섭취비율과 19가지 식품군은 평균안냄
        For j = 0 To 4
            For i = 1 To 28
                If intSectionCnt(j) <> 0 Then
                    s(i, j) = s(i, j) / intSectionCnt(j)
                End If
            Next i
        Next j
        '/////////////

        qrySelect = "DELETE FROM Nutrion WHERE CustomerNum=" & glngCustomerNum
        Call modSql.AdoExcuteSql(qrySelect)

        qrySelect = "DELETE FROM NutrionCont WHERE CustomerNum=" & glngCustomerNum
        Call modSql.AdoExcuteSql(qrySelect)

        If strSex = "M" Then
            intState = 1
        End If

        qrySelect = "SELECT ID, m1, m2, m3, m4, m5,m6, m7, m8, m9, m10,"
        qrySelect = qrySelect & "m11, m12, m13, m14, m15,m16, m17, m18, m19, m20,"
        qrySelect = qrySelect & "m21, m22, m23, m24, m25,m26, m27, m28 "
        qrySelect = qrySelect & "FROM Recommand WHERE Gender ='" & strSex & "' AND "
        qrySelect = qrySelect & "BodyState = " & intState
        qrySelect = qrySelect & " AND AgeLow <= " & intAge & " AND AgeHigh > " & intAge
        rValue2 = clsSelect.Query(qrySelect)

        Set clsSelect = Nothing

        Dim qryInsert As String
        For i = 0 To 4
            If i = 0 Then
                l = 5
            Else
                l = i
            End If
            qryInsert = "INSERT INTO Nutrion(CustomerNum, bt, btname, m1, m2, m3, m4, m5"
            qryInsert = qryInsert & ",m6, m7, m8, m9, m10,m11, m12, m13, m14, m15"
            qryInsert = qryInsert & ",m16, m17, m18, m19, m20,m21, m22, m23, m24, m25"
            qryInsert = qryInsert & ",m26, m27, m28, m29, m30,m31, m32, m33, m34, m35,m36) "
            qryInsert = qryInsert & "VALUES(" & glngCustomerNum & "," & l & ", '" & strTime(l) & "',"
            qryInsert = qryInsert & s(1, i) & "," & s(2, i) & "," & s(3, i) & "," & s(4, i) & "," & s(5, i) & "," & s(6, i) & "," & s(7, i) & "," & s(8, i) & "," & s(9, i) & "," & s(10, i) & ","
            qryInsert = qryInsert & s(11, i) & "," & s(12, i) & "," & s(13, i) & "," & s(14, i) & "," & s(15, i) & "," & s(16, i) & "," & s(17, i) & "," & s(18, i) & "," & s(19, i) & "," & s(20, i) & ","
            qryInsert = qryInsert & s(21, i) & "," & s(22, i) & "," & s(23, i) & "," & s(24, i) & "," & s(25, i) & "," & s(26, i) & "," & s(27, i) & "," & s(28, i) & "," & s(29, i) & "," & s(30, i) & ","
            qryInsert = qryInsert & s(31, i) & "," & s(32, i) & "," & s(33, i) & "," & s(34, i) & "," & s(35, i) & "," & s(36, i) & " )"

            Call modSql.AdoExcuteSql(qryInsert)
        Next i
        i = 0
        qryInsert = "INSERT INTO Nutrion (CustomerNum, bt,btname, m1, m2, m3, m4, m5,m6, m7, m8, m9, m10,m11, m12, m13, m14, m15,m16, m17, m18, m19, m20,m21, m22, m23, m24, m25,m26, m27, m28) "
        qryInsert = qryInsert & "VALUES ( " & glngCustomerNum & "," & 6 & ",'" & strTime(6) & "',"
        qryInsert = qryInsert & s(1, i) / sngDietCal * 100 & "," & s(2, i) / rValue2(2, 0) * 100 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & s(7, i) / rValue2(7, 0) * 100 & "," & s(8, i) / rValue2(8, 0) * 100 & "," & s(9, i) / rValue2(9, 0) * 100 & "," & 0 & ","
        qryInsert = qryInsert & 0 & "," & s(12, i) / rValue2(12, 0) * 100 & "," & s(13, i) / rValue2(13, 0) * 100 & "," & 0 & "," & 0 & "," & s(16, i) / rValue2(16, 0) * 100 & "," & s(17, i) / rValue2(17, 0) * 100 & "," & s(18, i) / rValue2(18, 0) * 100 & "," & s(19, i) / rValue2(19, 0) * 100 & "," & s(20, i) / rValue2(20, 0) * 100 & ","
        qryInsert = qryInsert & s(21, i) / rValue2(21, 0) * 100 & "," & s(22, i) / rValue2(22, 0) * 100 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & " )"
        Call modSql.AdoExcuteSql(qryInsert)

        qryInsert = "DELETE FROM NutrionGroup WHERE CustomerNum=" & glngCustomerNum
        Call modSql.AdoExcuteSql(qryInsert)

        For i = 0 To 4
            If i = 0 Then
               l = 5
            Else
                l = i
            End If
            qryInsert = "INSERT INTO NutrionGroup(CustomerNum,bt,btname, m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11, m12, m13, m14, m15, m16, m17, m18, m19) "
            qryInsert = qryInsert & "VALUES(" & glngCustomerNum & "," & l & ",'" & strTime(l) & "'," & c(1, i) & "," & c(2, i) & "," & c(3, i) & "," & c(4, i) & "," & c(5, i) & "," & c(6, i) & "," & c(7, i) & "," & c(8, i) & "," & c(9, i) & "," & c(10, i) & ","
            qryInsert = qryInsert & c(11, i) & "," & c(12, i) & "," & c(13, i) & "," & c(14, i) & "," & c(15, i) & "," & c(16, i) & "," & c(17, i) & "," & c(18, i) & "," & c(19, i) & ")"
            Call modSql.AdoExcuteSql(qryInsert)
        Next i

        Dim ss(4) As Single
        Dim ssr(4) As Single
        Dim ssc(4) As Single
        Dim fac(2 To 4) As Integer

        fac(2) = 4
        fac(3) = 9
        fac(4) = 4
        For i = 0 To 4
            ss(i) = 0
            ssr(i) = 0
            ssc(i) = 0
        Next i

        '각 영양소의 중량합을 구한다.(각 끼니별)
        For i = 2 To 4 '2=단백질, 3=지방 4=탄수화물
            ss(1) = ss(1) + s(i, 1) * fac(i) '아침
            ss(2) = ss(2) + s(i, 2) * fac(i) '점심
            ss(3) = ss(3) + s(i, 3) * fac(i) '저녁
            ss(4) = ss(4) + s(i, 4) * fac(i) '간식
        Next i
        For i = 1 To 4 '각 영양소
            If i = 1 Then '열량은 각 끼니 비율임
                For j = 1 To 4 '각 끼니

                     ssr(j) = Round(s(i, j) / s(i, 0) * 100, 2)
                Next j
            Else '영양소는 한끼니에서 각 영양소의 열량 비율임.
                For j = 1 To 4 '각 끼니
                    If ss(j) = 0 Then
                        ssr(j) = 0
                    Else
                        ssr(j) = Round(s(i, j) * fac(i) / ss(j) * 100, 2)
                    End If
                Next j
            End If

            qryInsert = "INSERT INTO NutrionCont(CustomerNum, element, m1, m2, m3, m4, m5,m6, m7, m8, m9) "
            qryInsert = qryInsert & "VALUES(" & glngCustomerNum & ",'" & strEelem(i) & "'," & s(i, 1) & "," & ssr(1) & "," & s(i, 2) & "," & ssr(2) & ","
            qryInsert = qryInsert & s(i, 3) & "," & ssr(3) & "," & s(i, 4) & "," & ssr(4) & "," & s(i, 0) & ")"
            Call modSql.AdoExcuteSql(qryInsert)
        Next i
        Calculate_Nut = True
    Else
        Calculate_Nut = False
    End If
End Function

'############################################################################
'
'  식습관평가 하단 그래프
'
'############################################################################
Private Sub CountHaveMeal()
'식사횟수를 체크하는 모듈
    Dim qrySelect As String, i As Integer
    Dim rValue As Variant
    Dim colCount As New Collection, colDate As New Collection
    Dim cfxArray As Object


On Error GoTo Err
    Set clsSelect = New clsSelect
    Set cfxArray = CreateObject("cfxdata.array")
    qrySelect = "SELECT Count(DietDiaryNum), MealDate "
    qrySelect = qrySelect & "FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum & " AND MealSection <> 4 AND NOT MealCalory IS NULL "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & "GROUP BY MealDate;"

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            colCount.Add Is_Null(rValue(0, i), 0)  '식사횟수
            If cmbPeriod.ListIndex = 0 Then
                colDate.Add Is_Null(rValue(1, i), "")
            Else
                colDate.Add Format(Is_Null(rValue(1, i), ""), "M/D")
            End If
        Next i
        cfxArray.AddArray colCount
        cfxArray.AddArray colDate

        ChartFX1.GetExternalData cfxArray
    Else
        ChartFX1.ClearData CD_VALUES
    End If

    Set colCount = Nothing
    Set colDate = Nothing
    Set clsSelect = Nothing
    Exit Sub
Err:
    '2004-12-23 류진선 로그기록
    'WriteLog "CountHaveMeal", "frmCounsel_5", Err.Number, Err.Description

End Sub

Private Sub TimeHaveMeal()
'각 끼니별 식사시간을 체크하는 모듈
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer

 On Error GoTo Err
    Set clsSelect = New clsSelect
    qrySelect = "SELECT MealDate, SUM(bf), SUM(lc), SUM(dn) FROM ( "
    qrySelect = qrySelect & "SELECT MealDate, MealTime AS bf, 0 AS lc, 0 AS dn FROM DietDiary WHERE CustomerNum=" & glngCustomerNum & " AND MealSection=1 "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & "UNION ALL "
    qrySelect = qrySelect & "SELECT MealDate, 0 AS bf, MealTime AS lc, 0 AS dn FROM DietDiary WHERE CustomerNum=" & glngCustomerNum & " AND MealSection=2 "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & "UNION ALL "
    qrySelect = qrySelect & "SELECT MealDate, 0 AS bf, 0 AS lc, MealTime AS dn FROM DietDiary WHERE CustomerNum=" & glngCustomerNum & " AND MealSection=3 "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & ") a GROUP BY MealDate;"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With chtTime
            .OpenDataEx COD_VALUES, 3, COD_UNKNOWN
            .Axis(AXIS_Y).Min = 0
            .Axis(AXIS_Y).Max = 24
            .Axis(AXIS_Y).STEP = 3
            .Series(0).Legend = "아침"
            .Series(1).Legend = "점심"
            .Series(2).Legend = "저녁"

            For i = 0 To UBound(rValue, 2)
                If rValue(1, i) < 1000 Then
                    .ValueEx(0, i) = Left(Is_Null(rValue(1, i), "0000"), 1)
                Else
                    .ValueEx(0, i) = Left(Is_Null(rValue(1, i), "0000"), 2)
                End If
                If rValue(2, i) < 1000 Then
                    .ValueEx(1, i) = Left(Is_Null(rValue(2, i), "0000"), 1)
                Else
                    .ValueEx(1, i) = Left(Is_Null(rValue(2, i), "0000"), 2)
                End If
                If rValue(3, i) < 1000 Then
                    .ValueEx(2, i) = Left(Is_Null(rValue(3, i), "0000"), 1)
                Else
                    .ValueEx(2, i) = Left(Is_Null(rValue(3, i), "0000"), 2)
                End If
                If cmbPeriod.ListIndex = 0 Then
                    .Axis(AXIS_X).Label(i) = Is_Null(rValue(0, i), "")
                Else
                    .Axis(AXIS_X).Label(i) = Format(Is_Null(rValue(0, i), ""), "M/D")
                End If
            Next i
            .CloseData COD_VALUES
        End With
    Else
        chtTime.ClearData CD_VALUES
    End If

    Set clsSelect = Nothing
    Exit Sub
Err:
    '2004-12-23 류진선 로그기록
    'WriteLog "TimeHaveMeal", "frmCounsel_5", Err.Number, Err.Description
    MsgBox Err.Number
End Sub

Private Sub SpeedHaveMeal()
'각 끼니마다 식사하는데 소요한 시간을 보여주는 모듈
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer

    Set clsSelect = New clsSelect
    qrySelect = "SELECT MealDate, SUM(bf), SUM(lc), SUM(dn) FROM ("
    qrySelect = qrySelect & "SELECT MealDate, MealNeedTime AS bf, 0 AS lc, 0 AS dn FROM DietDiary WHERE CustomerNum=" & glngCustomerNum & " AND MealSection=1 "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & "UNION ALL "
    qrySelect = qrySelect & "SELECT MealDate, 0 AS bf, MealNeedTime AS lc, 0 AS dn FROM DietDiary WHERE CustomerNum=" & glngCustomerNum & " AND MealSection=2 "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & "UNION ALL "
    qrySelect = qrySelect & "SELECT MealDate, 0 AS bf, 0 as lc, MealNeedTime AS dn FROM DietDiary WHERE CustomerNum=" & glngCustomerNum & " AND MealSection=3 "
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & ") a GROUP BY MealDate;"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With chtSpeed
            .OpenDataEx COD_VALUES, 3, COD_UNKNOWN
            .Axis(AXIS_Y).Min = 0
            .Axis(AXIS_Y).Max = 60
            .Axis(AXIS_Y).STEP = 10
            .Series(0).Legend = "아침"
            .Series(1).Legend = "점심"
            .Series(2).Legend = "저녁"

            For i = 0 To UBound(rValue, 2)
                .ValueEx(0, i) = CInt(Is_Null(rValue(1, i), 0))
                .ValueEx(1, i) = CInt(Is_Null(rValue(2, i), 0))
                .ValueEx(2, i) = CInt(Is_Null(rValue(3, i), 0))

                If cmbPeriod.ListIndex = 0 Then
                    .Axis(AXIS_X).Label(i) = Is_Null(rValue(0, i), "")
                Else
                    .Axis(AXIS_X).Label(i) = Format(Is_Null(rValue(0, i), ""), "M/D")
                End If
            Next i
            .CloseData COD_VALUES
        End With
    Else
        chtSpeed.ClearData CD_VALUES
    End If

    Set clsSelect = Nothing
End Sub

Private Sub EatingOut()
'외식횟수로 외식 정도를 보여주는 모듈
    Dim qrySelect As String, i As Integer
    Dim rValue As Variant
    Dim colCount As New Collection, colDate As New Collection
    Dim cfxArray As Object

    Set clsSelect = New clsSelect
    Set cfxArray = CreateObject("cfxdata.array")

    qrySelect = "SELECT COUNT(DietDiaryNum), MealDate "
    qrySelect = qrySelect & "FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND MealSection <> 4"  '간식제외
    qrySelect = qrySelect & " AND MealPlace = 3"     '장소가 외식인 것만
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY MealDate;"

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            colCount.Add Is_Null(rValue(0, i), 0)
            colDate.Add Is_Null(rValue(1, i), 0)
        Next i
        cfxArray.AddArray colCount
        cfxArray.AddArray colDate

        chtEatingOut.GetExternalData cfxArray
    Else
        chtEatingOut.ClearData CD_VALUES
    End If

    Set colCount = Nothing
    Set colDate = Nothing
    Set clsSelect = Nothing
End Sub

Private Sub WeekCalories()
    Dim qrySelect As String, rValue As Variant
    Dim sngCalory(7) As Single, nWeek As Integer, strWeek(7) As String
    Dim colWeek As New Collection
    Dim cfxArray As Object
    Dim i As Integer

    Set cfxArray = CreateObject("cfxdata.array")
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT MealDate, SUM(MealCalory) FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " GROUP BY MealDate"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            nWeek = Weekday(rValue(0, i), vbMonday) - 1
            sngCalory(nWeek) = sngCalory(nWeek) + CSng(rValue(1, i))
        Next i
        strWeek(0) = "월요일"
        strWeek(1) = "화요일"
        strWeek(2) = "수요일"
        strWeek(3) = "목요일"
        strWeek(4) = "금요일"
        strWeek(5) = "토요일"
        strWeek(6) = "일요일"
        cfxArray.AddArray sngCalory
        cfxArray.AddArray strWeek
        
        chtWeek.GetExternalData cfxArray
    Else
        chtWeek.ClearData CD_VALUES
    End If
    
    Set clsSelect = New clsSelect
End Sub

Private Sub InitialCountHaveMeal()
    With ChartFX1
        .ClearData CD_VALUES
        ' Chart Type Settings
        .Gallery = LINES
        .Chart3D = False
        .MarkerShape = MK_RECT

        ' Color Settings
        .BorderStyle = BORDER_FLAT
        .AxesStyle = CAS_FLATFRAME
        .Border = False
        .RGBBk = vbWhite

        ' Layout Settings
        .LegendBox = False
        .SerLegBox = False
        .ToolBar = False
        .Title(CHART_TOPTIT) = ""

        .PointLabels = True
        .Axis(AXIS_Y).Title = "횟수"
        .Axis(AXIS_Y).Min = 0
        .Axis(AXIS_Y).Max = 3
        .Axis(AXIS_Y).STEP = 1
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub InitialTimeHaveMeal()
    With chtTime
        .ClearData CD_VALUES
        ' Chart Type Settings
        .Gallery = LINES
        .Chart3D = False
    '    .MarkerShape = MK_NONE
        .MarkerShape = MK_RECT
        .RGBBk = vbWhite
        .Axis(0).Grid = True
        .Axis(2).Grid = True

        ' Color Settings
        .BorderStyle = BORDER_FLAT
        .AxesStyle = CAS_FLATFRAME
        .Border = False

        ' Layout Settings
        .SerLegBox = True
        .SerLegBoxObj.Docked = 515
        .ToolBar = False
        .Title(CHART_TOPTIT) = ""

        .Axis(AXIS_Y).Title = "시간:분"
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub InitialSpeedHaveMeal()
    With chtSpeed
        .ClearData CD_VALUES
        ' Chart Type Settings
        .Gallery = BAR
        .Chart3D = False
        .Stacked = CHART_NOSTACKED
        .Axis(0).Grid = True
        .RGBBk = vbWhite

        ' Color Settings
        .BorderStyle = BORDER_FLAT
        .Border = False

        ' Layout Settings
        .SerLegBox = True
        .SerLegBoxObj.Docked = 515
        .ToolBar = False
        .Title(CHART_TOPTIT) = ""

        .PointLabels = True
        .Axis(AXIS_Y).Title = "분"
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub InitialEatingOut()
    With chtEatingOut
        .ClearData CD_VALUES
        ' Chart Type Settings
        .Gallery = LINES
        .Chart3D = False
        .MarkerShape = MK_RECT
        .RGBBk = vbWhite

        ' Color Settings
        .BorderStyle = BORDER_FLAT
        .AxesStyle = CAS_FLATFRAME
        .Border = False

        ' Layout Settings
        .LegendBox = False
        .SerLegBox = False
        .ToolBar = False
        .Title(CHART_TOPTIT) = ""

        .PointLabels = True
        .Axis(AXIS_Y).Title = "횟수"
        .Axis(AXIS_Y).Min = 0
        .Axis(AXIS_Y).Max = 3
        .Axis(AXIS_Y).STEP = 1
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Sub InitialWeek()
    With chtWeek
        .ClearData CD_VALUES
        .Gallery = BAR
        .Chart3D = False
        .Stacked = CHART_NOSTACKED
        .Axis(0).Grid = True
        .Axis(AXIS_Y).Decimals = 0
        
        .MarkerShape = MK_RECT
        .RGBBk = vbWhite
        
        .BorderStyle = BORDER_FLAT
        .AxesStyle = CAS_FLATFRAME
        .Border = False
        
        .LegendBox = False
        .SerLegBox = False
        .ToolBar = False
        .Title(CHART_TOPTIT) = ""
        
        .PointLabels = True
        .Axis(AXIS_Y).Title = "열량kcal"
        .Axis(AXIS_X).Title = ""
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

'<< 식사일기 평가 >> 페이지를 출력하기 위해 준비하는 함수 /////////////////////////////////////////
Private Sub PrintData()
    Dim strConString As String
    Dim qrySelect As String, rValue As Variant
    Dim strBeginDay As String, strEndDay As String
    Dim i As Integer

On Error GoTo PrintErr
    '출력전에 선택된 기간내에 출력할 내용이 있는지를 먼저 확인할 것
    Set clsSelect = New clsSelect

    strBeginDay = Format(dtpBegin.Value, "YYYYMMDD")
    strEndDay = Format(dtpEnd.Value, "YYYYMMDD")

    qrySelect = "SELECT DISTINCT MealDate "
    qrySelect = qrySelect & "FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If

    rValue = clsSelect.Query(qrySelect)
    If IsNull(rValue) Then
        MsgBox "기간내에 입력된 식사일기가 없습니다.", vbInformation
        Set clsSelect = Nothing
        Exit Sub
    End If
    Set clsSelect = Nothing
    '리포터 연결 설정
    strServer = ServerName
'2005-01-18 류진선 DB정보수정
    strDBName = DBinfo.DBName
    strUID = DBinfo.DBID
    strPWD = DBinfo.DBPWD
'    strDBName = "Body"
'    strUID = "sa"
'    strPWD = "1111"

    Set crxReport = crxApplication.OpenReport(App.Path & "\Report\식사일기평가.rpt")
    crxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    crxReport.PaperOrientation = crPortrait
    With crxReport
        .RecordSelectionFormula = "{CustomerInfo.CustomerNum}=" & glngCustomerNum

        .Database.Tables(1).SetLogOnInfo strServer, strDBName, strUID, strPWD
'//////////////////////////////////////////  RDC 방식변경
        '+--------------------------------------------------
        '+ 1) 영양소 섭취상태
        '+--------------------------------------------------
        Call LoadCustomerInfo(glngCustomerNum)
        If typCustomerInfo.sngDietCal <> 0 Then
            With typCustomerInfo
                If Calculate_Nut(.sngDietCal, .intState, .intAge, .strSex) = True Then
                End If
            End With
        End If

        '1 : @Sex
        '2 : @FatPercent
        '3 : @Top5_Calory
        '4 : @Top5_Fat
        '5 : @Top5_SFA
        '6 : @Top5_Chol
        '7 : @Top5_Na
        '8 : @TreatCalory
        '9 : @Period
        '    - 섭취량, 가장 많이 포함된 다섯가지 음식
        '    - 열량 / 총지방량 / 포화지방 / 포화,불포화 / 콜레스테롤 / 나트륨
        .FormulaFields(3).Text = "'" & RPT_TopFood("열량") & "'"
        .FormulaFields(4).Text = "'" & RPT_TopFood("지방") & "'"
        .FormulaFields(5).Text = "'" & RPT_TopFood("포화지방") & "'"
        .FormulaFields(6).Text = "'" & RPT_TopFood("콜레스테롤") & "'"
        .FormulaFields(7).Text = "'" & RPT_TopFood("나트륨") & "'"
'        '    - 열량 권장량(선택된 기간내 처방된 칼로리들의 평균값)
        .FormulaFields(8).Text = "'" & Format(WhatTreatCalory, "#,###") & "'"
        '    - 선택된 기간 뿌려줌
        If cmbPeriod.ListIndex = 0 Then
            .FormulaFields(9).Text = "'" & Format(dtpBegin.Value, "YYYY.M.D") & "'"
        ElseIf cmbPeriod.ListIndex = 1 Then
            .FormulaFields(9).Text = "'" & Format(dtpBegin.Value, "YY.M.D") & " ~ " & Format(dtpEnd.Value, "YY.M.D") & "'"
        Else
            '초기 방문일부터 ~ ?
            Set clsSelect = New clsSelect

            qrySelect = "SELECT MIN(MealDate) FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum

            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue) Then
                .FormulaFields(9).Text = "'" & Format(rValue(0, 0), "YYYY.M.D") & " ~'"
            End If
        End If

        '+--------------------------------------------------
        '+ 2) 식습관
        '+--------------------------------------------------
        '    [1] 평균 일별 식사 횟수
        '10 : @Count
        Set clsSelect = New clsSelect

        qrySelect = "SELECT AVG(a) FROM ("
        qrySelect = qrySelect & "SELECT MealDate, COUNT(DietDiaryNum) AS a "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & "AND MealCalory IS NOT NULL"
        qrySelect = qrySelect & " GROUP BY MealDate) b;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            .FormulaFields(10).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(10).Text = "'0'"
        End If
        '    [2] 식사장소 / 시간
        '11 : @장소아침
        '12 : @장소점심
        '13 : @장소저녁
        '14 : @Time_B
        '15 : @Time_L
        '16 : @Time_D
        qrySelect = "SELECT MealSection, AVG(MealPlace), AVG(CAST(MealTime AS INT)) "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & " GROUP BY MealSection;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            For i = 0 To UBound(rValue, 2)
                Select Case rValue(0, i)
                    Case 1    ' 아침
                        .FormulaFields(11).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(14).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(14).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                    Case 2    ' 점심
                        .FormulaFields(12).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(15).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(15).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                    Case 3    ' 저녁
                        .FormulaFields(13).Text = "'" & Trim(rValue(1, i)) & "'"
                        If Len(Trim(rValue(2, i))) = 3 Then
                            .FormulaFields(16).Text = "'" & Mid(Trim(rValue(2, i)), 1, 1) & "'"
                        ElseIf Len(Trim(rValue(2, i))) = 4 Then
                            .FormulaFields(16).Text = "'" & Mid(Trim(rValue(2, i)), 1, 2) & "'"
                        End If
                End Select
            Next i
        End If
        '    [3] 걸린시간
        '17 : @NeedTime
        qrySelect = "SELECT AVG(MealNeedTime) FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            .FormulaFields(17).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(17).Text = "'0'"
        End If
        '    [4] 기분
        '18 : @Feeling
        qrySelect = "SELECT TOP 1 MealFeeling, COUNT(MealFeeling) AS a "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & " GROUP BY MealFeeling"
        qrySelect = qrySelect & " ORDER BY a DESC"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            .FormulaFields(18).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(18).Text = "'0'"
        End If
        '    [5] 식 후 배고픔 정도
        '19 : @Hungry
        qrySelect = "SELECT TOP 1 AfterMealHungry, COUNT(AfterMealHungry) AS a "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & " GROUP BY AfterMealHungry"
        qrySelect = qrySelect & " ORDER BY a DESC"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            .FormulaFields(19).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(19).Text = "'0'"
        End If
        '    [6] 외식횟수
        '20 : @EatOut
        qrySelect = "SELECT AVG(a) FROM ("
        qrySelect = qrySelect & "SELECT MealDate, COUNT(DietDiaryNum) AS a "
        qrySelect = qrySelect & "FROM DietDiary "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        If cmbPeriod.ListIndex = 0 Then
            qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
        ElseIf cmbPeriod.ListIndex = 1 Then
            qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
            qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
        End If
        qrySelect = qrySelect & "AND MealPlace=3 GROUP BY MealDate) b"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            .FormulaFields(20).Text = "'" & CInt(rValue(0, 0)) & "'"
        Else
            .FormulaFields(20).Text = "'0'"
        End If
    End With
    '+--------------------------------------------------
    '+ 두번째 장
    '+--------------------------------------------------
    Dim strTemp As String, strBeginDay1 As String, strEndDay1 As String
    Set crxReport2 = crxApplication.OpenReport(App.Path & "\Report\식사일기평가2.rpt")
    crxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    crxReport.PaperOrientation = crPortrait
    With crxReport2
        '1 : @GabCalory
        '2 : @GabMent
        '3 : @Rice
        '4 : @Exercise
        .Database.Tables(1).SetLogOnInfo strServer, strDBName, strUID, strPWD
        '    [1] 출력할 기간 선택
        If cmbPeriod.ListIndex = 1 Then
            strBeginDay1 = Left(strBeginDay, 4) & "," & Mid(strBeginDay, 5, 2) & "," & Right(strBeginDay, 2)
            strEndDay1 = Left(strEndDay, 4) & "," & Mid(strEndDay, 5, 2) & "," & Right(strEndDay, 2)
            strTemp = "{CustomerInfo.CustomerNum}=" & glngCustomerNum & " AND {DietDiary.MealDate} IN DateTime (" & strBeginDay1 & ") TO DateTime (" & strEndDay1 & ")"
        Else
            strTemp = "{CustomerInfo.CustomerNum}=" & glngCustomerNum
        End If
        .RecordSelectionFormula = strTemp

        '    [2] 하단에 요약정보
        Dim sngAvgTreatCal As Single, sngAvgMealCal As Single
        Dim sngAvgWeight As Single
        Dim sngGabCal As Single, sngRice As Single, intExercise As Integer
        '        - 해당기간내에 처방된 칼로리(Treat.TreatCalory)의 평균값
        sngAvgTreatCal = WhatTreatCalory
        If sngAvgTreatCal <> 0 Then
        '        - 해당기간내에 먹은 음식(일별)(DietDiary)의 평균값
            qrySelect = "SELECT AVG(a) FROM ("
            qrySelect = qrySelect & "SELECT MealDate, SUM(MealCalory) AS a "
            qrySelect = qrySelect & "FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            If cmbPeriod.ListIndex = 0 Then
                qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
            ElseIf cmbPeriod.ListIndex = 1 Then
                qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
                qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
            End If
            qrySelect = qrySelect & " GROUP BY MealDate) b"
            Set clsSelect = New clsSelect
            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue(0, 0)) Then
                sngAvgMealCal = CSng(rValue(0, 0))
                sngGabCal = sngAvgTreatCal - sngAvgMealCal
                .FormulaFields(1).Text = "'" & Format(Abs(sngGabCal), "#,###") & "'"
                If sngGabCal > 0 Then
                    .FormulaFields(2).Text = "'부족'"
                Else
                    .FormulaFields(2).Text = "'초과'"
                End If
                '    - 밥 한공기 300kcal
                sngRice = Abs(sngGabCal) / 300
                If sngRice >= 0.6 Then
                    .FormulaFields(3).Text = "'" & Format(sngRice, "#") & "'"
                ElseIf sngRice < 0.6 And sngRice >= 0.4 Then
                    .FormulaFields(3).Text = "'반'"
                Else
                    .FormulaFields(3).Text = "'0'"
                End If
                qrySelect = "SELECT AVG(a) FROM ("
                qrySelect = qrySelect & "SELECT TreatDay, SUM(Weight) AS a "
                qrySelect = qrySelect & "FROM BodyData LEFT JOIN Treat "
                qrySelect = qrySelect & "ON BodyData.TreatNum=Treat.TreatNum "
                qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
                If cmbPeriod.ListIndex = 0 Then
                    qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
                ElseIf cmbPeriod.ListIndex = 1 Then
                    qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
                    qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
                End If
                qrySelect = qrySelect & " GROUP BY TreatDay) b"

                rValue = clsSelect.Query(qrySelect)
                If Not IsNull(rValue(0, 0)) Then
                    sngAvgWeight = CSng(rValue(0, 0))
                    intExercise = sngGabCal / (sngAvgWeight * 0.16)
                    .FormulaFields(4).Text = "'" & intExercise & " 분'"
                Else  '기간내 입력된 체중이 없다면 가장 최근 체중
                    qrySelect = "SELECT TOP 1 Weight FROM BodyData LEFT JOIN Treat "
                    qrySelect = qrySelect & "ON Treat.TreatNum=BodyData.TreatNum "
                    qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
                    qrySelect = qrySelect & " ORDER BY TreatDay DESC;"

                    rValue = clsSelect.Query(qrySelect)
                    If Not IsNull(rValue(0, 0)) Then
                        sngAvgWeight = CSng(rValue(0, 0))
                        intExercise = sngGabCal / (sngAvgWeight * 0.16)
                        .FormulaFields(4).Text = "'" & intExercise & " 분'"
                    End If
                End If
                .PrintOut
            Else
                .FormulaFields(1).Text = "'0'"
                .FormulaFields(2).Text = "''"
                .FormulaFields(3).Text = "''"
                .FormulaFields(4).Text = "''"
            End If
        Else
            .FormulaFields(1).Text = "'0'"
            .FormulaFields(2).Text = "''"
            .FormulaFields(3).Text = "''"
            .FormulaFields(4).Text = "''"
        End If

    End With

    MsgBox "출력이 완료되었습니다.", vbOKOnly + vbInformation, "출력"
    Set clsSelect = Nothing

    Exit Sub
PrintErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "PrintData", "frmCounsel_5", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

Private Function WhatTreatCalory() As String
    Dim qrySelect As String, rValue As Variant

    Set clsSelect = New clsSelect
    qrySelect = "SELECT AVG(TreatCalory) FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue(0, 0)) Then
        WhatTreatCalory = rValue(0, 0)
    Else
        '만약 기간내에 입력된 처방칼로리가 없다면 가장 최근 것으로 사용한다.
        qrySelect = "SELECT TOP 1 TreatCalory FROM Treat "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & " ORDER BY TreatDay DESC;"

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue(0, 0)) Then
            WhatTreatCalory = rValue(0, 0)
        Else
            WhatTreatCalory = "0"   '이 단계까지 오면 안됨
        End If
    End If

    Set clsSelect = Nothing
End Function

Private Function RPT_TopFood(strNutrition) As String
    Dim qrySelect As String, rValue As Variant
    Dim intMealCal As Single, strFldNutrition As String
    Dim strTopFood As String
    Dim i As Integer

    Set clsSelect = New clsSelect
    Select Case strNutrition
        Case "열량"
            strFldNutrition = "tblFood.Energy"
        Case "지방"
            strFldNutrition = "tblFood.Fat"
        Case "포화지방"
            strFldNutrition = "tblFood.SFA"
        Case "콜레스테롤"
            strFldNutrition = "tblFood.Cholesterol"
        Case "나트륨"
            strFldNutrition = "tblFood.Na"
    End Select

    qrySelect = "SELECT DISTINCT(MealName),"
    qrySelect = qrySelect & "SUM((DietFood.FoodWeight*" & strFldNutrition & ")/100) AS a "
    qrySelect = qrySelect & "FROM DietDiary INNER JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "INNER JOIN tblMeal ON DietMeal.MealCode=tblMeal.MealID "
    qrySelect = qrySelect & "INNER JOIN DietFood ON DietMeal.DietMealNum=DietFood.DietMealNum "
    qrySelect = qrySelect & "INNER JOIN tblFood ON DietFood.FoodCode=tblFood.FoodID "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & "GROUP BY MealDate, MealSection, MealName "
    qrySelect = qrySelect & "ORDER BY a DESC;"

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        If UBound(rValue, 2) < 5 Then
            For i = 0 To UBound(rValue, 2) - 1
                strTopFood = ""
                strTopFood = strTopFood & Trim(rValue(0, i)) & " > "
                strTopFood = strTopFood & Trim(rValue(0, i))
            Next i
        Else
            strTopFood = Trim(rValue(0, 0)) & " > " & Trim(rValue(0, 1)) & " > " & Trim(rValue(0, 2))
            strTopFood = strTopFood & " > " & Trim(rValue(0, 3)) & " > " & Trim(rValue(0, 4))
        End If
    End If

    Set clsSelect = Nothing
    RPT_TopFood = strTopFood
End Function

Private Sub LoadCustomerInfo(lngCustomerNum As Long)
    Dim qrySelect As String, rValue As Variant

    Set clsSelect = New clsSelect
    qrySelect = "SELECT TOP 1 BodyData.BodyStatus, Age, Sex, Treat.TreatCalory "
    qrySelect = qrySelect & "FROM CustomerInfo INNER JOIN BodyData "
    qrySelect = qrySelect & "ON CustomerInfo.CustomerNum=BodyData.CustomerNum INNER JOIN "
    qrySelect = qrySelect & "Treat ON BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE CustomerInfo.CustomerNum=" & lngCustomerNum
    qrySelect = qrySelect & " ORDER BY Treat.TreatDay DESC;"
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With typCustomerInfo
            .intState = CInt(rValue(0, 0))
            .intAge = CInt(rValue(1, 0))
            .strSex = Trim(rValue(2, 0))
            .sngDietCal = Is_Null(rValue(3, 0), 0)
        End With
    End If

    Set clsSelect = Nothing
End Sub

'<3> 끼니별 비교/////////////////////////////////////////////////////////////////////////
Private Sub MealSectionRate()
    Dim qrySelect As String, rValue As Variant
    Dim intTotal As Integer, intBF As Integer, intLunch As Integer, intDinner As Integer, intSnack As Integer

On Error GoTo ShowErr
    Set clsSelect = New clsSelect

    qrySelect = "SELECT MealDate, SUM(a), SUM(b), SUM(c), SUM(d) FROM "
    qrySelect = "SELECT SUM(a), SUM(b), SUM(c), SUM(d) FROM "
    qrySelect = qrySelect & "(SELECT MealDate, SUM(Calories) AS a, 0 AS b, 0 AS c, 0 AS d "
    qrySelect = qrySelect & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " AND MealSection=1 GROUP BY MealDate, MealSection "
    qrySelect = qrySelect & "UNION ALL "
    qrySelect = qrySelect & "SELECT MealDate, 0 AS a, SUM(Calories) AS b, 0 AS c, 0 AS d "
    qrySelect = qrySelect & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " AND MealSection=2 GROUP BY MealDate, MealSection "
    qrySelect = qrySelect & "UNION ALL "
    qrySelect = qrySelect & "SELECT MealDate, 0 AS a, 0 AS b, SUM(Calories) AS c, 0 AS d "
    qrySelect = qrySelect & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " AND MealSection=3 GROUP BY MealDate, MealSection "
    qrySelect = qrySelect & "UNION ALL "
    qrySelect = qrySelect & "SELECT MealDate, 0 AS a, 0 AS b, 0 AS c, SUM(Calories) AS d "
    qrySelect = qrySelect & "FROM DietDiary LEFT JOIN DietMeal "
    qrySelect = qrySelect & "ON DietDiary.DietDiaryNum=DietMeal.DietDiaryNum "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    If cmbPeriod.ListIndex = 0 Then
        qrySelect = qrySelect & " AND MealDate='" & Format(cmbDaily.Text, "YYYYMMDD") & "' "
    ElseIf cmbPeriod.ListIndex = 1 Then
        qrySelect = qrySelect & " AND DietDiary.MealDate>='" & Format(dtpBegin.Value, "YYYYMMDD")
        qrySelect = qrySelect & "' AND DietDiary.MealDate<='" & Format(dtpEnd.Value, "YYYYMMDD") & "'"
    End If
    qrySelect = qrySelect & " AND MealSection=4 GROUP BY MealDate, MealSection) AS park "

    rValue = clsSelect.Query(qrySelect)
    
    If Not IsNull(rValue) Then
        intBF = Is_Null(rValue(0, 0), 0)
        intLunch = Is_Null(rValue(1, 0), 0)
        intDinner = Is_Null(rValue(2, 0), 0)
        intSnack = Is_Null(rValue(3, 0), 0)
        intTotal = intBF + intLunch + intDinner + intSnack
    End If

    If intTotal = 0 Then
        Exit Sub
    End If
    Set clsSelect = Nothing
    Exit Sub
ShowErr:
    '2004-12-23 류진선 로그기록
    'WriteLog "MealSectionRate", "frmCounsel_5", Err.Number, Err.Description
    MsgBox Err.Description
End Sub

Private Sub imgStart_Click()
    Dim datTemp As Date
    If cmbPeriod.ListIndex = 1 Then
        If dtpBegin.Value > dtpEnd.Value Then
            datTemp = dtpBegin.Value
            dtpBegin.Value = dtpEnd.Value
            dtpEnd.Value = datTemp
        End If
    End If
    Call ShowVal
End Sub

Private Sub imgStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgStart.Picture = LoadPicture(App.Path & "\Back\Counsel\on.jpg")
End Sub

Private Sub imgStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgStart.Picture = LoadPicture(App.Path & "\Back\Counsel\off.jpg")
End Sub

Private Sub imgSub_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 4
        imgSub(i).Enabled = True
    Next i
    Select Case Index
        Case 0:  '식사횟수
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB1_ON)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB5_OFF)
            
            chtTime.Visible = False
            chtSpeed.Visible = False
            chtEatingOut.Visible = False
            chtWeek.Visible = False
            ChartFX1.Visible = True
        Case 1:  '식사시간
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB2_ON)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB5_OFF)
            
            ChartFX1.Visible = False
            chtSpeed.Visible = False
            chtEatingOut.Visible = False
            chtWeek.Visible = False
            chtTime.Visible = True
        Case 2:  '식사속도
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB3_ON)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB5_OFF)
            
            ChartFX1.Visible = False
            chtTime.Visible = False
            chtEatingOut.Visible = False
            chtWeek.Visible = False
            chtSpeed.Visible = True
        Case 3:  '외식횟수
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB4_ON)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB5_OFF)
            
            ChartFX1.Visible = False
            chtTime.Visible = False
            chtSpeed.Visible = False
            chtWeek.Visible = False
            chtEatingOut.Visible = True
        Case 4:  '요일분석
            Set imgSub(0).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB1_OFF)
            Set imgSub(1).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB2_OFF)
            Set imgSub(2).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB3_OFF)
            Set imgSub(3).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB4_OFF)
            Set imgSub(4).Picture = LoadPicture(App.Path & PATH05 & IMG_SUB5_ON)
            
            ChartFX1.Visible = False
            chtTime.Visible = False
            chtSpeed.Visible = False
            chtEatingOut.Visible = False
            chtWeek.Visible = True
    End Select
End Sub

Private Function ExistDiary() As Boolean
'해당환자에 입력한 식사일기가 있는지 체크
    Dim qrySelect As String, rValue As Variant
    
    qrySelect = "SELECT DietDiaryNum FROM DietDiary "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        ExistDiary = True
    Else
        ExistDiary = False
    End If
End Function

