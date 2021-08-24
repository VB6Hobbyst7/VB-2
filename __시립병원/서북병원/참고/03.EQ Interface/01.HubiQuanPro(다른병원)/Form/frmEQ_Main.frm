VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmEQ_Main 
   Caption         =   "Hi Interface EQ"
   ClientHeight    =   10095
   ClientLeft      =   1110
   ClientTop       =   5790
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   15000
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6300
      TabIndex        =   38
      Top             =   60
      Width           =   2235
   End
   Begin VB.TextBox txtBuff 
      Height          =   1755
      Left            =   5820
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   6420
      Width           =   6195
   End
   Begin VB.TextBox txtSerialData 
      Height          =   5175
      Left            =   15180
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Top             =   4500
      Width           =   4995
   End
   Begin FPSpread.vaSpread sprDResult 
      Height          =   3075
      Left            =   5700
      TabIndex        =   7
      Top             =   1020
      Width           =   9255
      _Version        =   393216
      _ExtentX        =   16325
      _ExtentY        =   5424
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      MaxRows         =   10
      SpreadDesigner  =   "frmEQ_Main.frx":263A
      UserResize      =   1
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "화면정리(&C)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12660
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13980
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9660
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin MSComctlLib.ProgressBar prgPatient 
      Height          =   75
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar staCondition 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9720
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5398
            MinWidth        =   3528
            Text            =   "Copyright ⓒ 2010 Medimate Corp."
            TextSave        =   "Copyright ⓒ 2010 Medimate Corp."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10954
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "frmEQ_Main.frx":2D72
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1058
            Text            =   "Local DB"
            TextSave        =   "Local DB"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "HIS DB"
            TextSave        =   "HIS DB"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "COM"
            TextSave        =   "COM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "2011-10-16"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "오후 7:14"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpread.vaSpread sprLResult 
      Height          =   5175
      Left            =   60
      TabIndex        =   4
      Top             =   4500
      Width           =   14895
      _Version        =   393216
      _ExtentX        =   26273
      _ExtentY        =   9128
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   20
      OperationMode   =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmEQ_Main.frx":32E9
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   5640
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Sample No"
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
      Index           =   17
      Left            =   120
      TabIndex        =   40
      Top             =   1500
      Width           =   945
   End
   Begin VB.Label lblSAMPLENO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   39
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label lblDISKNOPOSNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   36
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Rack/Pos"
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
      Index           =   16
      Left            =   120
      TabIndex        =   35
      Top             =   1740
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사회차"
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
      Index           =   20
      Left            =   4260
      TabIndex        =   34
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label lblEXSEQ 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5160
      TabIndex        =   33
      Top             =   1080
      Width           =   120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  '투명하지 않음
      Height          =   255
      Index           =   5
      Left            =   9240
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Delta"
      Height          =   180
      Index           =   15
      Left            =   9540
      TabIndex        =   32
      Top             =   4260
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Low"
      Height          =   180
      Index           =   14
      Left            =   7920
      TabIndex        =   31
      Top             =   4260
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  '투명하지 않음
      Height          =   255
      Index           =   4
      Left            =   7620
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Panic"
      Height          =   180
      Index           =   13
      Left            =   10320
      TabIndex        =   30
      Top             =   4260
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '투명하지 않음
      Height          =   255
      Index           =   2
      Left            =   10020
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "High"
      Height          =   180
      Index           =   12
      Left            =   8700
      TabIndex        =   29
      Top             =   4260
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  '투명하지 않음
      Height          =   255
      Index           =   1
      Left            =   8400
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "결과 전송"
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
      Index           =   11
      Left            =   120
      TabIndex        =   28
      Top             =   2460
      Width           =   885
   End
   Begin VB.Label lblSDDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   27
      Top             =   2460
      Width           =   900
   End
   Begin VB.Shape shpCon 
      BackStyle       =   1  '투명하지 않음
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   1
      Left            =   4800
      Top             =   180
      Width           =   135
   End
   Begin VB.Shape shpCon 
      BackStyle       =   1  '투명하지 않음
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   135
   End
   Begin VB.Label lblRCDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   26
      Top             =   2220
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "결과 수신"
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
      Index           =   10
      Left            =   120
      TabIndex        =   25
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label lblORDGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   24
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label lblORDDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   23
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label lblSEXAGE 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   22
      Top             =   3900
      Width           =   900
   End
   Begin VB.Label lblPATNM 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   21
      Top             =   3660
      Width           =   900
   End
   Begin VB.Label lblPATNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   20
      Top             =   3420
      Width           =   900
   End
   Begin VB.Label lblEXDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   19
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "실시간 검사리스트"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   180
      TabIndex        =   18
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방 종류"
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
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3180
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방 일자"
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
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2940
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "처방 전송"
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
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1980
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "성별/연령"
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
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   3900
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "수검자 명"
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
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   3660
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "병록 번호"
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
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   3420
      Width           =   885
   End
   Begin VB.Label lblBARCD 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1260
      TabIndex        =   11
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검체 번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검체번호별 세부정보"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검체번호별 검사결과"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5820
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lbl장비명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사장비명"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   300
      TabIndex        =   1
      Top             =   60
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '단색
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   4875
   End
   Begin VB.Shape shpDResult 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   255
      Left            =   5700
      Shape           =   4  '둥근 사각형
      Top             =   720
      Width           =   9255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   255
      Index           =   0
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   720
      Width           =   5595
   End
   Begin VB.Shape shpLResult 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '하향 대각선
      Height          =   255
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   4200
      Width           =   6915
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File    "
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "환경설정    "
      Begin VB.Menu mnuSettingSub 
         Caption         =   "통신설정"
         Index           =   0
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "HIS DB 접속정보"
         Index           =   1
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "ETC DB 접속정보"
         Index           =   2
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "통신신호 추적"
         Index           =   3
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuJob 
      Caption         =   "작업    "
      Begin VB.Menu mnuJobSub 
         Caption         =   "WorkList 작업"
         Index           =   0
      End
      Begin VB.Menu mnuJobSub 
         Caption         =   "검사결과 관리"
         Index           =   1
      End
      Begin VB.Menu mnuJobSub 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuJobSub 
         Caption         =   "전송방식"
         Index           =   4
         Begin VB.Menu mnuJobModeAuto 
            Caption         =   "자동전송[Auto]"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuJobModeManual 
            Caption         =   "수동전송[Manual]"
         End
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "기초코드    "
      Begin VB.Menu mnuCodeSub 
         Caption         =   "장비검사코드 관리"
         Index           =   0
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "정보"
   End
End
Attribute VB_Name = "frmEQ_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height의 초기값
Dim lngMeWidth      As Long '/Me.Width의 초기값
Dim strOneLine      As String


Private Type ConWhere   ' 사용자 정의 형식을 만듭니다.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWhere

Public Function CHK_COMM_PORT() As Boolean
    CHK_COMM_PORT = False
    
On Error GoTo RTN_ERR_PORT

RE_CHK:

    MSComm1.CommPort = gtypEQ_INFO.SERIALPORT
    MSComm1.RTSEnable = gtypEQ_INFO.SERIALRTS
    MSComm1.DTREnable = gtypEQ_INFO.SERIALDTR
    MSComm1.Settings = gtypEQ_INFO.SERIALBAUD & "," & gtypEQ_INFO.SERIALPARITY & "," & gtypEQ_INFO.SERIALDATABIT & "," & gtypEQ_INFO.SERIALSTOPBIT

    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
    
    CHK_COMM_PORT = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR_PORT:
    If Err = 8002 Then      'Port
        If MsgBox("▶▶▶통신설정 Info" & vbCrLf & vbCrLf & _
                  "MSComm Port Setting 이 올바르지 않습니다." & vbCrLf & _
                  "(재)설정하겠습니까?", vbQuestion + vbYesNo, "질의") = vbNo Then
            
            MsgBox "▶▶▶통신설정 Info" & vbCrLf & vbCrLf & _
                   "계속 진행할 경우 일부 기능이 제한됩니다." & vbCrLf & _
                   "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
                
            Exit Function
        Else
            frmEQ공용_Set_Port.Show vbModal
            
            GoTo RE_CHK
        End If
    Else
        Resume Next
    End If
End Function

Public Function FUNC_LOC_VIEW(ArgSection As Integer) As Boolean
    Dim str처방코드     As String
    
    FUNC_LOC_VIEW = False
    
On Error GoTo RTN_ERR
    
    
    
'''    If ConnDB_LOC(gstrREG_DB_CONSTR) = True Then
'''        '/장비코드별 처방코드 가져오기
'''        gstrQuy = "SELECT ORDCD "
'''        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_EQORD "
'''        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
'''        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
'''
'''        If Not ADR_LOC Is Nothing Then
'''            Do Until ADR_LOC.EOF
'''                str처방코드 = str처방코드 & ",'" & Trim(ADR_LOC!ORDCD & "") & "'"
'''
'''                ADR_LOC.MoveNext
'''            Loop
'''            ADR_LOC.Close: Set ADR_LOC = Nothing
'''
'''            str처방코드 = Mid(str처방코드, 2)
'''        End If
    
    
    FUNC_LOC_VIEW = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Sub SUB_COMM_PART_HUBIQUANPRO_BAR(argCOMM_BF As String)
    '/Patient ID 가 바코드일 경우로 처리한다.
    
    '/일반검사 Sample
    'H|humasis|HUBI-QUAN pro|HP90003|46
    'P|110928-0001|20110928091712|P|CARDIAC 3/1|10-006
    'R1|CK-MB|0.00~5.00|MYO|0.00~100.00|TNI|0.00~0.40|
    'R2|CK-MB|>30.00|ng/mL| |
    'R2|MYO|>150.00|ng/mL| |
    'R2|TNI|7.00|ng/mL| |
    'L|1|N
    
    '/QC Sample
    'H|humasis|HUBI-QUAN pro|HP90003|
    'P|110928-0002|20110928095439|CARDIAC 3/1|10-006
    'R|CK-MB|ng/mL|10.53|Low|29.24|High
    'R|MYO|ng/mL|68.66|Low|>150.00|High
    'R|TNI|ng/mL|2.39|Low|6.87|High
    'L|1|N
    
    Dim strCrlf '/배열변수(CRLF 기준)
    Dim strPart '/배열변수(| 기준)
    Dim strLotNo '/Lot No
    Dim strEXSEQ            As String '/EXSEQ 값 증가여부
    Dim intLResultRow       As Integer
    Dim intLResultTarRow    As Integer
    Dim intLResultTarCol    As Integer
    Dim intCol              As Integer
    Dim intLineCnt          As Integer
    
    strCrlf = Split(argCOMM_BF, vbCrLf)
    
    For intLineCnt = 0 To UBound(strCrlf) - 1
        strPart = Split(strCrlf(intLineCnt), "|")
        
        Select Case strPart(0)
            Case "H" '/Hearder정보(결과신호시작)
                '/처리안함
            
            Case "P" '/Patient정보(환자정보)
                '/strPart(0): P
                '/strPart(1): Test No(Barcode) YYDDMM-XXX1 날짜별 SEQNO
                '/strPart(2): 날짜시간(검사시작일시?, 결과전송일시?) YYYYMMDDHHMMSS
                '/strPart(3): Patient ID (입력되지 않으면 P 로 표시됨)
                '/strPart(4): Device Name(코드를 읽고 제품명이 출력된다)
                '/strPart(5): Lot No
                
                gtypPAT_RES.SAMPLENO = Trim(strPart(1)) '/Sample No(Test No)
                
                '/Patient ID 가 바코드일 경우
                gtypPAT_RES.BARCD = Trim(strPart(3)) '/BARCD(검체번호(Barcode))
                '/Patient ID 가 바코드일 경우
                
                gtypPAT_RES.EXDT = Trim(Left(strPart(2), 8)) '/EXDT(검사처방전송일자(YYYYMMDD) HIEQ->의료장비)
                gtypPAT_RES.EXTM = Trim(Mid(strPart(2), 9)) '/EXTM(검사처방전송시간(24HHMMSS) HIEQ->의료장비)
                
                strLotNo = Split(strPart(5), "-")
                gtypPAT_RES.DISKNO = strLotNo(0) '/DISKNO(LotNo 의 앞부분)
                gtypPAT_RES.POSNO = strLotNo(1) '/POSNO(LotNo 의 뒷부분)
                
                Call FUNC_HIS_PATIENT '/HIS 환자정보 가져오기
                
                '/검체번호/SampleNo/Rack/Pos가 정의된 상태에서 진행할 것.
                If strEXSEQ <> "Y" Then
                    gtypPAT_RES.EXSEQ = FUNC_GET_EXSEQ(gtypPAT_RES.BARCD) '/검체번호(Barcode)별 검사회차
                    strEXSEQ = "Y"
                End If
                
                '/실시간 검사리스트 보여주기----------------------------------------------------------------------------------------------------/
                intLResultTarRow = 0
                For intLResultRow = 1 To sprLResult.DataRowCnt
                    If Trim(GET_CELL(sprLResult, 1, intLResultRow)) = gtypPAT_RES.BARCD And _
                       Trim(GET_CELL(sprLResult, 2, intLResultRow)) = gtypPAT_RES.EXSEQ And _
                       Trim(GET_CELL(sprLResult, 3, intLResultRow)) = gtypPAT_RES.SAMPLENO And _
                       Trim(GET_CELL(sprLResult, 4, intLResultRow)) = gtypPAT_RES.DISKNO And _
                       Trim(GET_CELL(sprLResult, 5, intLResultRow)) = gtypPAT_RES.POSNO Then
                       
                        intLResultTarRow = intLResultRow
                        Exit For
                    End If
                Next intLResultRow
            
                If intLResultTarRow = 0 Then
                    sprLResult.MaxRows = sprLResult.MaxRows + 1
                    intLResultTarRow = sprLResult.MaxRows
                    
                    Call SET_CELL(sprLResult, 1, intLResultTarRow, gtypPAT_RES.BARCD)
                    Call SET_CELL(sprLResult, 2, intLResultTarRow, gtypPAT_RES.EXSEQ)
                    Call SET_CELL(sprLResult, 3, intLResultTarRow, gtypPAT_RES.SAMPLENO)
                    Call SET_CELL(sprLResult, 4, intLResultTarRow, gtypPAT_RES.DISKNO)
                    Call SET_CELL(sprLResult, 5, intLResultTarRow, gtypPAT_RES.POSNO)
                    Call SET_CELL(sprLResult, 8, intLResultTarRow, IIf(gtypPAT_RES.EXDT <> "", Format(gtypPAT_RES.EXDT, "@@@@-@@-@@"), "") & IIf(gtypPAT_RES.EXTM <> "", " " & Format(gtypPAT_RES.EXTM, "@@:@@:@@"), ""))
                    Call SET_CELL(sprLResult, 13, intLResultTarRow, gtypPAT_RES.PATNO)
                    Call SET_CELL(sprLResult, 14, intLResultTarRow, gtypPAT_RES.PATNM)
                    If gtypPAT_RES.PATSEX <> "" Or gtypPAT_RES.PATAGE <> "" Then
                        Call SET_CELL(sprLResult, 15, intLResultTarRow, gtypPAT_RES.PATSEX & "/" & gtypPAT_RES.PATAGE)
                    End If
                End If
                '/실시간 검사리스트 보여주기----------------------------------------------------------------------------------------------------/
                
            Case "R" '/검사결과정보(QC)
                
            Case "R1" '/참고범위정보(일반검사)
                '/처리안함
            
            Case "R2" '/검사결과정보(일반검사)
                '/strPart(0): R2
                '/strPart(1): 장비검사코드
                '/strPart(2): 장비검사결과
                '/strPart(3): 결과단위
                
                'R2|CK-MB|>30.00|ng/mL| |
            
                gtypPAT_RES.EQCD = strPart(1) '/EQCD(장비검사코드)
                
                gtypPAT_RES.EQRESULT = strPart(2) '/EQRESULT(장비원시결과)
                gtypPAT_RES.Result = FUNC_RESULT_CHANGE(gtypPAT_RES.EQCD, gtypPAT_RES.EQRESULT) '/RESULT(검사결과(변형된 결과))
                gtypPAT_RES.RCDT = Format(Now, "YYYYMMDD") '/RCDT(검사결과수신일자(YYYYMMDD) 의료장비 ->HIEQ)
                gtypPAT_RES.RCTM = Format(Now, "HHMMSS") '/RCTM(검사결과수신시간(24HHMMSS) 의료장비 ->HIEQ)
                gtypPAT_RES.STATEFLAG = "1" '/STATEFLAG(결과진행상태 (0:처방, 1:결과))
                
                Call FUNC_HIS_ORDER_VIEW    '/처방내역 조회
                Call FUNC_HIS_RESULT_JUDGMENT   '/결과 판정
                
                gtypPAT_RES.SENDFLAG = "0"
                    
                '/실시간 검사리스트 보여주기----------------------------------------------------------------------------------------------------/
                Call SET_CELL(sprLResult, 9, intLResultTarRow, IIf(gtypPAT_RES.RCDT <> "", Format(gtypPAT_RES.RCDT, "@@@@-@@-@@"), "") & IIf(gtypPAT_RES.RCTM <> "", " " & Format(gtypPAT_RES.RCTM, "@@:@@:@@"), "")) '/RCDT(검사결과수신일자(YYYYMMDD) 의료장비 ->HIEQ)
                Call SET_CELL(sprLResult, 6, intLResultTarRow, IIf(gtypPAT_RES.STATEFLAG = "1", "결과", "처방"))
                Call SET_CELL(sprLResult, 6, intLResultTarRow, IIf(gtypPAT_RES.SENDFLAG = "1", "완료", "대기"))
                Call SET_CELL(sprLResult, 11, intLResultTarRow, gtypPAT_RES.ORDDT) '/ORDDT(처방일자)
                Select Case gtypPAT_RES.ORDGB '/ORDGB(처방종류(O.외래, I.입원, G.건강검진))
                    Case "O": Call SET_CELL(sprLResult, 12, intLResultTarRow, "외래")
                    Case "I": Call SET_CELL(sprLResult, 12, intLResultTarRow, "입원")
                    Case "G": Call SET_CELL(sprLResult, 12, intLResultTarRow, "검진")
                End Select
                
                '/장비검사항목 Column Set/결과
                For intCol = gintEQ_StartCol To sprLResult.MaxCols
                    If GET_CELL(sprLResult, intCol, -1000) = gtypPAT_RES.EQCD Then
                        Call SET_CELL(sprLResult, intCol, intLResultTarRow, gtypPAT_RES.Result) '/RESULT(검사결과(변형된 결과))
                        
                        sprLResult.Col = intCol
                        sprLResult.Row = intLResultTarRow
                        
                        If gtypPAT_RES.AFLAG = "L" Then
                            sprLResult.BackColor = &HFFFF&
                        End If
                        If gtypPAT_RES.AFLAG = "H" Then
                            sprLResult.BackColor = &HFFFF80
                        End If
                        If gtypPAT_RES.DFLAG = "D" Then
                            sprLResult.BackColor = &HFF8080
                        End If
                        If gtypPAT_RES.PFLAG = "P" Then
                            sprLResult.BackColor = &HFF&
                        End If

                        Exit For
                    End If
                Next intCol
                '/실시간 검사리스트 보여주기----------------------------------------------------------------------------------------------------/
                
                '/해당 Row로 Focus 이동
                sprLResult.Col = 1
                sprLResult.Row = intLResultTarRow
                sprLResult.Action = ActionActiveCell
                
                '/해당 Row 자료 세부정보와 검사결과 표시
                Call sprLResult_LeaveRow(intLResultTarRow - 1, False, False, False, intLResultTarRow, False, False)
                
                '/받은 자료 Local 저장
                If FUNC_LOC_SAVE_PAT_RES = True Then
                    If mnuJobModeAuto.Checked = True Then '/전송방식이 자동전송이면...
                        Call FUNC_HIS_SAVE '/HIS에 결과 전송
                        Call FUNC_LOC_SAVE_SEND(gtypPAT_RES.BARCD, gtypPAT_RES.EXSEQ, gtypPAT_RES.EQCD, gtypPAT_RES.SAMPLENO, gtypPAT_RES.DISKNO, gtypPAT_RES.POSNO, "1") '/HIS에 결과 전송
                    End If
                End If
            Case "L" '/종료정보
                '/처리안함
    
                gtypPAT_RES.BARCD = ""
                gtypPAT_RES.EXSEQ = ""            '/EXSEQ(검체번호(Barcode)별 검사회차)
                gtypPAT_RES.EQCD = ""             '/EQCD(장비검사코드)
                gtypPAT_RES.EXAMCD = ""           '/EXAMCD(처방코드(HIS or LIS의 검사코드))
                gtypPAT_RES.EXDT = ""             '/EXDT(검사처방전송일자(YYYYMMDD) HIEQ->의료장비)
                gtypPAT_RES.EXTM = ""             '/EXTM(검사처방전송시간(24HHMMSS) HIEQ->의료장비)
                gtypPAT_RES.RCDT = ""             '/RCDT(검사결과수신일자(YYYYMMDD) 의료장비 ->HIEQ)
                gtypPAT_RES.RCTM = ""             '/RCTM(검사결과수신시간(24HHMMSS) 의료장비 ->HIEQ)
                gtypPAT_RES.SDDT = ""             '/SDDT(검사결과전송일자(YYYYMMDD) HIEQ->HIS)
                gtypPAT_RES.SDTM = ""             '/SDTM(검사결과전송시간(24HHMMSS) HIEQ->HIS)
                gtypPAT_RES.Result = ""           '/RESULT(검사결과(변형된 결과))
                gtypPAT_RES.EQRESULT = ""         '/EQRESULT(장비원시결과)
                gtypPAT_RES.AFLAG = ""            '/AFLAG(Abnormal(정상참고치 기준 (H)High or (L)Low 값 표시))
                gtypPAT_RES.PFLAG = ""            '/PFLAG(Panic)
                gtypPAT_RES.DFLAG = ""            '/DFLAG(Delta)
                gtypPAT_RES.SAMPLENO = ""         '/Sample No(AU2700, Uriscan 등에 사용)
                gtypPAT_RES.DISKNO = ""           '/DISKNO(디스크번호 or 렉번호)
                gtypPAT_RES.POSNO = ""            '/POSNO(위치번호)
                gtypPAT_RES.ORDDT = ""            '/ORDDT(처방일자)
                gtypPAT_RES.ORDGB = ""            '/ORDGB(처방종류(O.외래, I.입원, G.건강검진))
                gtypPAT_RES.PATNO = ""            '/PATNO(병록번호)
                gtypPAT_RES.PATNM = ""            '/PATNM(수검자명)
                gtypPAT_RES.PATSEX = ""           '/PATSEX(성별)
                gtypPAT_RES.PATAGE = ""           '/PATAGE(연령)
                gtypPAT_RES.SENDFLAG = ""         '/SENDFLAG(HIS 전송 FLAG (0:대기, 1:완료))
                gtypPAT_RES.STATEFLAG = ""        '/STATEFLAG(결과진행상태 (0:처방, 1:결과))
        End Select
    Next intLineCnt
End Sub

Public Sub SUB_MM_CANCEL()
    lbl장비명 = ""
    prgPatient.Max = 100
    prgPatient.Value = 100
    
    Call SUB_MM_KEY_CLEAR("1") '/검체번호별 세부정보
    Call SUB_MM_KEY_CLEAR("2") '/검체번호별 검사결과
    Call SUB_MM_KEY_CLEAR("3") '/실시간 검사리스트
    
    mnuSetting.Visible = False '/구성메뉴 안보이기

    txtSerialData.Visible = False
End Sub

Public Sub SUB_MM_INITIAL()
    '/Resize를 위한 초기 Size Setting----------------------------------------------------------------------------------------------------/
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Timer
            Case TypeOf Me.Controls(intX) Is Menu
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is MSComm
            Case TypeOf Me.Controls(intX) Is CommonDialog
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Size Setting
    lngMeHeight = 10890
    lngMeWidth = 15150
    
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Show
    '/Resize를 위한 초기 Size Setting----------------------------------------------------------------------------------------------------/
    
    '/초기 자료 Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/초기 자료 Setting----------------------------------------------------------------------------------------------------/
    
    '/변동 컨트롤 초기화----------------------------------------------------------------------------------------------------/
    Call SUB_MM_CANCEL
    '/변동 컨트롤 초기화----------------------------------------------------------------------------------------------------/
    
    '/워크리스트 작업여부(Y.사용함, N.사용안함)
    If gtypEQ_INFO.WORKLISTGB = "Y" Then
        mnuJobSub(0).Visible = True
    Else
        mnuJobSub(0).Visible = False
    End If
    
    '/작업모드(A.자동, M.수동)
    If gtypEQ_INFO.AUTOGB = "Y" Then
        mnuJobModeAuto.Checked = True
        staCondition.Panels.Item(3).Picture = LoadPicture(App.Path & "\Auto.jpg")
        mnuJobModeManual.Checked = False
    Else
        mnuJobModeManual.Checked = True
        staCondition.Panels.Item(3).Picture = LoadPicture(App.Path & "\Manual.jpg")
        mnuJobModeAuto.Checked = False
    End If
    
    Me.Caption = Me.Caption & " For " & App.FileDescription
    Me.Caption = Me.Caption & Space(10) & "(사용자: " & gtypUSER.USERNM & " )"
    
    lbl장비명 = App.FileDescription

    '/작업 상태 Check----------------------------------------------------------------------------------------------------/
    If ConnDB_HIS = True Then
        Call CloseDB_HIS
        staCondition.Panels.Item(5).Enabled = True '/HISDB 활성화
    Else
        staCondition.Panels.Item(5).Enabled = False '/HISDB 비활성화
    End If
    
    If CHK_COMM_PORT = True Then
        staCondition.Panels.Item(6).Enabled = True '/COM Port 활성화
    Else
        staCondition.Panels.Item(6).Enabled = False '/COM Port 비활성화
    End If
    '/작업 상태 Check----------------------------------------------------------------------------------------------------/
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ADD_ITEM:
    Dim intCnt   As Integer
    
    '/검체번호별 검사결과 Title Clear
    '/검사명|장비검사코드|결과|Wall
    sprDResult.ClearRange 1, -1, 2, -1, True
    sprDResult.ClearRange 4, -1, 5, -1, True
    sprDResult.ClearRange 7, -1, 8, -1, True
    sprDResult.ClearRange 10, -1, 11, -1, True
    
    If sprLResult.MaxCols > gintEQ_StartCol - 1 Then sprLResult.MaxCols = gintEQ_StartCol - 1
    
    If ConnDB_LOC = False Then
        MsgBox "▶▶▶Local DataBase Info" & vbCrLf & vbCrLf & _
               "Local DataBase 를 접속할 수 없습니다." & vbCrLf & _
               "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
        End
    Else
        frmEQ_Main.staCondition.Panels.Item(4).Enabled = True '/Main 화면의 Local 상태 활성화

        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY EQSEQ "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then End
        
        If Not ADR_LOC Is Nothing Then
            Do Until ADR_LOC.EOF
                sprLResult.MaxCols = sprLResult.MaxCols + 1
                sprLResult.Col = sprLResult.MaxCols
                sprLResult.Row = -1
                sprLResult.BackColor = RGB(255, 255, 230)
                sprLResult.CellType = CellTypeStaticText
                sprLResult.TypeHAlign = TypeHAlignRight
                sprLResult.TypeVAlign = TypeVAlignCenter

                sprLResult.Row = -1000
                sprLResult.Text = Trim(ADR_LOC!EQCD & "")
                sprLResult.RowHidden = True

                sprLResult.Row = -999
                sprLResult.Text = Trim(ADR_LOC!EQNM & "")
                
                intCnt = intCnt + 1             '/실가간 검사리스트 검사항목 읽기 증가
        
                '/검체번호별 검사결과 Column
                Select Case intCnt
                    Case 1 To 10:  sprDResult.Col = 1
                    Case 11 To 20: sprDResult.Col = 4
                    Case 21 To 30: sprDResult.Col = 7
                    Case 31 To 40: sprDResult.Col = 10
                End Select
                
                '/검체번호별 검사결과 Row
                If (intCnt Mod 10) = 0 Then
                    sprDResult.Row = 10
                Else
                    sprDResult.Row = intCnt Mod 10
                End If
                sprDResult.Text = Trim(ADR_LOC!EQNM & "")
                
                ADR_LOC.MoveNext
            Loop
            
            ADR_LOC.Close: Set ADR_LOC = Nothing
        End If
        
        Call CloseDB_LOC
    End If
Return
End Sub

Public Sub SUB_MM_KEY_CLEAR(ArgSection As String)
    Select Case ArgSection
        Case "1" '/검체번호별 세부정보
            lblBARCD = ""
            lblEXSEQ = ""
            lblSAMPLENO = ""
            lblDISKNOPOSNO = ""
            lblEXDT = ""
            lblRCDT = ""
            lblSDDT = ""
            lblORDDT = ""
            lblORDGB = ""
            lblPATNO = ""
            lblPATNM = ""
            lblSEXAGE = ""
            
        Case "2" '/검체번호별 검사결과
            '/(lCol As Long, lRow As Long, lCol2 As Long, lRow2 As Long, bDataOnly As Boolean)
            sprDResult.ClearRange 2, -1, 2, -1, True
            sprDResult.ClearRange 5, -1, 5, -1, True
            sprDResult.ClearRange 8, -1, 8, -1, True
            sprDResult.ClearRange 11, -1, 11, -1, True
            
        Case "3": '/실시간 검사리스트
            If sprLResult.MaxRows > 0 Then sprLResult.MaxRows = 0
    End Select
End Sub

Public Sub SUB_MM_PRINT()
'''    Dim strFont1  As String
'''    Dim strFont2  As String
'''    Dim strHead1  As String
'''
'''    If sprVIEW.MaxRows = 0 Then MsgBox "출력할 자료가 없습니다.", vbInformation, "확인": Exit Function
'''
'''    If MsgBox("출력하겠습니까?", vbQuestion + vbOKCancel, "출력여부") = vbCancel Then Exit Function
'''
'''    strFont1 = "/fn""굴림체""/fz""15""/fb1/fi0/fu1/fk0/fs1"
'''    strFont2 = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs2"
'''
'''    strHead1 = "/f1/c" & "거래처 코드" & "/n/n/n"
'''
'''    With sprVIEW
'''        .PrintAbortMsg = "거래처 코드 출력 중..."
'''        .PrintHeader = strFont1 + strHead1 + strFont2
'''        .PrintFooter = "/c" & "PAGE : " & "/P"
'''        .PrintBorder = True
'''        .PrintGrid = True
'''        .PrintColHeaders = True
'''        .PrintRowHeaders = True
'''        .PrintColor = False
'''        .PrintMarginTop = 500
'''        .PrintMarginBottom = 500
'''        .PrintMarginLeft = 500
'''        .PrintMarginRight = 0
'''        .PrintType = PrintTypeAll
'''        .PrintShadows = False
'''        .PrintUseDataMax = False
'''        .Action = ActionSmartPrint
'''    End With
End Sub

Private Sub cmdClear_Click()
    Call SUB_MM_KEY_CLEAR("1") '/검체번호별 세부정보
    Call SUB_MM_KEY_CLEAR("2") '/검체번호별 검사결과
    Call SUB_MM_KEY_CLEAR("3") '/실시간 검사리스트
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    MsgBox InStr(txtBuff, chrCR)
    Call SUB_COMM_PART_HUBIQUANPRO_BAR(txtBuff)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown, Txt
   
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
   
    If KeyCode = vbKeyM Then   ' 키의 조합 상태를 출력합니다.
        If mnuSetting.Visible = True Then
            mnuSetting.Visible = False
        Else
            mnuSetting.Visible = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
    
    DoEvents
    DoEvents
    DoEvents
End Sub

Private Sub Form_Resize()
    Dim intCnt  As Integer

On Error Resume Next
    '/object.Move Left, Top, Width, Height
    '/(((Me.Height - lngMeHeight) / 3) * 2) : 높이가 늘어나는 개체 3개, 디자인상 해당 개체 위에 늘어난 개체가 2개
    For intCnt = 0 To UBound(CW)
        Select Case CW(intCnt).Nm
            Case cmdClear.Name:     cmdClear.Move CW(intCnt).Left + (Me.Width - lngMeWidth), CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height
            Case cmdExit.Name:      cmdExit.Move CW(intCnt).Left + (Me.Width - lngMeWidth), CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height
            Case prgPatient.Name:   prgPatient.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case shpDResult.Name:   shpDResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case sprDResult.Name:   sprDResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case shpLResult.Name:   shpLResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case sprLResult.Name:   sprLResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height + (Me.Height - lngMeHeight)
        End Select
    Next intCnt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Call CloseDB_HIS
    Call CloseDB_ETC
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    Set frmEQ_Main = Nothing
End Sub

Private Sub mnuCodeSub_Click(Index As Integer)
    Select Case Index
        Case 0:
            MsgBox "의료장비와 통신 중에 장비검사코드 정보를 수정하면" & vbCrLf & _
                   "의도되지 않은 결과를 초래할 수 있습니다." & vbCrLf & vbCrLf & _
                   "장비검사코드 정보를 수정한 후엔 프로그램을 재 실행하십시오", vbExclamation, "주의"
        
            frmEQ공용_장비검사코드관리_조회.Show vbModal
    End Select
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInfo_Click()
    frmEQ공용_Info.Show vbModal
End Sub

Private Sub mnuJobModeAuto_Click()
    If mnuJobModeAuto.Checked = False Then
        If MsgBox("검사결과에 대해 HIS(병원정보시스템)로의 전송방식을 자동전송[Auto] (으)로 하겠습니까?" & vbCrLf & vbCrLf & _
                  "(주의: 의료장비와 통신중 일때는 전송방식을 바꾸지 마십시오!)", vbQuestion + vbOKCancel + vbDefaultButton2, "전송방식 변경 확인") = vbCancel Then Exit Sub
        mnuJobModeAuto.Checked = True
        staCondition.Panels.Item(3).Picture = LoadPicture(App.Path & "\Auto.jpg")
        mnuJobModeManual.Checked = False
    End If
End Sub

Private Sub mnuJobModeManual_Click()
    If mnuJobModeManual.Checked = False Then
        If MsgBox("검사결과에 대해 HIS(병원정보시스템)로의 전송방식을 수동전송[Manual] (으)로 하겠습니까?" & vbCrLf & vbCrLf & _
                  "(주의: 의료장비와 통신중 일때는 전송방식을 바꾸지 마십시오!)", vbQuestion + vbOKCancel + vbDefaultButton2, "전송방식 변경 확인") = vbCancel Then Exit Sub
        mnuJobModeManual.Checked = True
        staCondition.Panels.Item(3).Picture = LoadPicture(App.Path & "\Manual.jpg")
        mnuJobModeAuto.Checked = False
    End If
End Sub

Private Sub mnuJobSub_Click(Index As Integer)
    Select Case Index
        Case 0: 'frmWorkList.Show vbModal
        Case 1: frmEQ_검사결과관리.Show vbModal
    End Select
End Sub

Private Sub mnuSettingSub_Click(Index As Integer)
    Select Case Index
        Case 0: frmEQ공용_Set_Port.Show vbModal
        Case 1: gstrArgTemp1 = "HIS": frmEQ공용_Set_DB.Show vbModal
        Case 2: gstrArgTemp1 = "ETC": frmEQ공용_Set_DB.Show vbModal
        Case 3:
            txtSerialData.Left = 7800
            txtSerialData.Top = 2775
            If txtSerialData.Visible = False Then
                txtSerialData.Visible = True
            Else
                txtSerialData.Visible = False
            End If
        End Select
End Sub

Private Sub MSComm1_OnComm()
    Dim strOneByte  As String
    Dim strBuff     As String
    
    If shpCon(0).FillColor = &HFF& Then
        shpCon(0).FillColor = &HFF0000
    Else
        shpCon(0).FillColor = &HFF&
    End If

    If shpCon(1).FillColor = &HFF& Then
        shpCon(1).FillColor = &HFF0000
    Else
        shpCon(1).FillColor = &HFF
    End If

    strOneByte = MSComm1.Input

    txtBuff = txtBuff & strOneByte          '/전체 문장
    strOneLine = strOneLine & strOneByte    '/한 라인 담는 변수

    Select Case strOneByte
        Case chrLF
            If Mid(strOneLine, 1, 1) = "L" Then
                strBuff = txtBuff
                
                txtBuff = ""
                strOneLine = ""
                
                Call SUB_COMM_PART_HUBIQUANPRO_BAR(strBuff)
            Else
                strOneLine = ""
            End If
    End Select
End Sub

Private Sub sprLResult_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    If Cancel = True Then Exit Sub  '/임의로 Cancel을 True로 만들어 호출할 경우 처리하지 않게 할 수 있다.
    If NewRow = Row Then Exit Sub   '/Row가 변동없을 때에는 처리하지 않는다.
    If NewRow < 1 Then Exit Sub     '/선택한 Row가 유효한 내용이지 않으면 처리하지 않는다.
    
    Dim intLResultCol   As Integer
    Dim intCnt          As Integer
    Dim strResult       As String
    
    Call SUB_MM_KEY_CLEAR("1") '/검체번호별 세부정보
            
    lblBARCD = GET_CELL(sprLResult, 1, NewRow)
    lblEXSEQ = GET_CELL(sprLResult, 2, NewRow)
    lblSAMPLENO = GET_CELL(sprLResult, 3, NewRow)
    lblDISKNOPOSNO = GET_CELL(sprLResult, 4, NewRow) & "/" & GET_CELL(sprLResult, 5, NewRow)
    
    lblEXDT = GET_CELL(sprLResult, 8, NewRow)
    lblRCDT = GET_CELL(sprLResult, 9, NewRow)
    lblSDDT = GET_CELL(sprLResult, 10, NewRow)
    
    lblORDDT = GET_CELL(sprLResult, 11, NewRow)
    lblORDGB = GET_CELL(sprLResult, 12, NewRow)
    
    lblPATNO = GET_CELL(sprLResult, 13, NewRow)
    lblPATNM = GET_CELL(sprLResult, 14, NewRow)
    lblSEXAGE = GET_CELL(sprLResult, 15, NewRow)
    
    
    
    Call SUB_MM_KEY_CLEAR("2") '/검체번호별 검사결과
    
    For intLResultCol = gintEQ_StartCol To sprLResult.MaxCols
        '/읽기----------------------------------------------------------------------------------------------------/
        sprLResult.Col = intLResultCol
        sprLResult.Row = NewRow         '/실시간 검사리스트 검사결과 Row
        strResult = sprLResult.Text     '/실시간 검사리스트 검사결과 값
        '/읽기----------------------------------------------------------------------------------------------------/

        '/쓰기----------------------------------------------------------------------------------------------------/
        intCnt = intCnt + 1             '/실가간 검사리스트 검사항목 읽기 증가

        '/검체번호별 검사결과 Column
        Select Case intCnt
            Case 1 To 10:  sprDResult.Col = 2
            Case 11 To 20: sprDResult.Col = 5
            Case 21 To 30: sprDResult.Col = 8
            Case 31 To 40: sprDResult.Col = 11
        End Select

        '/검체번호별 검사결과 Row
        If (intCnt Mod 10) = 0 Then
            sprDResult.Row = 10
        Else
            sprDResult.Row = intCnt Mod 10
        End If

        sprDResult.Text = strResult
        '/쓰기----------------------------------------------------------------------------------------------------/
    Next intLResultCol
End Sub

Private Sub staCondition_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel = "COM" Then
        If txtSerialData.Visible = False Then
            txtSerialData.Visible = True
        Else
            txtSerialData.Visible = False
        End If
    End If
End Sub
