VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   " STA-R Interface"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   15105
   Begin IF_STA_R_국립암센터.MDButton cmdReset 
      Height          =   585
      Left            =   12540
      TabIndex        =   43
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear"
   End
   Begin IF_STA_R_국립암센터.MDButton cmdSend 
      Height          =   585
      Left            =   11340
      TabIndex        =   41
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "선택전송"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   900
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4260
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   690
      TabIndex        =   22
      Top             =   1080
      Width           =   165
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   345
      Left            =   210
      TabIndex        =   21
      Top             =   1080
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "번호"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.76
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   9255
      Left            =   150
      TabIndex        =   20
      Top             =   1020
      Width           =   14775
      _Version        =   393216
      _ExtentX        =   26061
      _ExtentY        =   16325
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   50
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":0442
      UserResize      =   2
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   285
      Left            =   450
      TabIndex        =   19
      Top             =   1950
      Width           =   435
      _Version        =   65536
      _ExtentX        =   767
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "번호"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.76
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   1125
      Left            =   8220
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   1755
      _Version        =   393216
      _ExtentX        =   3096
      _ExtentY        =   1984
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmInterface.frx":4E22
   End
   Begin FPSpread.vaSpread vasRes 
      Height          =   8115
      Left            =   8400
      TabIndex        =   15
      Top             =   1890
      Visible         =   0   'False
      Width           =   6435
      _Version        =   393216
      _ExtentX        =   11351
      _ExtentY        =   14314
      _StockProps     =   64
      ColHeaderDisplay=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmInterface.frx":5064
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   10305
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5467
            MinWidth        =   5467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2010-04-27"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 8:29"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "메디메이트 ☎(051)462-1751"
            TextSave        =   "메디메이트 ☎(051)462-1751"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   750
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
      RTSEnable       =   -1  'True
      EOFEnable       =   -1  'True
   End
   Begin VB.TextBox txtErr 
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   10920
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1245
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   9915
      _Version        =   65536
      _ExtentX        =   17489
      _ExtentY        =   1138
      _StockProps     =   15
      Caption         =   "  STA-R INTERFACE"
      ForeColor       =   -2147483633
      BackColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   14.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7950
         Picture         =   "frmInterface.frx":55F5
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   58
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox Text_Today 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   5910
         TabIndex        =   2
         Text            =   "2002/02/18"
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         Caption         =   "사용자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   59
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   225
         Left            =   4800
         TabIndex        =   3
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.TextBox Text_ini 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8940
      TabIndex        =   10
      Top             =   7830
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame Frame7 
      Height          =   9345
      Left            =   90
      TabIndex        =   13
      Top             =   840
      Width           =   14925
      Begin VB.CommandButton cmdDown 
         Height          =   525
         Left            =   3750
         Picture         =   "frmInterface.frx":5B7F
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   8610
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdUp 
         Height          =   525
         Left            =   2580
         Picture         =   "frmInterface.frx":5CB1
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   8610
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   5850
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   5310
      Width           =   2175
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   1110
      TabIndex        =   23
      Top             =   3180
      Width           =   2055
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   2055
      Left            =   7290
      TabIndex        =   24
      Top             =   4470
      Visible         =   0   'False
      Width           =   3285
      _Version        =   393216
      _ExtentX        =   5794
      _ExtentY        =   3625
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmInterface.frx":5DE0
   End
   Begin IF_STA_R_국립암센터.MDButton Command_close 
      Height          =   585
      Left            =   13740
      TabIndex        =   42
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "종료"
   End
   Begin VB.Frame Frame1 
      Height          =   9255
      Left            =   120
      TabIndex        =   25
      Top             =   1320
      Width           =   14865
      Begin VB.TextBox txtEquipID 
         Height          =   345
         Left            =   3930
         TabIndex        =   39
         Text            =   "10"
         Top             =   1230
         Width           =   1875
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Rack Pos"
         Height          =   375
         Left            =   7890
         TabIndex        =   38
         Top             =   1200
         Width           =   1635
      End
      Begin VB.CommandButton Command10 
         Caption         =   "결과입력"
         Height          =   375
         Left            =   6210
         TabIndex        =   37
         Top             =   1200
         Width           =   1635
      End
      Begin VB.TextBox txtEquipCode 
         Height          =   345
         Left            =   2040
         TabIndex        =   36
         Text            =   "0ADVI120"
         Top             =   1215
         Width           =   1875
      End
      Begin VB.CommandButton Command9 
         Caption         =   "장비ID조회"
         Height          =   375
         Left            =   390
         TabIndex        =   35
         Top             =   1200
         Width           =   1635
      End
      Begin VB.CommandButton Command8 
         Caption         =   "미검사상세목록"
         Height          =   375
         Left            =   5340
         TabIndex        =   34
         Top             =   780
         Width           =   1635
      End
      Begin VB.CommandButton Command7 
         Caption         =   "미검사목록"
         Height          =   375
         Left            =   3690
         TabIndex        =   33
         Top             =   780
         Width           =   1635
      End
      Begin VB.CommandButton Command6 
         Caption         =   "검사상세목록"
         Height          =   375
         Left            =   2040
         TabIndex        =   32
         Top             =   780
         Width           =   1635
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   4335
         Left            =   390
         TabIndex        =   31
         Top             =   1650
         Width           =   10065
         _Version        =   393216
         _ExtentX        =   17754
         _ExtentY        =   7646
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":6022
      End
      Begin VB.TextBox txtID 
         Height          =   345
         Left            =   6990
         TabIndex        =   30
         Text            =   "05111000003"
         Top             =   810
         Width           =   1875
      End
      Begin VB.CommandButton Command5 
         Caption         =   "검사목록"
         Height          =   375
         Left            =   390
         TabIndex        =   29
         Top             =   780
         Width           =   1635
      End
      Begin VB.CommandButton Command4 
         Caption         =   "서버시간"
         Height          =   375
         Left            =   390
         TabIndex        =   26
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lblDate2 
         AutoSize        =   -1  'True
         Caption         =   "서버시간1"
         Height          =   195
         Left            =   2250
         TabIndex        =   28
         Top             =   420
         Width           =   945
      End
      Begin VB.Label lblDate1 
         AutoSize        =   -1  'True
         Caption         =   "서버시간1"
         Height          =   195
         Left            =   3480
         TabIndex        =   27
         Top             =   420
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   120
      TabIndex        =   5
      Top             =   9150
      Visible         =   0   'False
      Width           =   14925
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5850
         TabIndex        =   12
         Top             =   300
         Width           =   1965
      End
      Begin VB.TextBox txtStart 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         TabIndex        =   7
         Top             =   420
         Width           =   885
      End
      Begin VB.TextBox txtEnd 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         TabIndex        =   6
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   " - "
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
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "번호"
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
         Left            =   450
         TabIndex        =   8
         Top             =   480
         Width           =   450
      End
   End
   Begin IF_STA_R_국립암센터.MDButton cmdCall 
      Height          =   585
      Left            =   10140
      TabIndex        =   56
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "결과조회"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hide Control"
      Height          =   4515
      Left            =   750
      TabIndex        =   44
      Top             =   5430
      Visible         =   0   'False
      Width           =   12465
      Begin VB.CommandButton cmdQC 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3510
         TabIndex        =   57
         Top             =   1410
         Width           =   1185
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   4650
         Style           =   1  '그래픽
         TabIndex        =   53
         Top             =   660
         Value           =   1  '확인
         Width           =   1125
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   1980
         TabIndex        =   51
         Top             =   630
         Width           =   1545
      End
      Begin VB.CommandButton cmdResSave 
         Caption         =   "결과저장"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8430
         TabIndex        =   50
         Top             =   3300
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList 작성"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         TabIndex        =   48
         Top             =   2520
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   285
         Left            =   450
         TabIndex        =   47
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdResCall 
         Caption         =   "QC 결과 전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4830
         TabIndex        =   46
         Top             =   3690
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdChangeUser 
         Caption         =   "사용자변경"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3180
         TabIndex        =   45
         Top             =   3690
         Visible         =   0   'False
         Width           =   1545
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1125
         Left            =   6840
         TabIndex        =   49
         Top             =   2940
         Visible         =   0   'False
         Width           =   1755
         _Version        =   393216
         _ExtentX        =   3096
         _ExtentY        =   1984
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":6264
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   375
         Left            =   450
         TabIndex        =   52
         Top             =   630
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "바코드번호"
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.74
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin IF_STA_R_국립암센터.MDButton Command_Config 
         Height          =   735
         Left            =   7020
         TabIndex        =   54
         Top             =   660
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "통신설정"
      End
      Begin IF_STA_R_국립암센터.MDButton Command_setup 
         Height          =   735
         Left            =   8160
         TabIndex        =   55
         Top             =   660
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "파일"
      Begin VB.Menu MnMainExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConf 
      Caption         =   "설정"
      Begin VB.Menu MnConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnConfExam 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "전송"
      Begin VB.Menu MnTransAuto 
         Caption         =   "Atuo"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "Manual"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const colCheckBox = 1
'Const colBarcode = 2
'Const colSeqNo = 3
'Const colReceno = 4
'Const colRack = 5
'Const colPos = 6
'Const colPID = 7
'Const colPName = 8
'Const colPSex = 9
'Const colPAge = 10
'Const colPJumin = 11
'Const colState = 12

Const colCheckBox = 1
Const colBarCode = 8
Const colSeqNo = 4      '검체
Const colReceno = 2     '접수번호
Const colRack = 5
Const colPos = 6
Const colPID = 3
Const colPName = 7
Const colPSex = 9
Const colPAge = 10
Const colPJumin = 11
Const colState = 12

Const colOrd = 13
Const colRes = 14
Const colDate = 15
Const colTime = 16

Const colSampleType = 17

Const colResult = 18
Dim colResult1 As Long

'Const colEquipCode = 1
'Const colExamCode = 2
'Const colExamName = 3
'Const colResult = 4
'Const colSeq = 5
'Const colRCheck = 6
'
''2004/10/21 이상은
''Const colRefLow = 7
'Const colResult1 = 7
'
'Const colRefHigh = 8

Dim gRow As Long                'VasID에서 같은 바코번호 존재 여부로 위치 체크

Dim gsBarCode As String
Dim gsPID As String
Dim gsRackNo As String
Dim gsPosNo As String
Dim gsResDateTime As String
Dim gRCol As Integer
Dim gsSeqNo As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String
Dim gsSampleType As String

Sub Var_Clear()
    gsBarCode = ""
    gsPID = ""
    gsRackNo = ""
    gsPosNo = ""
    gsResDateTime = ""
    gsSeqNo = ""
    gsExamCode = ""
    gsExamName = ""
    gsOrder = ""
    gsResult = ""
End Sub

Public Function Check_Result(argBarCode As String, argPID As String, argExamCode As String, _
                            argResult As String, ByVal argRow As Long, asSex As String) As Integer
    Dim sDiffRet, sDiffRet1 As String
    Dim PreResult   As String
    
    Dim sResClassCode As String     '결과종류
    Dim sLow        As String       '참조치
    Dim sHigh       As String
    Dim RefRet      As String
    Dim sPanicGubun As String
    Dim sPanicLow   As String       'Panic
    Dim sPanicHigh  As String
    Dim PanicRet    As String
    Dim sDeltaGubun As String
    Dim sDeltaLow   As String       'Delta
    Dim sDeltaHigh  As String
    Dim DeltaRet    As String
    
    Dim sTmpRece1, sTmpRet1 As String
    Dim sTmpRece2, sTmpRet2 As String
    Dim sMax_ReceNo As String
    Dim i           As Integer
    Dim sReceNo     As String
    Dim sPID        As String
    
    Dim sTmpStr As String
    
    Check_Result = -1
    
    If argBarCode = "" Then
        Exit Function
    End If
    
    If argExamCode = "" Then
        Exit Function
    End If
    

    RefRet = ""
    PanicRet = ""
    DeltaRet = ""
    
    sDiffRet = argResult
    If sDiffRet = "" Then
        Check_Result = -1
        Exit Function
    End If
    
    SQL = " Use NeoSoft"
    res = SendQuery(gServer, SQL)
    
    SQL = " Select LABM_MAN_FRES, LABM_MAN_TRES, LABM_WOM_FRES, LABM_WOM_TRES " & CR & _
          "From CC_LABM " & CR & _
          " Where LABN_ID = '" & Trim(argExamCode) & "' "
          
    res = db_select_Col(gServer, SQL)
    
'    sResClassCode = Trim(gReadBuf(0))
    
'    If sResClassCode = "1" Then '숫자
'참조치 체크
        sLow = ""
        sHigh = ""
        
        '숫자인지 아닌지 확인
        If IsNumeric(sDiffRet) = False Then
           MsgBox "결과형식이 일치하지 않습니다.", vbInformation, "알림"
           Check_Result = -1
           Exit Function
        End If
        
'        If IsNumeric(gReadBuf(13)) Then
'            If CInt(gReadBuf(13)) > 0 Then
'                sTmpStr = "#0."
'                For i = 1 To CInt(gReadBuf(13))
'                    sTmpStr = sTmpStr & "0"
'                Next i
'            Else
'                sTmpStr = "#0"
'            End If
'            sDiffRet = Format(sDiffRet, sTmpStr)
'            SetText vasRes, sDiffRet, argRow, colResult
'            SetText vasRes, sDiffRet, argRow, colResult1
'        End If
        
        Select Case asSex
        Case "M", ""
            sLow = Trim(gReadBuf(0))
            sHigh = Trim(gReadBuf(1))
        Case "F"
            sLow = Trim(gReadBuf(2))
            sHigh = Trim(gReadBuf(3))
        End Select
        
        If sLow = "" And sHigh = "" Then
            RefRet = ""
        ElseIf sLow = "" And sHigh <> "" Then   '이상
            If CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            End If
        ElseIf sLow <> "" And sHigh = "" Then   '이하
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            End If
        Else
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
                RefRet = ""
            End If
        End If


''Panic 체크
'        sPanicLow = ""
'        sPanicHigh = ""
'
'        sPanicGubun = Trim(gReadBuf(5))
'
'        Select Case asSex
'        Case "M", ""
'            sPanicLow = Trim(gReadBuf(6))
'            sPanicHigh = Trim(gReadBuf(7))
'        Case "F"
'            sPanicLow = Trim(gReadBuf(8))
'            sPanicHigh = Trim(gReadBuf(9))
'        End Select
'
'        If sPanicGubun = "0" Then '상한/하한
'            If sPanicLow = "" Or sPanicHigh = "" Then
'                PanicRet = ""
'            Else
'                If CCur(sPanicLow) > CCur(sDiffRet) Then
'                    PanicRet = "L"
'                ElseIf CCur(sPanicHigh) < CCur(sDiffRet) Then
'                    PanicRet = "H"
'                ElseIf CCur(sPanicLow) <= CCur(sDiffRet) And CCur(sPanicHigh) <= CCur(sDiffRet) Then
'                    PanicRet = ""
'                End If
'            End If
'        ElseIf sPanicGubun = "1" Then 'percent
'            If sPanicLow = "" Then
'                PanicRet = ""
'            Else
'                If CCur(sPanicLow) - CCur(sDiffRet) > 0 Then
'                    If ((CCur(sPanicLow) - CCur(sDiffRet)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
'                        PanicRet = "L"
'                    Else
'                        PanicRet = ""
'                    End If
'                ElseIf CCur(sPanicHigh) - CCur(sDiffRet) < 0 Then
'                    If ((CCur(sDiffRet) - CCur(sPanicLow)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
'                        PanicRet = "H"
'                    Else
'                        PanicRet = ""
'                    End If
'                Else
'                    PanicRet = ""
'                End If
'            End If
'        End If
'
'
''Delta 체크
'        sDeltaLow = ""
'        sDeltaHigh = ""
'
'        sTmpRece1 = ""
'        sTmpRet1 = ""
'        sTmpRece2 = ""
'        sTmpRet2 = ""
'        PreResult = ""
'
'        sMax_ReceNo = ""
''        sTmpRece1 = Trim(argForm.dtpReceDate.Value)
'        sReceNo = argBarCode
'
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where HID = '115' " & CR & _
'              " And PID = '" & Trim(argPID) & "' " & CR & _
'              " And ReceNo < '" & argBarCode & "' " & CR & _
'              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'              " Group By Result"
'
'        res = db_select_Col(gServer, SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'            PreResult = gReadBuf(0)
'        Else
'            PreResult = ""
'        End If
'
'        If PreResult <> "" Then
'          'PreResult = Trim(gReadBuf(0))
'          sDeltaGubun = Trim(gReadBuf(10))
'
'          sDeltaLow = Trim(gReadBuf(11))
'          sDeltaHigh = Trim(gReadBuf(12))
'
'            '이전결과에서 현재결과 뺀값이 sDiffRet임 (2002년 3월 15일 수정)
''            sDiffRet = PreResult - sDiffRet
'            sDiffRet1 = sDiffRet - PreResult
'            If sDeltaGubun = "0" Then '상한/하한
'                If sDeltaLow = "" Or sDeltaHigh = "" Then
'                    DeltaRet = ""
'                Else
'                    If CCur(sDeltaLow) > CCur(sDiffRet1) Then
'                        DeltaRet = "L"
'                    ElseIf CCur(sDeltaHigh) < CCur(sDiffRet1) Then
'                        DeltaRet = "H"
'                    ElseIf CCur(sDeltaLow) <= CCur(sDiffRet1) And CCur(sDeltaHigh) <= CCur(sDiffRet1) Then
'                        DeltaRet = ""
'                    End If
'                End If
'
'            ElseIf sDeltaGubun = "1" Then 'percent
'               If CInt(PreResult) = 0 Or CInt(sDiffRet) = 0 Then
'                  DeltaRet = ""
'               Else
'                   If sDeltaLow = "" Then
'                        DeltaRet = ""
'                    Else
'                        If (Abs(CCur(PreResult) - CCur(sDiffRet)) / CCur(PreResult)) * 100 >= CCur(sDeltaLow) Then
'                            DeltaRet = "D"
'                        Else
'                            DeltaRet = ""
'                        End If
'                    End If
'               End If
'            End If
'        End If
'
'    ElseIf sResClassCode = "2" Then '문자
'        Dim sRefValue As String
'        Dim sPanicValue As String
'        Dim sResult As String
'
'        sLow = ""
'        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
'
'        '2003/03/17 이상은 수정
'        '검사 항목 결과 참조 코드 체크에서 1 이상일 경우만 판정되게
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            Exit Function
'        End If
'
'        '2002년 3월 12일 +-에서 +/-로 수정
'        '2002년 5월 13일 NON-REACTIVE 판정 안돼서 추가
'        '2003년 2월 4일 이상은 수정 - 0-1로 참조치는 1이나 판정됨
'        '=================================================================================
'        '2002년 5월 13일 1 : 40 미만 판정 안됨
'        '2002년 6월 11일 (결과참조가 1:로 시작하면 판정체크 안하게 수정)
'        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
'            Exit Function
'        End If
'        '=================================================================================
'
'        Select Case UCase(sDiffRet)
'        Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-1"
'            sResult = 1
'        Case "+/-", "2", "+-", "2-5"
'            sResult = 2
'        Case "+", "POSITIVE", "양성", "3", "6-10"
'            sResult = 3
'        Case "++", "4", "11-20"
'            sResult = 4
'        Case "+++", "5", "21-30"
'            sResult = 5
'        Case "++++", "6"
'            sResult = 6
'        Case "+++++", "7"
'            sResult = 7
'        Case "++++++", "8"
'            sResult = 8
'        Case Else
'            sResult = sDiffRet
'        End Select
'        'sLow = "0-2"
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성", "1", "NON-REACTIVE", "0-2"
'                sRefValue = 1
'            Case "+/-", "2", "+-"
'                sRefValue = 2
'            Case "+", "POSITIVE", "양성", "3"
'                sRefValue = 3
'            Case "++", "4"
'                sRefValue = 4
'            Case "+++", "5"
'                sRefValue = 5
'            Case "++++", "6"
'                sRefValue = 6
'            Case "+++++", "7"
'                sRefValue = 7
'            Case "++++++", "8"
'                sRefValue = 8
'            Case Else
'                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
'                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
'                    RefRet = sDiffRet
'                End If
'            End Select
'            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'
'            ElseIf sRefValue < sResult Then
''                RefRet = "H"
'                RefRet = sDiffRet
'
''                argTable.Row = argRow
''                argTable.Col = iresDecision
''                argTable.ForeColor = RGB(205, 55, 0)
'
'
'            End If
'        End If
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            RefRet = Trim(GetText(argTable, argRow, iresDecision))
'        End If
'        sLow = ""
'        sLow = Trim(GetText(argTable, argRow, iresPanicValue))
'        If Trim(sLow) <> "" Then
'            Select Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "음성"
'                sPanicValue = 1
'            Case "+/-"
'                sPanicValue = 2
'            Case "+", "POSITIVE", "양성"
'                sPanicValue = 3
'            Case "++"
'                sPanicValue = 4
'            Case "+++"
'                sPanicValue = 5
'            Case "++++"
'                sPanicValue = 6
'            Case "+++++"
'                sPanicValue = 7
'            Case "++++++"
'                sPanicValue = 8
'            Case Else
'                If UCase(sDiffRet) > UCase(sLow) Then
'                    PanicRet = sDiffRet
'                End If
'            End Select
'            If sPanicValue < sResult Then
'                'PanicRet = "H"
'                PanicRet = sDiffRet
'            End If
'        End If
'
'        'Delta Check
'        sMax_ReceNo = ""
'        DeltaRet = ""
'        sReceNo = Trim(GetText(argForm.vasPatient, 1, 1))
'        sPID = Trim(GetText(argForm.vasPatient, 1, 3))
'
'        SQL = "Select Result,Max(ReceNo) From ExamRes " & CR & _
'              " Where PID = '" & sPID & "' " & CR & _
'              " And ReceNo < '" & sReceNo & "' " & CR & _
'              " And ExamCode = '" & Trim(GetText(argTable, argRow, iresExamCode)) & "' " & CR & _
'              " Group By Result"
'
'        res = db_select_Col(SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'               If sDiffRet <> gReadBuf(0) Then
'                  DeltaRet = "D"
'               End If
'        Else
'            DeltaRet = ""
'        End If
'    End If
    
'    SetText vasRes, RefRet, argRow, colRCheck
    

    '2002년 2월 15일 수정 (판정시 H, L 일때 글자 색깔 변화)
    '2002년 3월 14일 수정 (판정시 L일때는 파란색 그 외는 빨간색)
'    If RefRet = "L" Then
'        vasRes.Row = argRow
'        vasRes.Col = colRCheck
'        vasRes.ForeColor = RGB(65, 105, 225)
'    Else
'        vasRes.Row = argRow
'        vasRes.Col = colRCheck
'        vasRes.ForeColor = RGB(205, 55, 0)
'    End If

    
    Check_Result = 1

End Function

Private Sub ChkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If

End Sub

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub

Private Sub cmdCall_Click()
    Dim iRow As Long
    Dim i As Long
    Dim j As Long
    Dim y As Long
    
    ClearSpread vasID
    ClearSpread vasRes
    
'    SQL = "select distinct levelname, '', '', '0', '0', examtime, '', '', '', 'F' " & vbCrLf & _
'          "from qc_res " & vbCrLf & _
'          "where equipno  = '" & Trim(gEquip) & "' " & vbCrLf & _
'          "  and examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' "
'    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    
    SQL = "select receno, pid, seqno, diskno, posno, pname, barcode, psex, page, jumin, sendflag, max(recedate)" & _
          "from pat_res " & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag in ('B','C') " & vbCrLf & _
          "group by barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag "
    SQL = SQL & vbCrLf & " Union " & vbCrLf
    SQL = SQL & vbCrLf & _
          "select receno, pid, seqno, diskno, posno, pname, barcode,psex, page, jumin, sendflag, max(recedate)" & _
          "from pat_res " & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag not in ('B','C') " & vbCrLf & _
          "group by barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag " & vbCrLf & _
          "order by diskno,posno"
    res = db_select_Vas(gLocal, SQL, vasID, vasID.DataRowCnt + 1, 2)
    
'    SQL = "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, refvalue, panicvalue, max(recedate)" & _
'          "from pat_res " & _
'          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'          "group by barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, refvalue, panicvalue " & vbCrLf & _
'          "order by diskno,posno"
'    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    vasSort vasID, colRack, colPos
    
    For iRow = 1 To vasID.DataRowCnt
        Select Case Trim(GetText(vasID, iRow, colState))
        Case "B"
            SetText vasID, "결과", iRow, colState
        Case "C"
            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
            'SetForeColor vasID, iRow, iRow, colState, colState, 255, 0, 0
            SetText vasID, "완료", iRow, colState
        Case "A"
            SetText vasID, "오더", iRow, colState
        End Select
        
        '결과 불러오기
        ClearSpread vasTemp
        
        SQL = " Select examcode, result From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasID, iRow, colBarCode)) & "' " & vbCrLf & _
              " And pid = '" & Trim(GetText(vasID, iRow, colPID)) & "' "
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        For i = 1 To vasTemp.DataRowCnt
            For j = 1 To UBound(gArrEquip)
                If Trim(GetText(vasTemp, i, 1)) = gArrEquip(j, 3) Then
                    y = colResult + (j - 1) * 4
                    Exit For
                End If
            Next j
            
            SetText vasID, Trim(GetText(vasTemp, i, 2)), iRow, y
        Next i
    Next iRow
    

    
End Sub

Private Sub cmdChangeUser_Click()
'    frmUserChange.Show 1
End Sub

Private Sub cmdDelete_Click()
    Dim lRow As Long
    Dim lsPID As String
    Dim lsReceNo1 As String
    Dim lsReceNo2 As String
    
    Dim sStart As String
    Dim send As String
    
    sStart = Trim(txtStart.Text)
    send = Trim(txtEnd.Text)
    
    If sStart <> "" And send <> "" Then
        For lRow = sStart To send
            lsPID = Trim(GetText(vasID, lRow, 5))
            lsReceNo1 = Trim(GetText(vasID, lRow, 11))
            lsReceNo2 = Trim(GetText(vasID, lRow, 12))
            SQL = "Delete from pat_res " & vbCrLf & _
                  "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                  "  and equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and pid = '" & lsPID & "' " & vbCrLf & _
                  "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                  "  and receno1 = '" & lsReceNo2 & "' "
            res = SendQuery(gLocal, SQL)
            
            DeleteRow vasID, lRow, lRow
        Next lRow
    Else
        lRow = 1
        Do While lRow <= vasID.DataRowCnt
            vasID.Row = lRow
            vasID.Col = 1
            If vasID.Value = 1 Then
                lsPID = Trim(GetText(vasID, lRow, 5))
                lsReceNo1 = Trim(GetText(vasID, lRow, 11))
                lsReceNo2 = Trim(GetText(vasID, lRow, 12))
                SQL = "Delete from pat_res " & vbCrLf & _
                      "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  and equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and pid = '" & lsPID & "' " & vbCrLf & _
                      "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                      "  and receno1 = '" & lsReceNo2 & "' "
                res = SendQuery(gLocal, SQL)
                
                DeleteRow vasID, lRow, lRow
            Else
                lRow = lRow + 1
            End If
        Loop
    End If
    
    MsgBox "삭제 완료"
    chkAll.Value = 0
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow + 1
    vasActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdOrder_Click()

End Sub

Private Sub cmdQC_Click()
    'frmQCResSch.Show
End Sub

Private Sub cmdResCall_Click()
'    frmResult.Show 1
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
'    If chkAll.Value = 1 Then
            For i = 1 To vasID.DataRowCnt
                vasID.Row = i
                vasID.Col = 1

                If vasID.Value = 1 Then
                    DeleteRow vasID, i, i
                    i = i - 1
                End If
            Next i
'
'            chkAll.Value = 0
'    Else
'        vasID.Row = 1
'        vasID.Row2 = vasID.MaxRows
'        vasID.Col = 1
'        vasID.Col2 = vasID.MaxCols
'        vasID.BlockMode = True
'        vasID.BackColor = RGB(255, 255, 255)
'        vasID.Action = 3
'        vasID.BlockMode = False
'    End If
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetBackColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols + 4, 255, 255, 255
    
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    'ClearSpread vasID
    ClearSpread vasRes
    
    Text_Today = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
End Sub

Private Sub cmdResSave_Click()
    'Proc_Result txtBarcode
End Sub

Private Sub cmdSend_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
'            If Left(Trim(GetText(vasID, lRow, colBarCode)), 1) = "9" Then
'                res = ToQCServer(lRow)
'            Else
                res = ToServer(lRow)
'            End If
            
            If res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "실패", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 0
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "완료", lRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      " And equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, lRow, colBarCode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
        End If
    Next lRow
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow - 1
    vasActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1
End Sub

Private Sub Command_close_Click()
    Unload Me
End Sub

'Private Sub Command_config_Click()
'    frmConfig.Show 1
'End Sub


'Private Sub Command_setup_Click()
'    frmOrderCode.Show 1
'    GetExamCode
'End Sub

Private Sub Command1_Click()
    STA_R txtData
    txtData = ""
End Sub

'Private Sub Command10_Click()
'    Dim oerrmsg$
'    Dim ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$
'    Dim lRow As Long
'
'    If vasList.DataRowCnt < 1 Then Exit Sub
'
'    ReDim ispcid(vasList.DataRowCnt)
'    ReDim iexamcode(vasList.DataRowCnt)
'    ReDim iresult(vasList.DataRowCnt)
'    ReDim ierrflag(vasList.DataRowCnt)
'    ReDim iequipcd(vasList.DataRowCnt)
'
'    For lRow = 1 To vasList.DataRowCnt
'        ispcid(lRow - 1) = Trim(GetText(vasList, lRow, 1))
'        iexamcode(lRow - 1) = Trim(GetText(vasList, lRow, 6))
'        iresult(lRow - 1) = Trim(GetText(vasList, lRow, 8))
'        ierrflag(lRow - 1) = ""
'        iequipcd(lRow - 1) = Trim(txtEquipCode)
'        'iequipcd(lRow - 1) = ""
'    Next lRow
'    res = sl_online_result_ul_e(oerrmsg, ispcid(), iexamcode(), iresult(), ierrflag(), iequipcd(), "")
'    If res < 0 Then
'        MsgBox "저장 에러"
'    Else
'        MsgBox "저장 확인 : " & res
'    End If
'
'End Sub

'Private Sub Command11_Click()
'    Dim oerrmsg$
'    Dim ispcid$(), imach_id$(), ipos_flag$(), irack_id$(), irack_pos$()
'
'    ReDim ispcid(0)
'    ReDim imach_id(0)
'    ReDim ipos_flag(0)
'    ReDim irack_id(0)
'    ReDim irack_pos(0)
'
'    ispcid(0) = Trim(txtID)
'    imach_id(0) = Trim(txtEquipID)
'    ipos_flag(0) = "E"
'    irack_id(0) = "1001"
'    irack_pos(0) = "1"
'
'    res = sl_upd_spc_pos("", ispcid(), imach_id(), ipos_flag(), irack_id(), irack_pos())
'    MsgBox res
'End Sub


Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
    SQL = "CREATE INDEX resindex1 ON pat_res (examdate,equipno,barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex1 created"
    Else
        MsgBox "resindex1 failed"
    End If
    SQL = "CREATE INDEX resindex2 ON pat_res (examdate,equipno,barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex2 created"
    Else
        MsgBox "resindex2 failed"
    End If
    
    SQL = "CREATE INDEX resindex3 ON pat_res (barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex3 created"
    Else
        MsgBox "resindex3 failed"
    End If
    
    SQL = "CREATE INDEX resindex4 ON pat_res (barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex4 created"
    Else
        MsgBox "resindex4 failed"
    End If
End Sub

Private Sub Form_Load()
    Dim sDate As String

    If App.PrevInstance Then
        End
    End If

    Me.Left = 0
    Me.Top = 0
'    Me.Height = 11190
'    Me.Width = 15360

    cmdReset_Click

'    GetSetup

    MSComm1.CommPort = gSetup.gPort
'    MSComm1.RTSEnable = gSetup.gRTSEnable
'    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If

    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If

    '서버일자 가져오기
    'Text_Today = Format(CDate(GetDateFull), "yyyy/mm/dd")

    '해당 장비의 검사코드 가져오기
    GetExamCode

    sDate = Format(DateAdd("y", CDate(Text_Today.Text), -30), "yyyymmdd")
    SQL = " Delete From pat_res Where examdate < '" & sDate & "' "
    res = SendQuery(gLocal, SQL)

    gRow = -1

    vasID.MaxRows = 0

    vasID.RowHeight(-1) = 13
    lblUser.Caption = gIFUser


End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim lsbarcode As String
    Dim i As Integer
    Dim sRes As String
    
    Get_Sample_Info = -1
    
    lsbarcode = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호

    sRes = Online_XML(gXml_S03, lsbarcode)
    
    SetText vasID, gPat_Info_Select.ACPTNO_1, asRow, colSeqNo
    SetText vasID, gPat_Info_Select.SPC_CD_1, asRow, colReceno

    SetText vasID, gPat_Info_Select.PT_NO, asRow, colPID
    SetText vasID, gPat_Info_Select.PT_NM, asRow, colPName

    SetText vasID, gPat_Info_Select.Sex, asRow, colPSex
    SetText vasID, gPat_Info_Select.Age, asRow, colPAge

    vasID.RowHeight(asRow) = 20
        
        
    
'    sID = ""
'    sID = Trim(GetText(vasID, asRow, colBarCode))
'
'    'Clear*****************************
'    SetText vasID, "", asRow, colSeqNo
'    SetText vasID, "", asRow, colReceno
'
'    SetText vasID, "", asRow, colPID
'    SetText vasID, "", asRow, colPName
'
'    SetText vasID, "", asRow, colPSex
'    SetText vasID, "", asRow, colPAge
'    '**********************************
'
'    i = Get_PatInfo(Trim(sID))
'
'    If i = 1 Then
'        'SetText vasID, gPatient_Info.SPC_CD, asRow, colSeqNo
'        SetText vasID, gPatient_Info.WD_NO, asRow, colSeqNo
'        SetText vasID, gPatient_Info.ACPT_NO, asRow, colReceno
'
'        SetText vasID, gPatient_Info.PTNO, asRow, colPID
'        SetText vasID, gPatient_Info.PATNAME, asRow, colPName
'
'        SetText vasID, gPatient_Info.SEX, asRow, colPSex
'        SetText vasID, gPatient_Info.AGE, asRow, colPAge
'
'        Get_Sample_Info = 1
'    Else
'        SetText vasID, "", asRow, colSeqNo
'        SetText vasID, "", asRow, colReceno
'
'        SetText vasID, "", asRow, colPID
'        SetText vasID, "", asRow, colPName
'
'        SetText vasID, "", asRow, colPSex
'        SetText vasID, "", asRow, colPAge
'    End If
    
End Function

Function Get_QC_Info(ByVal asRow As Long) As Integer
'    Dim sID As String
'    Dim i As Integer
'
'    Get_QC_Info = -1
'
'    sID = ""
'    sID = Trim(GetText(vasID, asRow, colBarCode))
'
'    'i = Get_QCInfo(Trim(sID))
'    i = Get_QCList(Trim(sID), 1)
'
'    If i = 1 Then
'        SetText vasID, "", asRow, colSeqNo
'        SetText vasID, gQC_Info(0).LOT_NO, asRow, colReceno
'
'        SetText vasID, gQC_Info(0).CTRL_CD, asRow, colPID
'        SetText vasID, "QC", asRow, colPName
'
'        SetText vasID, "", asRow, colPSex
'        SetText vasID, "", asRow, colPAge
'
'        Get_QC_Info = 1
'    Else
'        SetText vasID, "", asRow, colSeqNo
'        SetText vasID, "", asRow, colReceno
'
'        SetText vasID, "", asRow, colPID
'        SetText vasID, "", asRow, colPName
'
'        SetText vasID, "", asRow, colPSex
'        SetText vasID, "", asRow, colPAge
'    End If
End Function


Function EquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String

    EquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, vasTemp1)
    
    If vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        End If
    Next i

    SQL = " Select SUCD From LRESULT " & CR & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    res = db_select_Col(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        EquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function

Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    
    SQL = "Select equipcode, examcode, examname, reflow, refhigh " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  seqno "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    If res > 0 Then
        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 6)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    vasList.MaxCols = UBound(gArrEquip) + 1
    vasID.MaxCols = vasTemp.DataRowCnt * 5 + colResult - 1 + 1
    
    colResult1 = vasTemp.DataRowCnt * 4 + colResult
    
    For i = 1 To vasTemp.DataRowCnt
        gArrEquip(i, 1) = i
        For j = 1 To 5
            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
        
        SetText vasID, gArrEquip(i, 4), 0, colResult + (i - 1) * 4
        vasID.ColWidth(colResult + (i - 1) * 4 + 1) = 0
        vasID.ColWidth(colResult + (i - 1) * 4 + 2) = 0
        vasID.ColWidth(colResult + (i - 1) * 4 + 3) = 0
        
        vasID.ColWidth(colResult1 + i - 1) = 0
        
        SetText vasList, gArrEquip(i, 4), 0, i + 1

        SetText vasList, gArrEquip(i, 4), 0, i + 1
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Local
    
    Unload Me
    
    End
    
End Sub

Private Sub MDButton2_Click()

End Sub

Private Sub MnConfExam_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub MnConfig_Click()
    frmConfig.Show 1
End Sub


Private Sub MnTransAuto_Click()
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
End Sub

Private Sub MnTransManual_Click()
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
End Sub

Private Sub MSComm1_OnComm()
    Dim lsChar As String
    Dim sSendData As String
    
    lsChar = MSComm1.Input
    
    Select Case lsChar
    Case chrENQ
        If Text_Today <> Format(CDate(Date), "yyyy/mm/dd") Then
            cmdReset_Click
            
        End If
        txtData = lsChar
        SaveData "[RX]" & txtData
        
        txtData = ""
        MSComm1.Output = chrACK
        SaveData "[TX]" & chrACK
        
    Case chrSTX
        txtData = ""
        
    Case chrETX
        SaveData "[RX]" & txtData
        
        STA_R txtData
        
        txtData = ""
        MSComm1.Output = chrACK
        SaveData "[TX]" & chrACK
        
    Case chrEOT
        txtData = txtData & lsChar
        SaveData "[RX]" & txtData
        
        txtData = ""
        MSComm1.Output = chrACK
        SaveData "[TX]" & chrACK
        
        'SetText vasID, "수신", gRow, colState
        
        If gOrderMessage = "Q" Then
            gTxMsgFlag = ""
            gCurTxCnt = 1
            
            gPatient = "P|1|||^^^" & chrCR & chrETX
            gOrder = "O|1|" & gsBarCode & "|" & "|" & gsOrder & "|R" & chrCR & chrETX
            MSComm1.Output = chrENQ
            SaveData "[TX]" & chrENQ
        End If
    
    Case chrACK
        Select Case gTxMsgFlag
        Case ""  'ENQ 보낸 다음 처음 받은 ACK
            'Header 보내기
            gHeader = "H|\^&|||99^2.00|||||||||" & gDate & chrCR & chrETX
            sSendData = chrSTX & CStr(gCurTxCnt) & gHeader & ASTM_CSum(CStr(gCurTxCnt) & gHeader) & chrCR & chrLF
            
            gTxMsgFlag = "H"
            MSComm1.Output = sSendData
            
            gPreData = sSendData
            gCurTxCnt = gCurTxCnt + 1
            If gCurTxCnt = 8 Then
                gCurTxCnt = 0
            End If
        Case "H"
            'patient 보내기
            sSendData = chrSTX & CStr(gCurTxCnt) & gPatient & ASTM_CSum(CStr(gCurTxCnt) & gPatient) & chrCR & chrLF
            gTxMsgFlag = "P"
            
            MSComm1.Output = sSendData
            
            gPreData = sSendData
            gCurTxCnt = gCurTxCnt + 1
            If gCurTxCnt = 8 Then
                gCurTxCnt = 0
            End If
        Case "P"
            'Test Order 보내기
            sSendData = chrSTX & CStr(gCurTxCnt) & gOrder & ASTM_CSum(CStr(gCurTxCnt) & gOrder) & chrCR & chrLF
            gTxMsgFlag = "O"
            
            MSComm1.Output = sSendData
            
            gPreData = sSendData
            gCurTxCnt = gCurTxCnt + 1
            If gCurTxCnt = 8 Then
                gCurTxCnt = 0
            End If
        Case "O"
            'Message Terminator 보내기
            sSendData = chrSTX & CStr(gCurTxCnt) & gMsgEnd & ASTM_CSum(CStr(gCurTxCnt) & gMsgEnd) & chrCR & chrLF
            gTxMsgFlag = "L"
            
            MSComm1.Output = sSendData
            
            gPreData = sSendData
            gCurTxCnt = gCurTxCnt + 1
            If gCurTxCnt = 8 Then
                gCurTxCnt = 0
            End If
        Case "L"
            sSendData = chrEOT
            gPreData = sSendData
            
            MSComm1.Output = sSendData
            gTxMsgFlag = ""
                     
        End Select
        SaveData "[TX]" & sSendData
    Case Else
        txtData = txtData & lsChar
    End Select

End Sub

Sub STA_R(asData As String)
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    Dim iPos As Integer
    Dim iCnt As Integer
    
    Dim lsID As String
    Dim lsPID As String
    Dim lsPName As String
    Dim lsJumin1 As String
    Dim lsJumin2 As String
    Dim lsPSex As String
    Dim lsPage As String

    Dim lsTestID As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsResult As String
    Dim lsExamDate As String
    Dim sSampleType As String
    Dim lsResFlag As String
    Dim lsErrFlag As String
    
    Dim rv As Integer
    Dim vTemp As String
    
    Dim liRet As Long
    Dim iRow As Integer
    Dim jRow As Integer
    
    Dim lCol As Integer
    
    Dim sCnt As String
    Dim sOCnt As String
    Dim sRCnt As String
    
    If asData = "" Then
        Exit Sub
    End If

'    If Text_Today <> Format(GetDateFull, "YYYY-MM-DD") Then
'        Text_Today = Format(GetDateFull, "YYYY-MM-DD")
'    End If
    
    TablePtr = 1
' ----- for start
    For j = 1 To Len(asData)
        If (Mid(asData, j, 1) = "|") Then
            TablePtr = TablePtr + 1
            ResultTbl(TablePtr) = " "
        Else
            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
        End If
    Next j
' ------- for end
    
    If Mid(ResultTbl(1), 2, 1) = "H" Then     'Header Record
        Var_Clear
        
        iCnt = 0
        
        For i = 1 To Len(asData)
            If Mid(asData, i, 1) = "|" Then
                iCnt = iCnt + 1

                Select Case iCnt
                    Case 11
                        gsSampleType = Mid(asData, i + 1, 1)
                    Case 13
                        gDate = Mid(asData, i + 1, 14)      '장비에서 받은 날짜시간
                End Select
            End If
        Next i
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "Q" Then     'Request Information Record
        gOrderMessage = "Q"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        'sTmp = Mid(sTmp, i + 1, 11)
        sTmp = Trim(Mid(sTmp, i + 1))
        iPos = InStr(1, sTmp, Chr(13))
        If iPos > 0 Then
            gsBarCode = Mid(sTmp, 1, iPos - 1)
        Else
            gsBarCode = sTmp
        End If
        
        gRow = -1
        For i = 1 To vasID.DataRowCnt
            If gsBarCode <> "" Then  '메뉴얼일경우 바코드 없음
                If Trim(GetText(vasID, i, colBarCode)) = gsBarCode Then
                    gRow = i
                    Exit For
                End If
            End If
        Next i

        If gRow = -1 Then
            gRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < gRow Then
                vasID.MaxRows = gRow
            End If
        End If
    
        SetText vasID, gsBarCode, gRow, colBarCode
        
        vasActiveCell vasID, gRow, colPID
        
        SetForeColor vasID, gRow, gRow, 1, colState, 0, 0, 0
        
        '환자정보 불러오기
        If Trim(GetText(vasID, gRow, colPName)) = "" Then
'            If gsSampleType = "Q" Or Left(Trim(GetText(vasID, gRow, colBarCode)), 1) = "9" Then
'                'Get_QC_Info gRow
'
'                If Trim(GetText(vasID, gRow, colPName)) = "" Then
'                    SetText vasID, "QC", gRow, colPName
'                End If
'            Else
                Get_Sample_Info gRow
'            End If
        End If
    
        'Order 만들기***********************
'        If gsSampleType = "Q" Or Left(Trim(GetText(vasID, gRow, colBarCode)), 1) = "9" Then
'
'        Else
            res = MakeOrder(gRow, gsBarCode)
'        End If
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "O") Then          'Test Order Record
        iPos = InStr(1, Trim(ResultTbl(3)), Chr(13))
        If iPos > 0 Then
            gsBarCode = Mid(Trim(ResultTbl(3)), 1, iPos - 1)   '검체번호
        Else
            gsBarCode = Trim(ResultTbl(3))  '검체번호
        End If

        gRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarCode)) = gsBarCode Then
                gRow = i
                Exit For
            End If
        Next i
        
        If gRow < 0 Then
            gRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < gRow Then
                vasID.MaxRows = gRow
            End If
        End If
        
        SetText vasID, gsBarCode, gRow, colBarCode

        SetText vasID, gsSampleType, gRow, colSampleType
        
        '샘플정보 가져오기
'        If gsSampleType = "Q" Then
'            If Trim(GetText(vasID, gRow, colPName)) = "" Then
'                If gsSampleType = "Q" Or Left(Trim(GetText(vasID, gRow, colBarCode)), 1) = "9" Then
'                    'Get_QC_Info gRow
'
'                    If Trim(GetText(vasID, gRow, colPName)) = "" Then
'                        SetText vasID, "QC", gRow, colPName
'                    End If
'                End If
'            End If
'        Else
            If Trim(GetText(vasID, gRow, colPName)) = "" Then
                '환자정보 불러오기*****************
                Get_Sample_Info gRow
            End If
'        End If
    End If

    If (Mid(ResultTbl(1), 2, 1) = "R") Then     'Result
        gOrderMessage = "R"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        lsTestID = Trim(sTmp)           '장비코드
        gOrderCode = lsTestID
        
        
        sTmp = ResultTbl(4)
        lsResult = Trim(sTmp)           '결과
        
        
        gsResDateTime = ResultTbl(13)    'result time
            
        lsExamCode = ""
        lsExamName = ""
        
        For i = 1 To UBound(gArrEquip)
            If lsTestID = gArrEquip(i, 2) Then
                lCol = (gArrEquip(i, 1) - 1)
                
                lsExamCode = Trim(gArrEquip(i, 3))
                lsExamName = Trim(gArrEquip(i, 4))
                SetText vasID, lsResult, gRow, colResult + lCol * 4
                SetText vasID, lsResult, gRow, colResult1 + lCol
            End If
        Next i
    
        lsID = Trim(GetText(vasID, gRow, colBarCode))
        lsResult = lsResult
        gsSeqNo = Trim(GetText(vasID, gRow, colSeqNo))
        gsRackNo = Trim(GetText(vasID, gRow, colRack))
        gsPosNo = Trim(GetText(vasID, gRow, colPos))
        
        lsPName = Trim(GetText(vasID, gRow, colPName))
        lsPSex = Trim(GetText(vasID, gRow, colPSex))
        lsPage = Trim(GetText(vasID, gRow, colPAge))
        lsJumin1 = Trim(GetText(vasID, gRow, colPJumin))
    
        lsPID = Trim(GetText(vasID, gRow, colPID))
    '    sSampleType = Trim(GetText(vasID, gRow, colSampleType))
        
        'Local Table Insert
        '환자 데이타 ====================================================================================
        db_BeginTran gLocal
        
'        If sSampleType = "Q" Then   'QC Data Local에 저장
'            Save_Local_QC gsResDateTime, _
'                          Trim(GetText(vasID, gRow, colBarCode)), _
'                          lsExamCode, _
'                          lsResult, _
'                          lsResult
'
'        Else    'Sample Data Local에 저장
            sCnt = ""
            
            SQL = "Select count(*) from pat_res " & vbCrLf & _
                  "where examdate = '" & Format(Text_Today.Text, "yyyymmdd") & "' " & vbCrLf & _
                  "and equipno = '" & gEquip & "' " & vbCrLf & _
                  "and barcode = '" & lsID & "' and equipcode = '" & lsTestID & "' "
            'res = db_select_Var(gLocal, SQL, sCnt)
            res = db_select_Col(gLocal, SQL)
            If res <= 0 Then
                SaveQuery SQL
                db_RollBack gLocal
                Exit Sub
            End If
            
            If gReadBuf(0) = "" Then
                sCnt = 0
            Else
                sCnt = Trim(gReadBuf(0))
            End If
            
            If Not IsNumeric(lsPage) Then
                lsPage = "0"
            End If
            
    
            If CInt(sCnt) = 0 Then
                '입력
                SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
                      "pid, pname, jumin, page, psex, resdate, receno, " & _
                      "equipcode, examcode, result, result1, sendflag, examname, " & _
                      "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
                      "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
                      "'" & Trim(GetText(vasID, gRow, colBarCode)) & "', '" & Trim(GetText(vasID, gRow, colSeqNo)) & "'," & _
                      "'" & Trim(GetText(vasID, gRow, colRack)) & "', '" & Trim(GetText(vasID, gRow, colPos)) & "', " & _
                      "'" & Trim(GetText(vasID, gRow, colPID)) & "', " & vbCrLf & _
                      "'" & Trim(GetText(vasID, gRow, colPName)) & "', '" & Trim(GetText(vasID, gRow, colPJumin)) & "', " & _
                      "'" & Trim(GetText(vasID, gRow, colPAge)) & "', '" & Trim(GetText(vasID, gRow, colPSex)) & "', " & _
                      "'" & lsExamDate & "', '" & Trim(GetText(vasID, gRow, colReceno)) & "', " & vbCrLf & _
                      "'" & lsTestID & "', '" & lsExamCode & "',  " & _
                      "'" & Trim(lsResult) & "', '" & Trim(lsResult) & "', 'B', '" & lsExamName & "', " & vbCrLf & _
                      "'',  " & _
                      "'" & Trim(GetText(vasID, gRow, colOrd)) & "', '" & Trim(GetText(vasID, gRow, colRes)) & "', '" & Trim(GetText(vasID, gRow, colDate)) & "') "
              
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            ElseIf CInt(sCnt) > 0 Then
    
    
                SQL = " Update pat_res Set " & vbCrLf & _
                      " diskno = '" & Trim(GetText(vasID, gRow, colRack)) & "', " & vbCrLf & _
                      " posno  = '" & Trim(GetText(vasID, gRow, colPos)) & "', " & vbCrLf & _
                      " result = '" & Trim(lsResult) & "', " & vbCrLf & _
                      " result1 = '" & Trim(lsResult) & "', " & vbCrLf & _
                      " refflag = '', " & vbCrLf & _
                      " refvalue = '" & Trim(GetText(vasID, gRow, colOrd)) & "', " & vbCrLf & _
                      " panicvalue = '" & Trim(GetText(vasID, gRow, colRes)) & "', " & vbCrLf & _
                      " resdate = '" & lsExamDate & "' " & vbCrLf & _
                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      " And equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, gRow, colBarCode)) & "' " & vbCrLf & _
                      " And equipcode = '" & lsTestID & "' " & vbCrLf & _
                      " And examcode = '" & lsExamCode & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            End If
'        End If
        
        db_Commit gLocal
        
'        If chkMode.Value = 1 Then
'            Dim sBarCode1 As String
'            Dim sExamCode1 As String
'            Dim sResult1 As String
'            Dim sEquip1 As String
'            Dim rc As String
'
'            sBarCode1 = ""
'            sExamCode1 = ""
'            sResult1 = ""
'            sEquip1 = ""
'
'            sExamCode1 = chrTAB & lsExamCode & chrTAB
'
'            sBarCode1 = chrTAB & lsID & chrTAB
'
'            sResult1 = chrTAB & Trim(lsResult) & chrTAB
'
'            sEquip1 = chrTAB & gEquip & chrTAB
'
'            If Left(Trim(GetText(vasID, gRow, colBarCode)), 1) = "9" Then
'
'            Else
'                'rc = Online_Result(sBarCode1, sExamCode1, sResult1, sEquip1, 1)
'                rc = Online_Result_New(sBarCode1, sExamCode1, sResult1, sEquip1, 1, "", gWorker_Info.WK_ID)
'            End If
'
'            If rc = "N" Then
'                SetBackColor vasID, gRow, gRow, 1, vasID.MaxCols, 202, 255, 112
'                SetText vasID, "완료", gRow, colState
'
'                SQL = " Update pat_res Set " & vbCrLf & _
'                      " sendflag = 'C' " & vbCrLf & _
'                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'                      " And equipno = '" & gEquip & "' " & vbCrLf & _
'                      " And barcode = '" & Trim(GetText(vasID, gRow, colBarCode)) & "' "
'                res = SendQuery(gLocal, SQL)
'                If res = -1 Then
'                    SaveQuery SQL
'                    Exit Sub
'                End If
'            Else
'                SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
'                SetText vasID, "실패", gRow, colState
'            End If
'        End If
        '==============================================================================================

    End If
    
    If (Mid(ResultTbl(1), 2, 1) = "M") Then
        lsResFlag = Trim(ResultTbl(3))
        lsID = Trim(GetText(vasID, gRow, colBarCode))
        lsTestID = gOrderCode
        lsResult = ""
        
        
        Select Case lsResFlag
        Case "3" 'v-max
                        
            Select Case lsTestID
            Case "2"
                lsResult = ">120"
            Case "4"
                lsResult = ">180"
            Case "7"
                lsResult = ">20"
            Case "6"
                lsResult = ">120"
            End Select
        Case "4" 'v-min
            Select Case lsTestID
           
            Case "6"
                lsResult = "<2"
            End Select
            
        End Select
        
        If Trim(lsResult) <> "" Then
            For i = 1 To UBound(gArrEquip)
                If lsTestID = gArrEquip(i, 2) Then
                    lCol = (gArrEquip(i, 1) - 1)
                    
                    lsExamCode = Trim(gArrEquip(i, 3))
                    lsExamName = Trim(gArrEquip(i, 4))
                    SetText vasID, lsResult, gRow, colResult + lCol * 4
                    SetText vasID, lsResult, gRow, colResult1 + lCol
                End If
            Next i
            
            SQL = "update pat_res set result = '" & lsResult & "' where barcode = '" & lsID & "' and equipcode = '" & lsTestID & "'"
            res = SendQuery(gLocal, SQL)
        End If
        
        
        
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "L") Then
        If gOrderMessage = "R" Then
            If MnTransAuto.Checked = True Then
                res = ToServer(gRow)
        
                If res = -1 Then
                    SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                    SetText vasID, "실패", gRow, colState
                Else
                   
                    SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                    SetText vasID, "완료", gRow, colState
                    
                    SQL = " Update pat_res Set " & vbCrLf & _
                          " sendflag = 'C' " & vbCrLf & _
                          " Where equipno = '" & gEquip & "' " & vbCrLf & _
                          " And barcode = '" & Trim(GetText(vasID, gRow, colBarCode)) & "' "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                    
                End If
            Else
                SetBackColor vasID, gRow, gRow, 1, vasID.MaxCols, 202, 255, 112
                SetText vasID, "결과", gRow, colState
            End If
            
        End If
        
'        If gOrderMessage = "R" Then
'            sOCnt = "0"
'            sRCnt = "0"
'
'            Get_Order Trim(gsBarCode)
'            sOCnt = UBound(gOrder_List)
'
'            For i = colResult1 To vasID.MaxCols
'                If Trim(GetText(vasID, gRow, i)) <> "" Then
'                    sRCnt = sRCnt + 1
'                End If
'            Next i
'
'            '결과갯수
'            SetText vasID, sOCnt, gRow, colOrd
'            SetText vasID, sRCnt, gRow, colRes
'
'            '만약 Auto 버튼이 체크되면 자동전송 되도록 함
'            If chkMode.Value = 1 Then
'                If sOCnt = sRCnt Then
'                    liRet = -1
'
'                    jRow = CInt(gRow)
'                    liRet = Insert_Data(jRow)
'
'                    If liRet = -1 Then
'                        SetForeColor vasID, gRow, gRow, colState, colState, 255, 0, 0
'                        SetText vasID, "실패", gRow, colState
'                    Else
'                        SetForeColor vasID, gRow, gRow, colState, colState, 0, 0, 0
'                        SetText vasID, "완료", gRow, colState
'
'                        SQL = " Update pat_res Set " & vbCrLf & _
'                              " sendflag = 'B' " & vbCrLf & _
'                              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'                              " And equipno = '" & gEquip & "' " & vbCrLf & _
'                              " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
'                        res = SendQuery(gLocal, SQL)
'                        If res = -1 Then
'                            SaveQuery SQL
'                            Exit Sub
'                        End If
'
'                        vasID.Row = gRow
'                        vasID.Col = 1
'                        vasID.Value = 1
'                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
'                    End If
'                End If
'            End If
'        End If
    End If
End Sub

Sub STA_R_이전(asData As String)
'    Dim ResultTbl(1 To 40) As String
'    Dim TablePtr As Integer
'    Dim sTmp As String
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim X As Integer
'
'    Dim iCnt As Integer
'
'    Dim lsID As String
'    Dim lsPID As String
'    Dim lsPName As String
'    Dim lsJumin1 As String
'    Dim lsJumin2 As String
'    Dim lsPSex As String
'    Dim lsPage As String
'
'    Dim lsTestID As String
'    Dim lsExamCode As String
'    Dim lsResult As String
'    Dim lsExamDate As String
'
'    Dim sSampleType As String
'
'    Dim rv As Integer
'    Dim vTemp As String
'
'    Dim liRet As Long
'    Dim iRow As Integer
'
'    If asData = "" Then
'        Exit Sub
'    End If
'
'
'    TablePtr = 1
'' ----- for start
'    For j = 1 To Len(asData)
'        If (Mid(asData, j, 1) = "|") Then
'            TablePtr = TablePtr + 1
'            ResultTbl(TablePtr) = " "
'        Else
'            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
'        End If
'    Next j
'' ------- for end
'
'    If Mid(ResultTbl(1), 2, 1) = "H" Then     'Header Record
'        Var_Clear
'
'        iCnt = 0
'
'        For i = 1 To Len(asData)
'            If Mid(asData, i, 1) = "|" Then
'                iCnt = iCnt + 1
'
'                Select Case iCnt
'                    Case 13
'                        gDate = Mid(asData, i + 1, 14)      '장비에서 받은 날짜시간
'                End Select
'            End If
'        Next i
'    End If
'
'    If Mid(ResultTbl(1), 2, 1) = "Q" Then     'Request Information Record
'        gOrderMessage = "Q"
'
'        sTmp = ResultTbl(3)
'        i = InStr(1, sTmp, "^")
'        sTmp = Mid(sTmp, i + 1, 11)
'        gsBarCode = sTmp
'
'
'        gRow = -1
'        For i = 1 To vasID.DataRowCnt
'            If gsBarCode <> "" Then  '메뉴얼일경우 바코드 없음
'                If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
'                    gRow = i
'                    Exit For
'                End If
'            End If
'        Next i
'
'        If gRow = -1 Then
'            gRow = vasID.DataRowCnt + 1
'            If vasID.maxrows < gRow Then
'                vasID.maxrows = gRow
'            End If
'        End If
'
'        SetText vasID, gsBarCode, gRow, colBarcode
'
'        vasActiveCell vasID, gRow, colPID
'
'        SetForeColor vasID, gRow, gRow, 1, colState, 0, 0, 0
'
'        '환자정보 불러오기
'        If Trim(GetText(vasID, gRow, colPName)) = "" Then
'            Get_Sample_Info gRow
'        End If
'
'        'Order 만들기***********************
'        res = MakeOrder(gRow, gsBarCode)
'    End If
'
'
'    If (Mid(ResultTbl(1), 2, 1) = "O") Then          'Test Order Record
'        gsBarCode = Trim(ResultTbl(3))      '검체번호
'
'        sSampleType = "P"
'
'        gRow = -1
'        For i = 1 To vasID.DataRowCnt
'            If sSampleType = "P" Then
'                If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
'                    gRow = i
'                    Exit For
'                End If
'            ElseIf sSampleType = "Q" Then
''                If Trim(GetText(vasID, i, colBarcode)) = gsBarCode And _
''                   Trim(GetText(vasID, i, colReceDate)) = gsResDateTime Then
''                    gRow = i
''                    Exit For
''                End If
'            End If
'        Next i
'
'        If gRow < 0 Then
'            gRow = vasID.DataRowCnt + 1
'            If vasID.maxrows < gRow Then
'                vasID.maxrows = gRow
'            End If
'        End If
'
'        SetText vasID, gsBarCode, gRow, colBarcode
'
'        SetText vasID, sSampleType, gRow, colSampleType
'
'        '샘플정보 가져오기
'        If sSampleType = "Q" Then
'            SetText vasID, "QC", gRow, colState
'            'SetText vasID, gsResDateTime, gRow, colReceDate
'        Else
'            '검체번호에 대한 검사목록 전체조회 (상세조회)
''            i_spc_no = gsBarCode
''
''            rv = sl_sel_spcno_tstcd_all_sub(i_spc_no, i_equip_cd, v_spc_no(), v_pt_no(), v_pt_nm(), _
''                                            v_tst_frct_cd(), v_tst_frct_nm(), v_acpt_dte(), v_acpt_no(), v_sex(), _
''                                            v_age(), v_spc_cd(), v_spc_nm(), v_tst_cd(), v_tst_nm())
''
''            If rv < 1 Then
''                SetText vasID, "", gRow, colState
''            Else
''                SetText vasID, v_pt_no(0), gRow, colPID
''                SetText vasID, v_pt_nm(0), gRow, colPName
''
''                SetText vasID, v_sex(0), gRow, colPSex
''                SetText vasID, v_age(0), gRow, colPAge
''            End If
'
'            If Trim(GetText(vasID, gRow, colPName)) = "" Then
'            '환자정보 불러오기*****************
'                Get_Sample_Info gRow
'            End If
'        End If
'    End If
'
'    If (Mid(ResultTbl(1), 2, 1) = "R") Then     'Result
'        gOrderMessage = "R"
'
'        sTmp = ResultTbl(3)
'        i = InStr(1, sTmp, "^")
'        sTmp = Mid(sTmp, i + 1)
'        i = InStr(1, sTmp, "^")
'        sTmp = Mid(sTmp, i + 1)
'        i = InStr(1, sTmp, "^")
'        sTmp = Mid(sTmp, i + 1)
'        i = InStr(1, sTmp, "^")
'        lsTestID = Trim(sTmp)           '장비코드
'
'        sTmp = ResultTbl(4)
'        lsResult = Trim(sTmp)           '결과
'
'
'        gsResDateTime = ResultTbl(13)    'result time
'
'        ClearSpread vasTemp
'
'        SQL = "Select examcode, examname From equipexam" & vbCrLf & _
'              "Where equipno = '" & gEquip & "' " & vbCrLf & _
'              "And equipcode = '" & lsTestID & "'"
'        res = db_select_Vas(gLocal, SQL, vasTemp)
'
'        If vasTemp.DataRowCnt > 0 Then
'            k = -1
'            If vasTemp.DataRowCnt > 1 And vasResTemp.DataRowCnt > 0 Then
'                For j = 1 To vasTemp.DataRowCnt
'                    k = -1
'                    For X = 1 To vasResTemp.DataRowCnt
'                        If Trim(GetText(vasResTemp, X, 1)) = Trim(GetText(vasTemp, j, 1)) Then
'                            k = j
'                            Exit For
'                        End If
'                    Next X
'                    If k > 0 Then
'                        Exit For
'                    End If
'                Next j
'            End If
'
'            If k < 1 Then
'                k = 1
'            Else
'                vasTemp.maxrows = k
'            End If
'
'            For j = k To vasTemp.DataRowCnt
'                i = i + 1
'
'                If IsNumeric(lsTestID) = True And IsNumeric(lsResult) = True Then
'                    If i > vasRes.maxrows Then
'                        vasRes.maxrows = i
'                    End If
'
'                    If i > 0 Then
'                        SetText vasRes, lsTestID, i, colEquipCode                           '장비코드
'                        SetText vasRes, Trim(GetText(vasTemp, j, 1)), i, colExamCode        '검사코드
'                        SetText vasRes, Trim(GetText(vasTemp, j, 2)), i, colExamName        '검사명
'
'                        SetText vasRes, lsResult, i, colResult                              '검사결과
'                        SetText vasRes, lsResult, i, colResult1                             '검사결과
'
'                        Save_Local_One_1 gRow, i, "A"
'                    End If
'
'                    lsExamCode = Trim(GetText(vasTemp, j, 1))
'
'                End If
'            Next j
'        End If
'    End If
'
'    If (Mid(ResultTbl(1), 2, 1) = "L") Then
'        If gOrderMessage = "R" Then
'            '결과갯수
'            SetText vasID, vasRes.DataRowCnt, gRow, colRes
'
'            '만약 Auto 버튼이 체크되면 자동전송 되도록 함
'            If chkMode.Value = 1 Then
'                liRet = -1
'
'                iRow = CInt(gRow)
'                liRet = Insert_Data(iRow)
'
'                If liRet = -1 Then
'                    SetForeColor vasID, iRow, iRow, colState, colState, 255, 0, 0
'                    SetText vasID, "실패", iRow, colState
'                Else
'                    SetForeColor vasID, iRow, iRow, colState, colState, 0, 0, 0
'                    SetText vasID, "완료", iRow, colState
'
'                    SQL = " Update pat_res Set " & vbCrLf & _
'                          " sendflag = 'B' " & vbCrLf & _
'                          " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'                          " And equipno = '" & gEquip & "' " & vbCrLf & _
'                          " And barcode = '" & Trim(GetText(vasID, iRow, colBarcode)) & "' "
'                    res = SendQuery(gLocal, SQL)
'                    If res = -1 Then
'                        SaveQuery SQL
'                        Exit Sub
'                    End If
'
'                    vasID.Row = iRow
'                    vasID.Col = 1
'                    vasID.Value = 1
'                    SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
'                End If
'            End If
'        End If
'    End If
    
End Sub

Function MakeOrder(argRow As Long, argID As String) As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim lsSpcCode As String
    Dim lsExamCode As String
    Dim lsEquipCode As String
    
    Dim sCnt As String
    Dim iCnt As String
    Dim iGet As Integer

    
    Dim rv As Integer
    Dim vTemp As String
    Dim sRv As String
    
    
    MakeOrder = -1
    gsOrder = ""
    
    iGet = 1

'2009.10.06 이상은 - 무조건 검사되게 할 것
'    sCnt = ""
'    SQL = "Select count(*) from pat_res " & vbCrLf & _
'          "where barcode = '" & argID & "'  "
'    res = db_select_Var(gLocal, SQL, sCnt)
'    If res > 0 Then
'        If Not IsNumeric(sCnt) Then
'            sCnt = "0"
'        End If
'        If CInt(sCnt) > 0 Then
'            iGet = 2
'        End If
'    ElseIf res = -1 Then
'        SaveQuery SQL
'    End If

    ClearSpread vasTemp

    If iGet = 2 Then    '재검
        SQL = "select equipcode from pat_res where barcode = '" & argID & "' and result <> '' "
        res = db_select_Vas(gLocal, SQL, vasTemp)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        If vasTemp.DataRowCnt > 0 Then
            For i = 1 To vasTemp.DataRowCnt
                If Trim(GetText(vasTemp, i, 1)) <> "" Then
                    If gsOrder = "" Then
                        gsOrder = "^^^" & Trim(GetText(vasTemp, i, 1)) & "^0"
                    Else
                        gsOrder = gsOrder & "\^^^" & Trim(GetText(vasTemp, i, 1)) & "^0"
                    End If

                End If
            Next i
            MakeOrder = 1
            SetText vasID, "재검", argRow, colState
            Exit Function
        Else
            iGet = 1
        End If
    End If
    
    If iGet = 1 Then    '검사항목 불러오기
        gAllExam = ""
        
'        rv = Get_Order(Trim(argID))

        Clear_XML_Exam
        sRv = Online_XML(gXml_S07, Trim(argID))
        
        
        If Trim(sRv) <> "" Then
            MakeOrder = 0
            
            SetText vasID, "없음", argRow, colState
            SetForeColor vasID, argRow, argRow, 2, 2, 255, 0, 0
        Else
            lsExamCode = ""

            ClearSpread vasTemp1
                    
            For i = 0 To UBound(gExam_Select)
'                vasTemp1.SetText 1, i + 1, gOrder_List(i).TST_CD
                vasTemp1.SetText 1, i + 1, gExam_Select(i).TST_CD
                
                
                If lsExamCode = "" Then
                    lsExamCode = "'" & Trim(GetText(vasTemp1, i + 1, 1)) & "'"
                Else
                    lsExamCode = lsExamCode & ",'" & Trim(GetText(vasTemp1, i + 1, 1)) & "'"
                End If
                
            Next i
    
            SaveData argID & " : " & lsExamCode
            
            If lsExamCode <> "" Then
                ClearSpread vasTemp
                
                SQL = "Select EquipCode, ExamCode, Examname, examflag from EquipExam " & _
                      "where EquipNo = '" & gEquip & "' and ExamCode in (" & lsExamCode & ")  "
                res = db_select_Vas(gLocal, SQL, vasTemp)
                
                iCnt = 0
                For i = 1 To vasTemp.DataRowCnt
'                    If Trim(GetText(vasTemp, i, 4)) <> "1" Then
                        iCnt = iCnt + 1
                        
                        If gsOrder = "" Then
                            gsOrder = "^^^" & Trim(GetText(vasTemp, i, 1))
                        Else
                            gsOrder = gsOrder & "\^^^" & Trim(GetText(vasTemp, i, 1))
                        End If
'                    End If
                Next i
            End If
    
            MakeOrder = 1

            SetText vasID, iCnt, argRow, colOrd
            SetText vasID, "오더", argRow, colState
            SetForeColor vasID, argRow, argRow, 2, 2, 0, 0, 0
        End If
    End If
    
End Function

Function Save_Local_One_1(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
'    Dim sCnt As String
'    Dim sExamDate As String
'
'    sExamDate = GetDateFull
'
'    If UCase(Left(Trim(GetText(vasID, asRow1, colPJumin)), 1)) = "F" Then
''        Save_Local_QC Trim(Text_Today.Text) & " " & Format(Time, "hh:nn:ss"), _
'                      Trim(GetText(vasID, asRow1, colBarcode)), _
'                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
'                      Trim(GetText(vasRes, asRow2, colResult)), _
'                      Trim(GetText(vasRes, asRow2, colResult1))
'        Save_Local_QC Trim(Text_Today.Text) & " " & Trim(GetText(vasID, asRow1, colPID)), _
'                      Trim(GetText(vasID, asRow1, colBarcode)), _
'                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
'                      Trim(GetText(vasRes, asRow2, colResult)), _
'                      Trim(GetText(vasRes, asRow2, colResult1))
'        Exit Function
'    End If
'
'    sCnt = ""
'    If Trim(GetText(vasRes, asRow2, colEquipCode)) = "" Then Exit Function
'
'    SQL = "select count(*) from pat_res " & vbCrLf & _
'          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and barcode = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & CStr(CLng(GetText(vasRes, asRow2, colEquipCode))) & "'" & vbCrLf & _
'          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'    res = db_select_Col(gLocal, SQL)
'    sCnt = Trim(gReadBuf(0))
'    If res = -1 Then
'        SaveQuery SQL, 1
'        Exit Function
'    End If
'
'    If Not IsNumeric(sCnt) Then
'        sCnt = "0"
'    End If
'
'    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
'        SetText vasID, "0", asRow1, colPAge
'    End If
''    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
''        SetText vasExam, "1900-01-01", asRow, colExamDate
''    End If
'
'    If sCnt = "0" Then
'        SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
'              "pid, pname, jumin, page, psex, resdate, receno, " & _
'              "equipcode, examcode, result, result1, sendflag, examname, " & _
'              "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
'              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
'              "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(vasID, asRow1, colSeqNo)) & "'," & _
'              "'" & Trim(GetText(vasID, asRow1, colRack)) & "', '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & _
'              "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
'              "'" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colPJumin)) & "', " & _
'              "'" & Trim(GetText(vasID, asRow1, colPAge)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & _
'              "'" & sExamDate & "', '" & Trim(GetText(vasID, asRow1, colReceno)) & "', " & vbCrLf & _
'              "'" & CStr(CLng(GetText(vasRes, asRow2, colEquipCode))) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "',  " & _
'              "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
'              "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "',  " & _
'              "'" & Trim(GetText(vasID, asRow1, colOrd)) & "', '" & Trim(GetText(vasID, asRow1, colRes)) & "', '" & Trim(GetText(vasID, asRow1, colDate)) & "') "
'
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'    Else
'        SQL = " Update pat_res Set " & vbCrLf & _
'              " diskno = '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
'              " posno  = '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf & _
'              " result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
'              " result1 = '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', " & vbCrLf & _
'              " refflag = '" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', " & vbCrLf & _
'              " refvalue = '" & Trim(GetText(vasID, asRow1, colOrd)) & "', " & vbCrLf & _
'              " panicvalue = '" & Trim(GetText(vasID, asRow1, colRes)) & "', " & vbCrLf & _
'              " resdate = '" & sExamDate & "' " & vbCrLf & _
'              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'              " And equipcode = '" & CStr(CLng(GetText(vasRes, asRow2, colEquipCode))) & "' " & vbCrLf & _
'              " And examcode = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "' "
'
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Function
'        End If
'
'    End If
    
End Function

Function ToQCServer(ByVal asRow As Long, Optional ByVal asSend As Integer = 0, Optional ByVal asRef As Integer = 0) As Integer
'서버의 데이타 베이스에 QC저장
'    Dim iRow As Long
'    Dim i As Integer
'
'    Dim sLotNo As String
'    Dim sExamTime As String
'
'
'    Dim lsbarcode As String
'    Dim sBarCode As String
'    Dim sBarCode1 As String
'    Dim sExamCode As String
'    Dim sResult As String
'    Dim sEquip As String
'    Dim sRefFlag As String
'
'    Dim sRet As String
'
'    Dim sErrFlag As String
'
'    Dim sRCnt As String
'
'    ToQCServer = -1
'
'    sBarCode = Trim(GetText(vasID, asRow, colBarCode))
'
'    Get_QCList Mid(sBarCode, 1, 11), 1
'
'    'Local에서 환자별로 결과값 가져오기
'    ClearSpread vasResTemp
'
'    SQL = " Select equipcode, examcode, result, resdate, pid " & vbCrLf & _
'          " From pat_res " & vbCrLf & _
'          " Where examdate = '" & Format(Trim(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & sBarCode & "' and examcode <> '' and result <> '' "
'    res = db_select_Vas(gLocal, SQL, vasResTemp)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    sLotNo = Trim(GetText(vasResTemp, 1, 5))
'
'    sExamTime = gQC_Info(0).INST_DTM
'    sExamTime = Format(sExamTime, "yyyymmddhhmmss")
'
'    lsbarcode = Mid(Trim(GetText(vasID, asRow, colBarCode)), 1, 11)
'    sEquip = gEquip
'
'    sExamCode = ""
'    sResult = ""
'
'    sRCnt = "0"
'
'    For i = 1 To vasResTemp.DataRowCnt
'        sRCnt = sRCnt + 1
'
'        If sExamCode = "" Then
'            sExamCode = chrTAB & Trim(GetText(vasResTemp, i, 2)) & chrTAB
'        Else
'            sExamCode = sExamCode & Trim(GetText(vasResTemp, i, 2)) & chrTAB
'        End If
'
'        If sResult = "" Then
'            sResult = chrTAB & Trim(GetText(vasResTemp, i, 3)) & chrTAB
'        Else
'            sResult = sResult & Trim(GetText(vasResTemp, i, 3)) & chrTAB
'        End If
'    Next i
'
'    If sRCnt = "0" Then
'        ToQCServer = 0
'    End If
'
'    sRet = Online_QCResult(lsbarcode, sEquip, sLotNo, sExamTime, sRCnt, sExamCode, sResult, gWorker_Info.WK_ID)
'
'    If sRet = "N" Then
'        ToQCServer = 1
'
'        SQL = " update pat_res set sendflag = 'C' " & vbCrLf & _
'              " Where examdate = '" & Format(Trim(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & sBarCode & "'"
'        res = SendQuery(gLocal, SQL)
'    End If
End Function

Function ToServer(ByVal argSpcRow As Long, Optional ByVal asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
        Dim sDpcd, sDate1, sSlip, sItem, sOitp, sWkno As String
    Dim sIDNo, sSmyr, sSmsn, sSms1 As String
    Dim tSmsn As String
    Dim lsExamCode, lsResult As String
    Dim lPanicLow, lPanicHigh As Currency
    Dim lDeltaLow, lDeltaHigh, lDeltaMeth, lDeltaGap
    Dim lsPanic, lsDelta As String
    Dim lsPreDate, lsPreResult As String
    Dim lsNState, lsWState As String
    Dim lStdVal
    Dim lTerm As Long
    Dim lsQCChk As String

    Dim iNone, iDP

    Dim sResDate As String
    Dim sRDate As String
    Dim sRTime As String

    Dim lsID As String

    Dim i, j As Long
    Dim lRow As Long
    Dim lsQCOn As String
    
    Dim sResult As String
    Dim sExamCode As String
    Dim sBarCode As String
    Dim sEquipCode As String
    Dim sResStr As String
    Dim sResRow As Long
    Dim sResCnt As String
    Dim sEquipRes As String
    Dim sParam As String
    Dim X As Integer
    
    ToServer = -1

    lsQCOn = ""

    lRow = argSpcRow

    If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Function

    lsID = Trim(GetText(vasID, lRow, colBarCode))
    sBarCode = ""
    sEquipCode = ""
    sResult = ""
    sExamCode = ""
    
    If lsID = "" Then Exit Function

    ClearSpread vasTemp
    ClearSpread vasTemp1

    iNone = 0
    iDP = 0
    
    gOrderExam = ""
    Online_XML gXml_S07, lsID
    
    
    SQL = "Select equipcode, examcode, examname, result, result " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & lsID & "' " & vbCrLf & _
          "  and examcode in (" & gOrderExam & ") " & vbCrLf & _
          "  and result <> '' "
    If asSend = 0 Then
'        SQL = SQL & vbCrLf & _
'          "  and sendflag <> 'C' "
    End If
    
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If vasTemp.DataRowCnt < 1 Then Exit Function

    Save_Raw_Data lsID & " : 서버 결과 전송 시작"
    Save_Raw_Data lsID & " : 장부 정보 가져오기"

    On Error GoTo ErrHandle
    
    sParam = ""
    
    For sResRow = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" Then
            sParam = sParam & "<Table>" & _
                    "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                    "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                    "<USERID><![CDATA[LIA]]></USERID>" & _
                    "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                    "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                    "<P0><![CDATA[" & lsID & "]]></P0>" & _
                    "<P1><![CDATA[" & Trim(GetText(vasTemp, sResRow, 2)) & "]]></P1>" & _
                    "<P2><![CDATA[" & Trim(GetText(vasTemp, sResRow, 5)) & "]]></P2>" & _
                    "<P3><![CDATA[]]></P3>" & _
                    "<P4><![CDATA[" & gEquip & "]]></P4>" & _
                    "<P5><![CDATA[" & gIFUser & "]]></P5>" & _
                    "<P6><![CDATA[]]></P6>" & _
                    "<P7><![CDATA[]]></P7>" & _
                    "<P8><![CDATA[]]></P8>" & _
                    "<P9><![CDATA[]]></P9>" & _
                    "</Table>"
            SQL = "Update pat_res set sendflag = 'C' " & vbCrLf & _
                  "where equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and barcode = '" & lsID & "' and examcode = '" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        End If
    Next
    
    If sParam = "" Then Exit Function
    
    sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
    
    Online_Result_Qry sParam
    
    ToServer = 1

    Save_Raw_Data lsID & " : 서버 결과 전송 완료!"

    Exit Function

ErrHandle:
    Save_Raw_Data Err.Number & " : " & Err.Description & vbCrLf & _
                  SQL
    Resume Next
End Function

Private Sub Picture1_Click()
    frmUser.Show 0
End Sub

Private Sub Text_Today_GotFocus()
    SelectFocus Text_Today
End Sub

Private Sub Text_Today_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCall_Click
    End If
End Sub

Private Sub txtBarcode_GotFocus()
    SelectFocus txtBarcode
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    Dim i As Integer
    
    If KeyCode = vbKeyReturn Then
        If txtBarcode <> "" Then
            For lRow = 1 To vasID.DataRowCnt
                If Trim(txtBarcode.Text) = Trim(GetText(vasID, lRow, colBarCode)) Then
                    i = lRow
                    Exit For
                End If
            Next lRow
        End If
        
        If i < 1 Then
            i = vasID.DataRowCnt + 1
            If i > vasID.MaxRows Then
                vasID.MaxRows = i + 1
            End If
            
            SQL = " Select receno, pid, seqno, pname from pat_res " & CR & _
                  " Where equipno = '" & gEquip & "' " & CR & _
                  " And examdate = '" & Format(Text_Today.Text, "YYYYMMDD") & "' " & CR & _
                  " And barcode= '" & Trim(txtBarcode) & "' "
            res = db_select_Col(gLocal, SQL)
            If res = 1 Then
                SetText vasID, Trim(txtBarcode), i, colBarCode
                SetText vasID, Trim(gReadBuf(0)), i, colReceno
                SetText vasID, Trim(gReadBuf(1)), i, colPID
                SetText vasID, Trim(gReadBuf(2)), i, colSeqNo
                SetText vasID, Trim(gReadBuf(3)), i, colPName
                
                vasID_Click colBarCode, i
            End If
        Else
            vasActiveCell vasID, lRow, 2
            vasID.SetFocus
        End If
        
        txtBarcode.Text = ""
    End If
End Sub

Private Sub txtEnd_GotFocus()
    SelectFocus txtEnd
End Sub

Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtEnd) = False Then
            txtEnd.SetFocus
            Exit Sub
        End If
        cmdSend.SetFocus
    End If
End Sub

Private Sub txtID_GotFocus()
    SelectFocus txtID
End Sub

Private Sub txtStart_GotFocus()
    SelectFocus txtStart
End Sub

Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtStart) = False Then
            txtStart.SetFocus
            Exit Sub
        End If
        txtEnd.SetFocus
    End If
End Sub


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        If Col = colRack Or Col = colPos Then
            vasSort vasID, colRack, colPos
        Else
            vasSort vasID, Col
        End If
    End If
    
    If Row < 0 Or Row > vasID.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasID.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim lsCnt As String
'    Dim lsID As String
'    Dim lsDate As String
'    Dim lsTime As String
'
'    Dim iRow As Long
'
'    'cmdCall_Click
'
'    If Row < 1 Or Row > vasID.DataRowCnt Then
'        Exit Sub
'    End If
'
'    lsID = Trim(GetText(vasID, Row, colBarcode))
'
'    'Local에서 불러오기
'    ClearSpread vasRes
'
'    If Trim(GetText(vasID, Row, colPJumin)) = "F" Then
'        lsTime = Trim(GetText(vasID, Row, colPID))
'        If Len(lsTime) = 4 Then
'        Else
'            lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
'        End If
'        SQL = "select a.equipcode, min(b.examcode), min(b.examname), a.result, b.seqno, a.resflag, a.result " & vbCrLf & _
'              " From qc_res a, equipexam b " & vbCrLf & _
'              "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'              "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'              "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
'              "  and a.levelname = '" & lsID & "' " & vbCrLf & _
'              "  and b.equipno = a.equipno " & vbCrLf & _
'              "  and b.equipcode = a.equipcode " & vbCrLf & _
'              "group by a.equipcode, a.result, b.seqno, a.resflag, a.result "
'        res = db_select_Vas(gLocal, SQL, vasRes)
'    End If
'
'
'    '장비코드, 검사코드, 검사명, 결과, 순번
'    SQL = "Select a.equipcode, a.examcode, b.examname, a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
'          "from pat_res a, equipexam b " & vbCrLf & _
'          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
'          "  and a.examcode <> a.equipcode " & vbCrLf & _
'          "  and b.equipno = a.equipno " & vbCrLf & _
'          "  and b.equipcode = a.equipcode " & vbCrLf & _
'          "  and b.examcode = a.examcode "
'    res = db_select_Vas(gLocal, SQL, vasRes)
'    SQL = "Select a.equipcode, a.examcode, max(b.examname), a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
'          "from pat_res a, equipexam b " & vbCrLf & _
'          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
'          "  and a.examcode = a.equipcode " & vbCrLf & _
'          "  and b.equipno = a.equipno " & vbCrLf & _
'          "  and b.equipcode = a.equipcode " & vbCrLf & _
'          "group by a.equipcode, a.examcode, a.result, b.seqno, a.refflag, a.result1 "
'    res = db_select_Vas(gLocal, SQL, vasRes, vasRes.DataRowCnt + 1, 1)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    For iRow = 1 To vasRes.DataRowCnt
'        If Trim(GetText(vasRes, iRow, colRCheck)) <> "" Then
'            SetForeColor vasRes, iRow, iRow, colResult, colResult, 255, 0, 0
'        Else
'            SetForeColor vasRes, iRow, iRow, colResult, colResult, 0, 0, 0
'        End If
'    Next iRow
'    vasRes.maxrows = vasRes.DataRowCnt
    'vasSort vasRes, 5, 2
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    
    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasID, iRow, colBarCode))
        
        If Trim(GetText(vasID, iRow, colPJumin)) = "F" Then
            If MsgBox("해당 QC 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                Exit Sub
            End If
            
            lsTime = Trim(GetText(vasID, iRow, colPID))
            If Len(lsTime) = 4 Then
            Else
                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
            End If
            
            SQL = "Delete From qc_res a " & vbCrLf & _
                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
                  "  and a.levelname = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
                
            Exit Sub
        End If
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
            
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasID, iRow, iRow
        ClearSpread vasRes
    End If
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
        vasID_DblClick colBarCode, lRow
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'Dim iRow As Long
'Dim lsID As String
'
'    If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'        Exit Sub
'    End If
'
'    iRow = Row
'
'    lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'    SQL = " Delete From pat_res " & vbCrLf & _
'          " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & lsID & "' "
'    res = SendQuery(gLocal, SQL)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    DeleteRow vasID, iRow, iRow
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim vasResRow As Long
'    Dim vasResCol As Long
'    Dim vasIDRow As Long
'
'    Dim lCCR, lM_C_ratio, lP_C_ratio As Long
'    Dim sCCR, sCrea_S, sCrea_U, sM_ALB_U, sTP_U As String
'
'    Dim sResult As String
'    Dim sResult1 As String
'
'    Dim i As Integer
'
'    Dim sTotalVol As String
'
'    Dim lsTime As String
'
'    vasIDRow = vasID.ActiveRow
'    vasResRow = vasRes.ActiveRow
'    vasResCol = vasRes.ActiveCol
'
'    If KeyCode = vbKeyReturn Then
'
'        If vasResCol = colResult Then
'
'            If Trim(GetText(vasRes, vasResRow, colEquipCode)) = "88888" Then
'                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
'                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
'                Save_Local_One_1 vasIDRow, vasResRow, "A"
'
'            ElseIf Trim(GetText(vasRes, vasResRow, colEquipCode)) = "99999" Then
'                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
'                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
'                Save_Local_One_1 vasIDRow, vasResRow, "A"
'
'                If IsNumeric(sTotalVol) Then
'                    lCCR = -1
'                    sCCR = ""
'                    sCrea_S = ""
'                    sCrea_U = ""
'                    sM_ALB_U = ""
'                    sTP_U = ""
'
'                    i = 1
'                    Do While i <= vasRes.DataRowCnt
'                        Select Case Trim(GetText(vasRes, i, colExamCode))
'                        Case "L3117", "L3101", "L3102", "L3103"  'Microalbumun(24hr),Na,K,Cl
'                            sResult = Trim(GetText(vasRes, i, colResult1))
'                            'SetText vasRes, sResult, i, colResult1
'                            If IsNumeric(sResult) Then
'                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 1000, "0.00")
'                                SetText vasRes, sResult, i, colResult
'                            End If
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'
'                        Case "L3104", "L3106", "L3107", "L3109" 'Ca,Pi,UA,Protein(24hr)
'                            sResult = Trim(GetText(vasRes, i, colResult1))
'                            'SetText vasRes, sResult, i, colResult1
'                            If IsNumeric(sResult) Then
'                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
'                                SetText vasRes, sResult, i, colResult
'                            End If
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'                        Case "L31094", "L31095" 'Protein 16hr, 8hr
'                            sResult = Trim(GetText(vasRes, i, colResult1))
'                            'SetText vasRes, sResult, i, colResult1
'                            If IsNumeric(sResult) Then
'                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
'                                SetText vasRes, sResult, i, colResult
'                            End If
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'                        Case "L31111", "L31112", "L31123", "L3113" 'Creatinie 16hr, 8hr,24hr, BUN(24hr UR)
'                            sResult = Trim(GetText(vasRes, i, colResult1))
'                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
'                            'SetText vasRes, "L31123", i, colExamCode
'                            'SetText vasRes, sResult, i, colResult1
'                            If IsNumeric(sResult) Then
'                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100 / 1000, "0.00")
'                                SetText vasRes, sResult, i, colResult
'                            End If
'
'                            Save_Local_One_1 vasIDRow, i, "A"
'                        Case "L3041", "88888"   'Serum Creatinine
'                            sCrea_S = Trim(GetText(vasRes, i, colResult1))
'
'                            'Save_Local_One_1 vasIDRow, i, "A"
'                        Case "L31121"   'CCR
'                            sCCR = Trim(GetText(vasRes, i, colResult1))
'                            lCCR = i
'                        Case "L31171"   'Microalbumin(random)
'                            sM_ALB_U = Trim(GetText(vasRes, i, colResult1))
'                        Case "L31110"  'Creatinine(random)
'                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
'                        Case "L31090"   'Protein(random)
'                            sTP_U = Trim(GetText(vasRes, i, colResult1))
'                        Case "L31172"   'Microalbumin / creatinine (random urine)
'                            lM_C_ratio = i
'                        Case "L31172"   'protein / creatinie (random)
'                            lP_C_ratio = i
'                        End Select
'                        i = i + 1
'                    Loop
'
'                    If lCCR > 0 And lCCR <= vasRes.DataRowCnt And IsNumeric(sCrea_U) = True And IsNumeric(sCrea_S) = True Then
'                        sResult = Format(CCur(sCrea_U) * CCur(sTotalVol) / 1440 / CCur(sCrea_S), "0.000")
'                        SetText vasRes, sResult, lCCR, colResult
'                        SetText vasRes, sResult, lCCR, colResult1
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
'
''                    If IsNumeric(sM_ALB_U) = True And IsNumeric(sCrea_U) = True Then
''                        sResult = Format(CCur(sM_ALB_U) / CCur(sCrea_U), "0.00") * 100
''                        If lM_C_ratio > 0 And lM_C_ratio <= vasRes.DataRowCnt Then
''                            SetText vasRes, sResult, lM_C_ratio, colResult
''                        Else
''                            i = vasRes.DataRowCnt + 1
''                            If i > vasRes.maxrows Then
''                                vasRes.maxrows = i
''                            End If
''
''                            SetText vasRes, "101", i, colEquipCode
''                            SetText vasRes, "L31172", i, colExamCode
''                            SetText vasRes, "Microalbumin / Urine Creatinine", i, colExamName
''                            SetText vasRes, sResult, i, colResult
''                            SetText vasRes, sResult, i, colResult1
''                        End If
''
''                        Save_Local_One_1 vasIDRow, i, "A"
''                    End If
''
''                    If IsNumeric(sTP_U) = True And IsNumeric(sCrea_U) = True Then
''                        sResult = Format(CCur(sTP_U) / CCur(sCrea_U), "0.00") * 1000
''                        If lP_C_ratio > 0 And lP_C_ratio <= vasRes.DataRowCnt Then
''                            SetText vasRes, sResult, lM_C_ratio, colResult
''                        Else
''                            i = vasRes.DataRowCnt + 1
''                            If i > vasRes.maxrows Then
''                                vasRes.maxrows = i
''                            End If
''
''                            SetText vasRes, "102", i, colEquipCode
''                            SetText vasRes, "L31201", i, colExamCode
''                            SetText vasRes, "Urine Protein / Urine Creatinine", i, colExamName
''                            SetText vasRes, sResult, i, colResult
''                            SetText vasRes, sResult, i, colResult1
''                        End If
''
''                        Save_Local_One_1 vasIDRow, i, "A"
''                    End If
'                End If
'            Else
'
'                If Trim(GetText(vasRes, vasIDRow, colPJumin)) = "F" Then
'
'                    If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 수정 하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'                        Exit Sub
'                    End If
'
'                    lsTime = Trim(GetText(vasID, vasIDRow, colPID))
'                    If Len(lsTime) = 4 Then
'                    Else
'                        lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
'                    End If
'
'                    SQL = "update qc_res set result = '" & sResult & "' " & vbCrLf & _
'                          "where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'                          "  and examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'                          "  and examtime = '" & lsTime & "' " & vbCrLf & _
'                          "  and levelname = '" & Trim(GetText(vasID, vasIDRow, colBarcode)) & "' " & vbCrLf & _
'                          "  and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
'                    res = SendQuery(gLocal, SQL)
'
'                    Exit Sub
'                Else
'
'
'                    sResult = Trim(GetText(vasRes, vasResRow, colResult))
'                    If MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!") = vbYes Then
'                        sResult = Trim(GetText(vasRes, vasResRow, colResult))
'
'                        SQL = " update pat_res set " & vbCrLf & _
'                              " Result = '" & sResult & "' " & vbCrLf & _
'                              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'                              " And equipno = '" & gEquip & "' " & vbCrLf & _
'                              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarcode)) & "' " & vbCrLf & _
'                              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
'                        res = SendQuery(gLocal, SQL)
'
'                        If res = -1 Then
'                            SaveQuery SQL
'                            Exit Sub
'                        End If
'
'                        'SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
'
'                    End If
'                End If
'            End If
'
'
'        End If
'    ElseIf KeyCode = vbKeyDelete Then
'        If Trim(GetText(vasID, vasIDRow, colPJumin)) = "F" Then
'
'            If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'                Exit Sub
'            End If
'
'            lsTime = Trim(GetText(vasID, vasIDRow, colPID))
'            If Len(lsTime) = 4 Then
'            Else
'                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
'            End If
'
'            SQL = "Delete From qc_res a " & vbCrLf & _
'                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
'                  "  and a.levelname = '" & Trim(GetText(vasID, vasIDRow, colBarcode)) & "' " & vbCrLf & _
'                  " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
'            res = SendQuery(gLocal, SQL)
'
'            Exit Sub
'        End If
'        If MsgBox("해당 환자의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = " Delete From pat_res " & vbCrLf & _
'              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'              " And equipno = '" & gEquip & "' " & vbCrLf & _
'              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarcode)) & "' " & vbCrLf & _
'              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' " & vbCrLf & _
'              " and examcode =  '" & Trim(GetText(vasRes, vasResRow, colExamCode)) & "' "
'        res = SendQuery(gLocal, SQL)
'
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        DeleteRow vasRes, vasResRow, vasResRow
'
'    End If
End Sub

Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
    Dim sResDateTime As String
    Dim sControl As String
    Dim sLotNo As String
    
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sRefFlag As String
    
    Dim sCnt As String
    
    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
    'sControl = Trim(Left(asBarcode, 2))
    'sLotNo = Trim(Mid(asBarcode, 3))
    sControl = asBarcode
    sRefFlag = ""
    
    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Col(gLocal, SQL)
    If res > 0 Then
        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
            If CCur(sRefHigh) < CCur(asRes2) Then
                sRefFlag = "H"
            End If
            If CCur(sRefLow) > CCur(asRes2) Then
                sRefFlag = "L"
            End If
        End If
    End If
    
    sCnt = ""
    SQL = "Select count(*) from qc_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        db_RollBack gLocal
        Exit Function
    End If
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        Exit Function
    End If
    If Not IsNumeric(sCnt) Then sCnt = "0"
    
    If CInt(sCnt) > 0 Then
        SQL = "delete from qc_res " & vbCrLf & _
              "where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
              "  and levelname = '" & sControl & "' " & vbCrLf & _
              "  and equipcode = '" & asExamCode & "' "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            'db_RollBack gLocal
            SaveQuery SQL
            Exit Function
        End If
    End If
    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        'db_RollBack gLocal
        SaveQuery SQL
        Exit Function
    End If
    
End Function


