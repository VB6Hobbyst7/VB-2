VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmComSetup 
   BorderStyle     =   1  '단일 고정
   Caption         =   "장비 설정"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   FillStyle       =   2  '수평선
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9165
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame5 
      Caption         =   " Communication Information "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   945
      TabIndex        =   15
      Top             =   2085
      Width           =   7935
      Begin VB.TextBox txtRThreshold 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   1965
         TabIndex        =   33
         Text            =   "1"
         ToolTipText     =   $"frmComSetup.frx":0000
         Top             =   3465
         Width           =   945
      End
      Begin VB.TextBox txtParityReplace 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         Left            =   5595
         TabIndex        =   32
         Text            =   "1"
         ToolTipText     =   "패리티 오류가 발생했을 때 데이터 스트림에서 유효하지 않은 문자를 대체하는 문자를 반환하거나 설정합니다."
         Top             =   4260
         Width           =   945
      End
      Begin VB.TextBox txtOutBufferSize 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   5595
         TabIndex        =   31
         Text            =   "1"
         ToolTipText     =   "전송 버퍼의 크기를 바이트 단위로 반환하거나 설정합니다."
         Top             =   3855
         Width           =   945
      End
      Begin VB.TextBox txtInBuf 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   1965
         TabIndex        =   29
         Text            =   "1"
         ToolTipText     =   "수신 버퍼의 크기를 바이트 단위로 반환하거나 설정합니다"
         Top             =   4290
         Width           =   945
      End
      Begin VB.TextBox txtSThreshold 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   1965
         TabIndex        =   28
         Text            =   "1"
         ToolTipText     =   $"frmComSetup.frx":0087
         Top             =   3885
         Width           =   945
      End
      Begin VB.CheckBox EOFEnable 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "End Of File"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4125
         TabIndex        =   27
         ToolTipText     =   $"frmComSetup.frx":0122
         Top             =   1440
         Width           =   2250
      End
      Begin VB.CheckBox NullDiscard 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "NullDiscard"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4125
         TabIndex        =   25
         ToolTipText     =   "널 문자가 포트에서 수신 버퍼로 전송되는지의 여부를 결정합니다"
         Top             =   1920
         Width           =   2250
      End
      Begin VB.CheckBox RTSEnable 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "Ready To Send"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4125
         TabIndex        =   24
         ToolTipText     =   $"frmComSetup.frx":01DD
         Top             =   2415
         Width           =   2250
      End
      Begin VB.CheckBox chkEcho 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "Echo"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4125
         TabIndex        =   23
         Top             =   2895
         Width           =   2250
      End
      Begin VB.ComboBox cboPort 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   22
         Top             =   435
         Width           =   1740
      End
      Begin VB.ComboBox cboHandshaking 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   2865
         Width           =   1740
      End
      Begin VB.ComboBox cboInputMode 
         Height          =   300
         ItemData        =   "frmComSetup.frx":026B
         Left            =   5190
         List            =   "frmComSetup.frx":026D
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   405
         Width           =   1950
      End
      Begin VB.ComboBox cboSpeed 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   19
         Top             =   930
         Width           =   1740
      End
      Begin VB.ComboBox cboStopBits 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   2370
         Width           =   1740
      End
      Begin VB.ComboBox cboParity 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         Top             =   1875
         Width           =   1740
      End
      Begin VB.ComboBox cboDataBits 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   1395
         Width           =   1740
      End
      Begin VB.CheckBox DTREnable 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "Data Terminal Ready "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4125
         TabIndex        =   26
         ToolTipText     =   $"frmComSetup.frx":026F
         Top             =   975
         Width           =   2250
      End
      Begin VB.TextBox txtInPutLen 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   5580
         TabIndex        =   30
         Text            =   "1"
         ToolTipText     =   "Input 속성이 수신 버퍼에서 읽는 문자의 수를 반환하거나 설정합니다"
         Top             =   3450
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(0)"
         Height          =   180
         Index           =   10
         Left            =   6630
         TabIndex        =   57
         Top             =   3510
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(512)"
         Height          =   180
         Index           =   9
         Left            =   6615
         TabIndex        =   56
         Top             =   3900
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(?)"
         Height          =   180
         Index           =   8
         Left            =   6645
         TabIndex        =   55
         Top             =   4305
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(1)"
         Height          =   180
         Index           =   7
         Left            =   3000
         TabIndex        =   54
         Top             =   3510
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(1)"
         Height          =   180
         Index           =   2
         Left            =   3000
         TabIndex        =   53
         Top             =   3930
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(False)"
         Height          =   180
         Index           =   6
         Left            =   6525
         TabIndex        =   52
         Top             =   1455
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(True)"
         Height          =   180
         Index           =   5
         Left            =   6525
         TabIndex        =   51
         Top             =   1935
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(True)"
         Height          =   180
         Index           =   4
         Left            =   6525
         TabIndex        =   50
         Top             =   2430
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(False)"
         Height          =   180
         Index           =   3
         Left            =   6525
         TabIndex        =   49
         Top             =   2910
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(1024)"
         Height          =   180
         Index           =   1
         Left            =   2955
         TabIndex        =   48
         Top             =   4335
         Width           =   510
      End
      Begin VB.Label lblInBufferSize 
         AutoSize        =   -1  'True
         Caption         =   "InBufferSize :"
         Height          =   180
         Left            =   720
         TabIndex        =   46
         Top             =   4335
         Width           =   1125
      End
      Begin VB.Label lblInputLen 
         AutoSize        =   -1  'True
         Caption         =   "InputLen :"
         Height          =   180
         Left            =   4620
         TabIndex        =   45
         Top             =   3495
         Width           =   840
      End
      Begin VB.Label lblOutBufferSize 
         AutoSize        =   -1  'True
         Caption         =   "OutBufferSize :"
         Height          =   180
         Left            =   4200
         TabIndex        =   44
         Top             =   3900
         Width           =   1260
      End
      Begin VB.Label lblParityReplace 
         AutoSize        =   -1  'True
         Caption         =   "ParityReplace :"
         Height          =   180
         Left            =   4170
         TabIndex        =   43
         Top             =   4305
         Width           =   1290
      End
      Begin VB.Label lblRThreshold 
         AutoSize        =   -1  'True
         Caption         =   "RThreshold :"
         Height          =   180
         Left            =   750
         TabIndex        =   42
         Top             =   3510
         Width           =   1095
      End
      Begin VB.Label lblSThreshold 
         AutoSize        =   -1  'True
         Caption         =   "SThreshold :"
         Height          =   180
         Left            =   750
         TabIndex        =   41
         Top             =   3930
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "송수신 포트 :"
         Height          =   180
         Index           =   2
         Left            =   495
         TabIndex        =   40
         Top             =   495
         Width           =   1080
      End
      Begin VB.Label lblInputMode 
         AutoSize        =   -1  'True
         Caption         =   "InputMode :"
         Height          =   180
         Left            =   4110
         TabIndex        =   39
         Top             =   465
         Width           =   1005
      End
      Begin VB.Label lblHandshaking 
         AutoSize        =   -1  'True
         Caption         =   "흐름 제어 :"
         Height          =   180
         Left            =   675
         TabIndex        =   38
         Top             =   2925
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "전송속도 :"
         Height          =   180
         Left            =   735
         TabIndex        =   37
         Top             =   990
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "중단 비트 :"
         Height          =   180
         Left            =   675
         TabIndex        =   36
         Top             =   2430
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "패리티 :"
         Height          =   180
         Left            =   915
         TabIndex        =   35
         Top             =   1935
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "데이터 비트 :"
         Height          =   180
         Index           =   5
         Left            =   495
         TabIndex        =   34
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(True)"
         Height          =   180
         Index           =   0
         Left            =   6525
         TabIndex        =   47
         Top             =   990
         Width           =   540
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " General Information "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   960
      TabIndex        =   9
      Top             =   750
      Width           =   7905
      Begin VB.TextBox txtEqu_NM 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   3045
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   360
         Width           =   2985
      End
      Begin VB.TextBox txtEqu_Cd 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1545
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   58
         Top             =   360
         Width           =   1005
      End
      Begin VB.TextBox txtWS_CD 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1965
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   765
         Width           =   2385
      End
      Begin VB.TextBox txtSave_DT 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   6570
         MaxLength       =   10
         TabIndex        =   10
         Top             =   765
         Width           =   840
      End
      Begin BHButton.BHImageButton cmdWS_CD 
         Height          =   285
         Left            =   4410
         TabIndex        =   12
         Top             =   765
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmComSetup.frx":02FD
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdEQP 
         Height          =   285
         Left            =   2610
         TabIndex        =   60
         Top             =   345
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmComSetup.frx":0457
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdEqpDel 
         Height          =   330
         Left            =   6975
         TabIndex        =   63
         Top             =   330
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   582
         Caption         =   "삭제"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdEqpAdd 
         Height          =   330
         Left            =   6150
         TabIndex        =   64
         Top             =   330
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   582
         Caption         =   "추가"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label7 
         Caption         =   "일"
         Height          =   195
         Left            =   7500
         TabIndex        =   62
         Top             =   825
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "장비코드 : "
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
         Left            =   540
         TabIndex        =   61
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "적용 검사 파트 :"
         Height          =   180
         Index           =   4
         Left            =   540
         TabIndex        =   14
         Top             =   825
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "결과 보관 일수 :"
         Height          =   180
         Index           =   6
         Left            =   5160
         TabIndex        =   13
         Top             =   825
         Width           =   1320
      End
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   6015
      Left            =   135
      ScaleHeight     =   6015
      ScaleWidth      =   675
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   825
      Width           =   675
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmComSetup.frx":05B1
      Caption         =   " Instruments Information"
      SubCaption      =   "장비 정보를 설정 합니다."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Lock_False 
         Height          =   300
         Left            =   8850
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   30
      TabIndex        =   0
      Top             =   7035
      Width           =   9105
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "CButton"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   1380
         TabIndex        =   6
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Save"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   2670
         TabIndex        =   7
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   3960
         TabIndex        =   8
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
   End
   Begin HSCotrol.UserPanel pnlPoplist 
      Height          =   4710
      Left            =   -30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   8308
      Bevel           =   2
      Moveble         =   -1  'True
      CloseEnabled    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ListView lvwPoplist 
         Height          =   4350
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   285
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   7673
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmComSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const POP_EQP   As String = "EQP"
Private Const POP_WSC   As String = "WSC"
'Private POP_STA   As String

Private iFlow           As Integer
Private iTempEcho       As Boolean
Private mAdoRs          As ADODB.Recordset

Private WithEvents PopUp_List As Listview
Attribute PopUp_List.VB_VarHelpID = -1

Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case 0
        Case 1
            Call cmdSave_Click
        Case 2
            Call cmdClear_Click
        Case 3 'cmd close
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub cmdSave_Click()
    Dim Eqp_Property    As Scripting.Dictionary
    Dim objEqp_Property As clsCommon
    Dim msgRst          As VbMsgBoxResult
    Dim objLogo         As clsLogo
    
    If Trim(txtEqu_Cd) = "" Then
        Call ShowMessage(" 장비 코드가 없습니다.")
        txtEqu_Cd.SetFocus
        Exit Sub
    End If
    
    INS_CODE = Trim(txtEqu_Cd)
    INS_NAME = Trim(txtEqu_NM)
    
    MainForm.Caption = INS_NAME
    
    Set objLogo = New clsLogo
    With objLogo
        .DrawingObject = picLogo
        .Caption = INS_NAME
        .StartColor = vbActiveTitleBar
        .EndColor = vbWhite
        .Draw
    End With
    Set objLogo = Nothing
    
    Call SaveString(HKEY_CURRENT_USER, REG_POSITION, REG_EQPCODE, Trim(txtEqu_Cd))
    Call SaveString(HKEY_CURRENT_USER, REG_POSITION, REG_EQPNAME, Trim(txtEqu_NM))
    
    Set Eqp_Property = New Scripting.Dictionary
    
    With Eqp_Property
        .Add "EQP_CD", Trim(txtEqu_Cd)
        .Add "CALL_ID", "" 'EQP_CALL
        .Add "EQP_NM", Trim(txtEqu_NM)
        .Add "WORK_ST", Trim(txtWS_CD)
        .Add "SAVE_DT", Trim(txtSave_DT)
        .Add "COM_PORT", cboPort.ListIndex + 1
        .Add "COM_SPEED", cboSpeed.Text
        .Add "COM_DATABIT", cboDataBits.Text
        Select Case cboParity.ListIndex
            Case 0
                .Add "COM_PARITYBIT", "e"
            Case 1
                .Add "COM_PARITYBIT", "m"
            Case 2
                .Add "COM_PARITYBIT", "n"
            Case 3
                .Add "COM_PARITYBIT", "o"
            Case 4
                .Add "COM_PARITYBIT", "s"
        End Select
        .Add "COM_STOPBIT", cboStopBits.Text
        .Add "COM_HANDSHAK", cboHandshaking.ListIndex
        .Add "COM_INPUTMOD", cboInputMode.ListIndex
        .Add "COM_DTR", DTREnable.Value
        .Add "COM_EOF", EOFEnable.Value
        .Add "COM_NULDIS", NullDiscard.Value
        .Add "COM_RTS", RTSEnable.Value
        .Add "COM_ECHO", chkEcho.Value
        .Add "COM_IBS", txtInBuf.Text
        .Add "COM_INLEN", txtInPutLen.Text
        .Add "COM_OBS", txtOutBufferSize.Text
        .Add "COM_PTR", txtParityReplace.Text
        .Add "COM_RTH", txtRThreshold.Text
        .Add "COM_STH", txtSThreshold
    End With
    
    Set objEqp_Property = New clsCommon
    
    With objEqp_Property
        .SetAdoCn AdoCn_Jet
        If Not .Let_EqpProperty(Eqp_Property) Then
            Call ShowMessage("오류가있어 저장 하지 못했습니다.")
        End If
    End With
    
    Set Eqp_Property = Nothing
    Set objEqp_Property = Nothing
    
End Sub

Private Sub cmdClear_Click()
    txtEqu_Cd = ""
    txtEqu_NM = ""

    txtWS_CD = ""
'    lblWS_Name.Caption = ""
    txtSave_DT = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEQP_Click()
    
    Dim objEqpList  As clsCommon
    Dim itemX       As ListItem
    Dim vntRs       As Variant
    Dim intRow      As Integer
    
    Set objEqpList = New clsCommon
    With objEqpList
        Call .SetAdoCn(AdoCn_SQL)
        vntRs = .Get_InstrumentList
    End With
    Set objEqpList = Nothing
    
    With PopUp_List
        With .ColumnHeaders
            .Clear
            .Add , , "장비코드", (PopUp_List.Width - 310) * 0.3
            .Add , , "장비명", (PopUp_List.Width - 310) * 0.7
        End With
        .ListItems.Clear
        .Tag = POP_EQP
    End With
    
    If IsNull(vntRs) = False Then
        For intRow = 0 To UBound(vntRs, 2)
            Set itemX = PopUp_List.ListItems.Add(, , Trim(vntRs(0, intRow) & ""))
            itemX.SubItems(1) = Trim(vntRs(1, intRow) & "")
        Next intRow
                
        With pnlPoplist
            Call .Move(cmdEQP.left + 750, cmdEQP.Top + 750)
            .Visible = True
            .ZOrder
        End With
        PopUp_List.SetFocus
    Else
        Call ShowMessage("등록된 장비가 없습니다.")
    End If
      
End Sub

Private Sub cmdEqpAdd_Click()
  
    Dim PartCode As String

        PartCode = Mid(txtWS_CD.Text, InStr(txtWS_CD.Text, "(") + 1)
        PartCode = Replace(PartCode, ")", "")
 
        If cmdEqpAdd_Insert(txtEqu_Cd.Text, txtEqu_NM.Text, PartCode) = True Then
            MsgBox "정상적으로 처리 되었습니다.", vbInformation, INS_NAME
        Else
            MsgBox "이미 사용중인 장비코드 입니다.", vbInformation, INS_NAME
        End If

End Sub

Private Function cmdEqpAdd_Insert(ByVal MachineCode As String, ByVal MachineName As String, ByVal PartCode As String) As Boolean
'Sub cmdEqpAdd_Insert(ByVal MachineCode As String, ByVal MachineName As String, ByVal PartCode As String)
    
    On Error GoTo frmComSetup_Error
    
    Dim strSql   As String, strSql1   As String

        strSql = "INSERT LabMachineSource (MachineCode, MachineName, MaxRack, S_Position, E_Position, UseYN, ComPort, isRackNumeric) " & _
                 "VALUES('" & Trim(MachineCode) & "', '" & Trim(MachineName) & "', '', '1', '999', '0', '1', '1')"
    
        AdoCn_SQL.Execute strSql
        
        strSql1 = "INSERT LabPartMachine (PartCode, MachineCode) " & _
                  "VALUES('" & Trim(PartCode) & "', '" & Trim(MachineCode) & "')"
                 
        AdoCn_SQL.Execute strSql1
        
        cmdEqpAdd_Insert = True
        
Exit Function

frmComSetup_Error:

    cmdEqpAdd_Insert = False
'    Call ErrMsgProc("frmTestEqp - cmdEqpAdd_Click()")

End Function

Private Sub cmdEqpDel_Click()

    Dim PartCode As String

        PartCode = Mid(txtWS_CD.Text, InStr(txtWS_CD.Text, "(") + 1)
        PartCode = Replace(PartCode, ")", "")
        
'        Call cmdEqpDel_DELETE(txtEqu_Cd.Text, PartCode)
        
        If cmdEqpDel_DELETE(txtEqu_Cd.Text, PartCode) = True Then
    
            txtEqu_Cd = ""
            txtEqu_NM = ""
            txtWS_CD = ""
            
            MsgBox "정상적으로 처리 되었습니다.", vbInformation, INS_NAME
        Else
            MsgBox "삭제할 코드가 없습니다.", vbInformation, INS_NAME
        End If
        
End Sub

Private Function cmdEqpDel_DELETE(ByVal MachineCode As String, ByVal PartCode As String) As Boolean
'Sub cmdEqpDel_DELETE(ByVal MachineCode As String, ByVal PartCode As String)
    
    On Error GoTo cmdEqpDel_DELETE_Error
    
    Dim strSql   As String, strSql1   As String
    Dim Call_Procedure      As String

        strSql = "DELETE FROM LabPartMachine" & _
                 " WHERE PartCode = '" & Trim(PartCode) & "'" & _
                 " AND MachineCode = '" & Trim(MachineCode) & "'"

        AdoCn_SQL.Execute strSql
        
        strSql1 = "DELETE FROM LabMachineSource" & _
                 " WHERE MachineCode = '" & Trim(MachineCode) & "'"
        
        AdoCn_SQL.Execute strSql1
        
        cmdEqpDel_DELETE = True
    
Exit Function
    
cmdEqpDel_DELETE_Error:
    
    cmdEqpDel_DELETE = False
'    Call ErrMsgProc("frmTestEqp - Private Sub ccmdEqpDel_DELETE()")
    
End Function

Private Sub cmdWS_CD_Click()
    Dim objStation As clsCommon
    Dim itemX As ListItem
    
    Set objStation = New clsCommon
    With objStation
        Call .SetAdoCn(AdoCn_SQL)
        Set mAdoRs = .Get_StationList
    End With
    Set objStation = Nothing
    
    With PopUp_List
        With .ColumnHeaders
            .Clear
            .Add , , "Code", (PopUp_List.Width - 310) * 0.3
            .Add , , "Work Station", (PopUp_List.Width - 310) * 0.7
        End With
        .ListItems.Clear
        .Tag = POP_WSC
    End With
    
    If Not mAdoRs Is Nothing Then
        Do Until mAdoRs.EOF
            Set itemX = PopUp_List.ListItems.Add(, , Trim(mAdoRs("STACD") & ""))
            itemX.SubItems(1) = Trim(mAdoRs("STANM") & "")
            mAdoRs.MoveNext
        Loop
                
        With pnlPoplist
            Call .Move(Frame4.left + txtWS_CD.left, Frame4.Top + txtWS_CD.Top)
            .Visible = True
            .ZOrder
        End With
        PopUp_List.SetFocus
    Else
        Call ShowMessage("등록된 Work Station 없습니다.")
    End If
    
    Set mAdoRs = Nothing
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
'    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
End Sub

Sub Property_Settings()
    Dim i As Integer
    
    '포트 설정을 불러옵니다.
    cboPort.Clear
    For i = 1 To 16
        cboPort.AddItem "Com" & Trim$(Str$(i))
    Next i
    
    With cboSpeed
        .Clear
        .AddItem "110"
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "28800"
        .AddItem "38400"
        .AddItem "56000"
        .AddItem "57600"
        .AddItem "115200"
        .AddItem "128000"
        .AddItem "256000"
    End With
    
    '데이터 비트 설정을 불러옵니다.
    With cboDataBits
        .Clear
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
    End With
    
    '패리티 설정을 불러옵니다.
    With cboParity
        .Clear
        .AddItem "Even"
        .AddItem "Odd"
        .AddItem "None"
        .AddItem "Mark"
        .AddItem "Space"
    End With
    
    '중단 비트 설정을 불러옵니다.
    With cboStopBits
        .Clear
        .AddItem "1"
        .AddItem "1.5"
        .AddItem "2"
    End With
    
    'cboHandshaking
    With cboHandshaking
        .Clear
        .AddItem "comNone", 0           '초기 접속 신호 없음
        .AddItem "comXonXoff", 1        'XOn/XOff 초기 접속 신호
        .AddItem "comRTS", 2            'RTS/CTS 초기 접속 신호
        .AddItem "comRTSXOnXOff", 3     'RTS와 Xon/XOff 초기 접속 신호
    End With
    
    'InputMode
    With cboInputMode
        .Clear
        .AddItem "comInputModeText", 0      '(기본값)데이터가 Input 속성을 통해서 텍스트로 변환됩니다.
        .AddItem "comInputModeBinary", 1    '데이터가 Input 속성을 통해서 이진 데이터로 변환됩니다.
    End With
    
    '기본 설정을 지정합니다.
    cboPort.ListIndex = 0
    cboSpeed.Text = "9600"
    cboParity.ListIndex = 2
    cboDataBits.Text = "8"
    cboStopBits.Text = "1"
    cboHandshaking.ListIndex = 0
    cboInputMode.ListIndex = 1
    DTREnable.Value = 1
    EOFEnable.Value = 0
    txtInBuf.Text = 1024
    txtInPutLen.Text = 0
    NullDiscard.Value = 1
    txtOutBufferSize.Text = 512
    txtParityReplace.Text = "?"
    txtRThreshold.Text = 1
    RTSEnable.Value = 1
    txtSThreshold.Text = 1
End Sub

Private Sub Form_Load()

    Dim objLogo     As New clsLogo
    With objLogo
        .DrawingObject = picLogo
        .Caption = INS_NAME
        .StartColor = vbActiveTitleBar
        .EndColor = vbWhite
        .Draw
    End With
    
    Call cmdClear_Click
    Call Property_Settings
    Call GetEqp_Setting(INS_CODE)
    
    Set PopUp_List = lvwPoplist

    With PopUp_List
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With
End Sub

Private Sub GetEqp_Setting(EqpCD As String)
    Dim strTmp_LH() As String
    Dim objEqp_Property As clsCommon
    
    Set objEqp_Property = New clsCommon
    With objEqp_Property
        Call .SetAdoCn(AdoCn_Jet)
        Set mAdoRs = .Get_EqpProperty(EqpCD)
    End With
    Set objEqp_Property = Nothing

    If Not mAdoRs Is Nothing Then
        If Not mAdoRs.EOF Then
            CaptionBar1.Caption = Trim(mAdoRs.Fields("EQP_NM") & "") & " Instruments Information"
            txtEqu_Cd = Trim(mAdoRs("EQP_CD") & "")
            txtEqu_NM = Trim(mAdoRs("EQP_NM") & "")
            txtWS_CD = Trim(mAdoRs("WORK_ST") & "")
            txtSave_DT = Trim(mAdoRs("SAVE_DT") & "")
            
            cboPort.ListIndex = Val(Trim(mAdoRs("COM_PORT") & "")) - 1
            cboSpeed.Text = Trim(mAdoRs("COM_SPEED") & "")
            cboDataBits.Text = Trim(mAdoRs("COM_DATABIT") & "")
            Select Case Trim(mAdoRs("COM_PARITYBIT") & "")
                Case "e"
                    cboParity.ListIndex = 0
                Case "m"
                    cboParity.ListIndex = 1
                Case "n"
                    cboParity.ListIndex = 2
                Case "o"
                    cboParity.ListIndex = 3
                Case "s"
                    cboParity.ListIndex = 4
                Case Else
                    cboParity.ListIndex = 2
            End Select
            cboStopBits.Text = Trim(mAdoRs("COM_STOPBIT") & "")
            cboHandshaking.ListIndex = Trim(mAdoRs("COM_HANDSHAK") & "")
            cboInputMode.ListIndex = Trim(mAdoRs("COM_INPUTMOD") & "")
            DTREnable.Value = Trim(mAdoRs("COM_DTR") & "")
            EOFEnable.Value = Trim(mAdoRs("COM_EOF") & "")
            NullDiscard.Value = Trim(mAdoRs("COM_NULDIS") & "")
            RTSEnable.Value = Trim(mAdoRs("COM_RTS") & "")
            txtInBuf.Text = Trim(mAdoRs("COM_IBS") & "")
            txtInPutLen.Text = Trim(mAdoRs("COM_INLEN") & "")
            txtOutBufferSize.Text = Trim(mAdoRs("COM_OBS") & "")
            txtParityReplace.Text = Trim(mAdoRs("COM_PTR") & "")
            txtRThreshold.Text = Trim(mAdoRs("COM_RTH") & "")
            txtSThreshold.Text = Trim(mAdoRs("COM_STH") & "")
        End If
    End If
    Set mAdoRs = Nothing
End Sub

Private Sub Lock_False_DblClick()
    If txtEqu_Cd.Locked = True Then
        txtEqu_Cd.Locked = False
        txtEqu_NM.Locked = False
        cmdEqpAdd.Visible = True
        cmdEqpDel.Visible = True
    Else
        txtEqu_Cd.Locked = True
        txtEqu_NM.Locked = True
        cmdEqpAdd.Visible = False
        cmdEqpDel.Visible = False
    End If
End Sub

Private Sub pnlPoplist_CloseMe()
    pnlPoplist.Visible = False
End Sub

Private Sub PopUp_List_DblClick()
    Dim itemX As ListItem
    Set itemX = PopUp_List.SelectedItem
    
    If itemX Is Nothing Then Exit Sub
            
    Select Case PopUp_List.Tag
        Case POP_EQP
            txtEqu_Cd = Trim(itemX.Text)
            Call GetEqp_Setting(txtEqu_Cd)
            txtEqu_NM = Trim(itemX.SubItems(1))
'        Case POP_STA
'            txtStation = Trim(itemX.SubItems(1))
        Case POP_WSC
            txtWS_CD = Trim(itemX.SubItems(1)) & "(" & Trim(itemX.Text) & ")"
    End Select
    
    Set itemX = Nothing
    
    Call pnlPoplist_CloseMe
End Sub

Private Sub PopUp_List_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PopUp_List_DblClick
    End If
End Sub

Private Sub txtEqu_Cd_Change()
    txtEqu_NM = ""
End Sub

Private Sub txtEqu_Cd_GotFocus()
    Call TextBoxs_GotFocus(txtEqu_Cd)
End Sub

Private Sub txtInBuf_GotFocus()
    Call TextBoxs_GotFocus(txtInBuf)
End Sub

Private Sub txtInBuf_KeyDown(KeyCode As Integer, Shift As Integer)
    txtInBuf.IMEMode = 8
End Sub

Private Sub txtInBuf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtInPutLen_GotFocus()
    Call TextBoxs_GotFocus(txtInPutLen)
End Sub

Private Sub txtInPutLen_KeyDown(KeyCode As Integer, Shift As Integer)
    txtInPutLen.IMEMode = 8
End Sub

Private Sub txtInPutLen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtOutBufferSize_GotFocus()
    Call TextBoxs_GotFocus(txtOutBufferSize)
End Sub

Private Sub txtOutBufferSize_KeyDown(KeyCode As Integer, Shift As Integer)
    txtOutBufferSize.IMEMode = 8
End Sub

Private Sub txtOutBufferSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtParityReplace_GotFocus()
    Call TextBoxs_GotFocus(txtParityReplace)
End Sub

Private Sub txtRThreshold_GotFocus()
    Call TextBoxs_GotFocus(txtRThreshold)
End Sub

Private Sub txtRThreshold_KeyDown(KeyCode As Integer, Shift As Integer)
    txtRThreshold.IMEMode = 8
End Sub

Private Sub txtRThreshold_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtSave_DT_GotFocus()
    Call TextBoxs_GotFocus(txtSave_DT)
End Sub

Private Sub txtSave_DT_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSave_DT.IMEMode = 8
End Sub

Private Sub txtSave_DT_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
End Sub

'Private Sub txtStation_GotFocus()
'    Call TextBoxs_GotFocus(txtStation)
'End Sub

Private Sub txtSThreshold_GotFocus()
    Call TextBoxs_GotFocus(txtSThreshold)
End Sub

Private Sub txtSThreshold_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSThreshold.IMEMode = 8
End Sub

Private Sub txtSThreshold_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtWS_CD_GotFocus()
    Call TextBoxs_GotFocus(txtWS_CD)
End Sub
