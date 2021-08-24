VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmComSetup 
   Caption         =   "장비 설정"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11970
   WindowState     =   2  '최대화
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
      Height          =   5760
      Left            =   75
      ScaleHeight     =   5760
      ScaleWidth      =   780
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   645
      Width           =   780
   End
   Begin VB.TextBox txtEqu_Cd 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   2250
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "1234567890"
      Top             =   705
      Width           =   1005
   End
   Begin VB.TextBox txtEqu_NM 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   3645
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   705
      Width           =   3045
   End
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
      Height          =   2685
      Left            =   1140
      TabIndex        =   51
      Top             =   3705
      Width           =   10800
      Begin VB.TextBox txtRThreshold 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   8835
         TabIndex        =   30
         Text            =   "1"
         ToolTipText     =   $"frmComSetup.frx":0000
         Top             =   1875
         Width           =   945
      End
      Begin VB.TextBox txtParityReplace 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   8835
         TabIndex        =   29
         Text            =   "1"
         ToolTipText     =   "패리티 오류가 발생했을 때 데이터 스트림에서 유효하지 않은 문자를 대체하는 문자를 반환하거나 설정합니다."
         Top             =   1470
         Width           =   945
      End
      Begin VB.TextBox txtOutBufferSize 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   8835
         TabIndex        =   28
         Text            =   "1"
         ToolTipText     =   "전송 버퍼의 크기를 바이트 단위로 반환하거나 설정합니다."
         Top             =   1050
         Width           =   945
      End
      Begin VB.TextBox txtInPutLen 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   8835
         TabIndex        =   27
         Text            =   "1"
         ToolTipText     =   "Input 속성이 수신 버퍼에서 읽는 문자의 수를 반환하거나 설정합니다"
         Top             =   645
         Width           =   945
      End
      Begin VB.TextBox txtInBuf 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   8835
         TabIndex        =   26
         Text            =   "1"
         ToolTipText     =   "수신 버퍼의 크기를 바이트 단위로 반환하거나 설정합니다"
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtSThreshold 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   8835
         TabIndex        =   31
         Text            =   "1"
         ToolTipText     =   $"frmComSetup.frx":0087
         Top             =   2280
         Width           =   945
      End
      Begin VB.CheckBox EOFEnable 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "End Of File"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3930
         TabIndex        =   22
         ToolTipText     =   $"frmComSetup.frx":0122
         Top             =   1080
         Width           =   2250
      End
      Begin VB.CheckBox DTREnable 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "Data Terminal Ready "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3930
         TabIndex        =   21
         ToolTipText     =   $"frmComSetup.frx":01DD
         Top             =   675
         Width           =   2250
      End
      Begin VB.CheckBox NullDiscard 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "NullDiscard"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3930
         TabIndex        =   23
         ToolTipText     =   "널 문자가 포트에서 수신 버퍼로 전송되는지의 여부를 결정합니다"
         Top             =   1500
         Width           =   2250
      End
      Begin VB.CheckBox RTSEnable 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "Ready To Send"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3930
         TabIndex        =   24
         ToolTipText     =   $"frmComSetup.frx":026B
         Top             =   1905
         Width           =   2250
      End
      Begin VB.CheckBox chkEcho 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Caption         =   "Echo"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3930
         TabIndex        =   25
         Top             =   2295
         Width           =   2250
      End
      Begin VB.ComboBox cboPort 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   225
         Width           =   1740
      End
      Begin VB.ComboBox cboHandshaking 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   19
         Top             =   2265
         Width           =   1740
      End
      Begin VB.ComboBox cboInputMode 
         Height          =   300
         Left            =   4995
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   225
         Width           =   1950
      End
      Begin VB.ComboBox cboSpeed 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   630
         Width           =   1740
      End
      Begin VB.ComboBox cboStopBits 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   1860
         Width           =   1740
      End
      Begin VB.ComboBox cboParity 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         Top             =   1455
         Width           =   1740
      End
      Begin VB.ComboBox cboDataBits 
         Height          =   300
         Left            =   1695
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   1035
         Width           =   1740
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(0)"
         Height          =   180
         Index           =   10
         Left            =   9855
         TabIndex        =   75
         Top             =   690
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(512)"
         Height          =   180
         Index           =   9
         Left            =   9855
         TabIndex        =   74
         Top             =   1095
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(?)"
         Height          =   180
         Index           =   8
         Left            =   9855
         TabIndex        =   73
         Top             =   1515
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(1)"
         Height          =   180
         Index           =   7
         Left            =   9855
         TabIndex        =   72
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(1)"
         Height          =   180
         Index           =   2
         Left            =   9855
         TabIndex        =   71
         Top             =   2325
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(False)"
         Height          =   180
         Index           =   6
         Left            =   6330
         TabIndex        =   70
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(True)"
         Height          =   180
         Index           =   5
         Left            =   6330
         TabIndex        =   69
         Top             =   1515
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(True)"
         Height          =   180
         Index           =   4
         Left            =   6330
         TabIndex        =   68
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(False)"
         Height          =   180
         Index           =   3
         Left            =   6330
         TabIndex        =   67
         Top             =   2325
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(1024)"
         Height          =   180
         Index           =   1
         Left            =   9855
         TabIndex        =   66
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(True)"
         Height          =   180
         Index           =   0
         Left            =   6330
         TabIndex        =   65
         Top             =   690
         Width           =   540
      End
      Begin VB.Label lblInBufferSize 
         AutoSize        =   -1  'True
         Caption         =   "InBufferSize :"
         Height          =   180
         Left            =   7560
         TabIndex        =   64
         Top             =   285
         Width           =   1125
      End
      Begin VB.Label lblInputLen 
         AutoSize        =   -1  'True
         Caption         =   "InputLen :"
         Height          =   180
         Left            =   7845
         TabIndex        =   63
         Top             =   690
         Width           =   840
      End
      Begin VB.Label lblOutBufferSize 
         AutoSize        =   -1  'True
         Caption         =   "OutBufferSize :"
         Height          =   180
         Left            =   7425
         TabIndex        =   62
         Top             =   1095
         Width           =   1260
      End
      Begin VB.Label lblParityReplace 
         AutoSize        =   -1  'True
         Caption         =   "ParityReplace :"
         Height          =   180
         Left            =   7395
         TabIndex        =   61
         Top             =   1515
         Width           =   1290
      End
      Begin VB.Label lblRThreshold 
         AutoSize        =   -1  'True
         Caption         =   "RThreshold :"
         Height          =   180
         Left            =   7590
         TabIndex        =   60
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblSThreshold 
         AutoSize        =   -1  'True
         Caption         =   "SThreshold :"
         Height          =   180
         Left            =   7590
         TabIndex        =   59
         Top             =   2325
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "송수신 포트 :"
         Height          =   180
         Index           =   2
         Left            =   495
         TabIndex        =   58
         Top             =   285
         Width           =   1080
      End
      Begin VB.Label lblInputMode 
         AutoSize        =   -1  'True
         Caption         =   "InputMode :"
         Height          =   180
         Left            =   3930
         TabIndex        =   57
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label lblHandshaking 
         AutoSize        =   -1  'True
         Caption         =   "흐름 제어 :"
         Height          =   180
         Left            =   675
         TabIndex        =   56
         Top             =   2325
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "전송속도 :"
         Height          =   180
         Left            =   735
         TabIndex        =   55
         Top             =   690
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "중단 비트 :"
         Height          =   180
         Left            =   675
         TabIndex        =   54
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "패리티 :"
         Height          =   180
         Left            =   915
         TabIndex        =   53
         Top             =   1515
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "데이터 비트 :"
         Height          =   180
         Index           =   5
         Left            =   495
         TabIndex        =   52
         Top             =   1095
         Width           =   1080
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
      Height          =   2460
      Left            =   1140
      TabIndex        =   38
      Top             =   1140
      Width           =   10800
      Begin VB.TextBox txtStation 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   2490
      End
      Begin VB.TextBox txtWS_CD 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   6570
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   2385
      End
      Begin VB.TextBox txtTmp_L 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   6570
         TabIndex        =   7
         Top             =   930
         Width           =   675
      End
      Begin VB.TextBox txtTmp_H 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   7590
         TabIndex        =   8
         Top             =   930
         Width           =   675
      End
      Begin VB.TextBox txtVendor 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1710
         TabIndex        =   9
         Top             =   1335
         Width           =   2490
      End
      Begin VB.TextBox txtBuy_Tel 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   6570
         TabIndex        =   11
         Top             =   1680
         Width           =   2490
      End
      Begin VB.TextBox txtProduct 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1710
         TabIndex        =   12
         Top             =   2025
         Width           =   2490
      End
      Begin VB.TextBox txtModelNo 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   6570
         TabIndex        =   13
         Top             =   2025
         Width           =   2490
      End
      Begin VB.TextBox txtSave_DT 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   270
         IMEMode         =   8  '영문
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   3
         Top             =   585
         Width           =   1560
      End
      Begin VB.TextBox txtCharge 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   1710
         TabIndex        =   10
         Top             =   1680
         Width           =   2490
      End
      Begin MSComCtl2.DTPicker dtpBuy_DT 
         Height          =   300
         Left            =   1710
         TabIndex        =   6
         Top             =   930
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   23658496
         CurrentDate     =   37035
      End
      Begin HSCotrol.CButton cmdStation 
         Height          =   270
         Left            =   4215
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmComSetup.frx":02F9
         MaskColor       =   0
         PicCapAlign     =   1
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdWS_CD 
         Height          =   270
         Left            =   8970
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmComSetup.frx":0453
         MaskColor       =   0
         PicCapAlign     =   1
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "적정 온도 :"
         Height          =   180
         Index           =   0
         Left            =   5625
         TabIndex        =   50
         Top             =   990
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "적용 검사실 :"
         Height          =   180
         Index           =   1
         Left            =   525
         TabIndex        =   49
         Top             =   285
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "구 입 일 :"
         Height          =   180
         Index           =   2
         Left            =   825
         TabIndex        =   48
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "적용 Work Station :"
         Height          =   180
         Index           =   4
         Left            =   4950
         TabIndex        =   47
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Model No :"
         Height          =   180
         Index           =   3
         Left            =   5580
         TabIndex        =   46
         Top             =   2070
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "제 조 사 :"
         Height          =   180
         Index           =   7
         Left            =   825
         TabIndex        =   45
         Top             =   2070
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "전화 번호 :"
         Height          =   180
         Index           =   8
         Left            =   5625
         TabIndex        =   44
         Top             =   1725
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "담 당 자 :"
         Height          =   180
         Index           =   9
         Left            =   825
         TabIndex        =   43
         Top             =   1725
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "구 입 처 :"
         Height          =   180
         Index           =   10
         Left            =   825
         TabIndex        =   42
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "~"
         Height          =   180
         Left            =   7320
         TabIndex        =   41
         Top             =   990
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "결과 보관 일수 :"
         Height          =   180
         Index           =   6
         Left            =   285
         TabIndex        =   40
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "일"
         Height          =   180
         Index           =   11
         Left            =   3390
         TabIndex        =   39
         Top             =   675
         Width           =   180
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmComSetup.frx":05AD
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
      Left            =   15
      TabIndex        =   32
      Top             =   6450
      Width           =   11940
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   33
         Top             =   135
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "CButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   1
         Left            =   1515
         TabIndex        =   34
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   2
         Left            =   2895
         TabIndex        =   35
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Clear"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   3
         Left            =   4260
         TabIndex        =   36
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
   End
   Begin HSCotrol.CButton cmdEQP 
      Height          =   270
      Left            =   3270
      TabIndex        =   78
      Top             =   705
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   476
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmComSetup.frx":182F
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin HSCotrol.UserPanel pnlPoplist 
      Height          =   4710
      Left            =   7215
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   900
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
         TabIndex        =   80
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
      Left            =   1260
      TabIndex        =   77
      Top             =   750
      Width           =   1005
   End
End
Attribute VB_Name = "frmComSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const POP_EQP   As String = "EQP"
Private Const POP_STA   As String = "STA"
Private Const POP_WSC   As String = "WSC"

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
    MainForm.Caption = "LAB(" & INS_NAME & ")"
    Set objLogo = New clsLogo
    With objLogo
        .DrawingObject = picLogo
        .Caption = INS_NAME
'        .StartColor = vbBlue
'        .EndColor = vbWhite
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
        .Add "STATION", Trim(txtStation)
        .Add "WORK_ST", Trim(txtWS_CD)
        .Add "SAVE_DT", Trim(txtSave_DT)
        .Add "BUY_DT", IIf(Trim(dtpBuy_DT.Value & "") = "", "", Trim(dtpBuy_DT.Value & ""))
        .Add "TMP_LH", Trim(txtTmp_L) & vbTab & Trim(txtTmp_H)
        .Add "VENDOR", Trim(txtVendor)
        .Add "CHARGE", Trim(txtCharge)
        .Add "BUY_TEL", Trim(txtBuy_Tel)
        .Add "PRODUCT", Trim(txtProduct)
        .Add "MODEL_NO", Trim(txtModelNo)
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
    txtStation = ""
    txtWS_CD = ""
'    lblWS_Name.Caption = ""
    txtSave_DT = ""
    dtpBuy_DT.Value = Null
    txtTmp_L = ""
    txtTmp_H = ""
    txtVendor = ""
    txtCharge = ""
    txtBuy_Tel = ""
    txtProduct = ""
    txtModelNo = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEQP_Click()
    
    Dim objEqpList  As clsCommon
    Dim itemX       As ListItem
    
    Set objEqpList = New clsCommon
    With objEqpList
        Call .SetAdoCn(AdoCn_Jet)
        Set mAdoRs = .Get_InstrumentList
    End With
    Set objEqpList = Nothing
    
    With PopUp_List
        With .ColumnHeaders
            .Clear
            .Add , , "코드", (PopUp_List.Width - 310) * 0.3
            .Add , , "장비명", (PopUp_List.Width - 310) * 0.7
        End With
        .ListItems.Clear
        .tag = POP_EQP
    End With
    
    If Not mAdoRs Is Nothing Then
        Do Until mAdoRs.EOF
            Set itemX = PopUp_List.ListItems.Add(, , Trim(mAdoRs("EQP_CD") & ""))
            itemX.SubItems(1) = Trim(mAdoRs("EQP_NM") & "")
            mAdoRs.MoveNext
        Loop
                
        With pnlPoplist
            Call .Move(cmdEQP.left, cmdEQP.Top)
            .Visible = True
            .ZOrder
        End With
        PopUp_List.SetFocus
    Else
        Call ShowMessage("등록된 장비가 없습니다.")
    End If
    
    Set mAdoRs = Nothing
        
End Sub

Private Sub cmdStation_Click()
    Dim objRoom As clsCommon
    Dim itemX As ListItem
    
    Set objRoom = New clsCommon
    With objRoom
        .SetAdoCn AdoCn_SQL
        Set mAdoRs = .Get_RoomList
    End With
    Set objRoom = Nothing
    
    With PopUp_List
        With .ColumnHeaders
            .Clear
            .Add , , "코드", 0
            .Add , , "검사실", (PopUp_List.Width - 310)
        End With
        .ListItems.Clear
        .tag = POP_STA
    End With
    
    If Not mAdoRs Is Nothing Then
        Do Until mAdoRs.EOF
            Set itemX = PopUp_List.ListItems.Add(, , Trim(mAdoRs("ROOMCD") & ""))
            itemX.SubItems(1) = Trim(mAdoRs("ROOMNM") & "")
            mAdoRs.MoveNext
        Loop
                
        With pnlPoplist
            Call .Move(Frame4.left + txtStation.left, Frame4.Top + txtStation.Top)
            .Visible = True
            .ZOrder
        End With
        PopUp_List.SetFocus
    Else
        Call ShowMessage("등록된 검사실이 없습니다.")
    End If
    
    Set mAdoRs = Nothing
End Sub

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
            .Add , , "Code", 0
            .Add , , "Work Station", (PopUp_List.Width - 310)
        End With
        .ListItems.Clear
        .tag = POP_WSC
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
    
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
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
'        .StartColor = vbBlue
'        .EndColor = vbWhite
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
        With .ColumnHeaders
            .Add 1, , "검사코드", (PopUp_List.Width - 310) * 0.4
            .Add 2, , "검사항목", (PopUp_List.Width - 310) * 0.6
            .Add 3, , "타입", 0
            .Add 4, , "Unit", 0
            .Add 5, , "PrtSeq", 0
        End With
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
            txtStation = Trim(mAdoRs("STATION") & "")
            txtWS_CD = Trim(mAdoRs("WORK_ST") & "")
            txtSave_DT = Trim(mAdoRs("SAVE_DT") & "")
            dtpBuy_DT.Value = IIf(Trim(mAdoRs("BUY_DT") & "") = "", Null, Trim(mAdoRs("BUY_DT") & ""))
            strTmp_LH = Split(Trim(mAdoRs("TMP_LH") & ""), vbTab)
            txtTmp_L = strTmp_LH(0)
            txtTmp_H = strTmp_LH(1)
            txtVendor = Trim(mAdoRs("VENDOR") & "")
            txtCharge = Trim(mAdoRs("CHARGE") & "")
            txtBuy_Tel = Trim(mAdoRs("BUY_TEL") & "")
            txtProduct = Trim(mAdoRs("PRODUCT") & "")
            txtModelNo = Trim(mAdoRs("MODEL_NO") & "")

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

Private Sub pnlPoplist_CloseMe()
    pnlPoplist.Visible = False
End Sub

Private Sub PopUp_List_DblClick()
    Dim itemX As ListItem
    Set itemX = PopUp_List.SelectedItem
    
    If itemX Is Nothing Then Exit Sub
            
    Select Case PopUp_List.tag
        Case POP_EQP
            txtEqu_Cd = Trim(itemX.Text)
            Call GetEqp_Setting(txtEqu_Cd)
            txtEqu_NM = Trim(itemX.SubItems(1))
        Case POP_STA
            txtStation = Trim(itemX.SubItems(1))
        Case POP_WSC
            txtWS_CD = Trim(itemX.SubItems(1))
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

Private Sub txtBuy_Tel_GotFocus()
    Call TextBoxs_GotFocus(txtBuy_Tel)
End Sub

Private Sub txtCharge_GotFocus()
    Call TextBoxs_GotFocus(txtCharge)
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

Private Sub txtModelNo_GotFocus()
    Call TextBoxs_GotFocus(txtModelNo)
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

Private Sub txtProduct_GotFocus()
    Call TextBoxs_GotFocus(txtProduct)
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

Private Sub txtStation_GotFocus()
    Call TextBoxs_GotFocus(txtStation)
End Sub

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

Private Sub txtTmp_H_GotFocus()
    Call TextBoxs_GotFocus(txtTmp_H)
End Sub

Private Sub txtTmp_H_KeyDown(KeyCode As Integer, Shift As Integer)
    txtTmp_H.IMEMode = 8
End Sub

Private Sub txtTmp_H_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtTmp_L_GotFocus()
    Call TextBoxs_GotFocus(txtTmp_L)
End Sub

Private Sub txtTmp_L_KeyDown(KeyCode As Integer, Shift As Integer)
    txtTmp_L.IMEMode = 8
End Sub

Private Sub txtTmp_L_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
        Exit Sub
    End If

    If (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtVendor_GotFocus()
    Call TextBoxs_GotFocus(txtVendor)
End Sub

Private Sub txtWS_CD_GotFocus()
    Call TextBoxs_GotFocus(txtWS_CD)
End Sub
