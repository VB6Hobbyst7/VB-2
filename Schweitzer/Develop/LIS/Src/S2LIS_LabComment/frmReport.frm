VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmReport 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "결과 보고"
   ClientHeight    =   9075
   ClientLeft      =   5055
   ClientTop       =   330
   ClientWidth     =   10080
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Tag             =   "ResultView2"
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00DBE6E6&
      Caption         =   "전체"
      Height          =   225
      Left            =   1320
      TabIndex        =   47
      Top             =   1905
      Width           =   675
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00EFE9E4&
      Caption         =   "삭제(&D)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8775
      Style           =   1  '그래픽
      TabIndex        =   46
      Top             =   120
      Width           =   1230
   End
   Begin VB.CommandButton cmdSupp 
      BackColor       =   &H00F5FFF4&
      Caption         =   "Supplemental"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7305
      Style           =   1  '그래픽
      TabIndex        =   43
      Top             =   5445
      Width           =   1335
   End
   Begin VB.CheckBox chkFinal 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종결 보류"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   7095
      TabIndex        =   42
      Top             =   255
      Width           =   1080
   End
   Begin VB.CommandButton cmdEtcResult 
      BackColor       =   &H00E0E0E0&
      Caption         =   "기타결과"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7545
      Style           =   1  '그래픽
      TabIndex        =   40
      Tag             =   "0"
      Top             =   1830
      Width           =   1095
   End
   Begin VB.CommandButton cmdCommentTemplete 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      Picture         =   "frmReport.frx":038A
      Style           =   1  '그래픽
      TabIndex        =   29
      Top             =   5460
      Width           =   315
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00F5FFF4&
      Caption         =   "미리보기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8655
      Style           =   1  '그래픽
      TabIndex        =   27
      Top             =   5445
      Width           =   1335
   End
   Begin VB.CommandButton cmdAllResult 
      BackColor       =   &H00F5FFF4&
      Caption         =   "전체결과보기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8655
      Style           =   1  '그래픽
      TabIndex        =   26
      Top             =   1815
      Width           =   1335
   End
   Begin VB.Frame fraMethod 
      BackColor       =   &H00DBE6E6&
      Caption         =   "◈ 검증방법"
      Height          =   870
      Left            =   135
      TabIndex        =   20
      Top             =   4530
      Width           =   9885
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Calibration Verification"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   375
         TabIndex        =   1
         Top             =   225
         Width           =   2145
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Internal Quality Control"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   3420
         TabIndex        =   2
         Top             =   225
         Width           =   2220
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Delta Check Verification"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   5925
         TabIndex        =   25
         Top             =   225
         Width           =   2685
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Panic/Alert Value Verification"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   375
         TabIndex        =   24
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Repeat / Recheck"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   3420
         TabIndex        =   23
         Top             =   480
         Width           =   1830
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Others;"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   5925
         TabIndex        =   22
         Top             =   480
         Width           =   900
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6810
         TabIndex        =   21
         Top             =   495
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8760
      Style           =   1  '그래픽
      TabIndex        =   16
      Top             =   555
      Width           =   1230
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8760
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   1350
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      Style           =   1  '그래픽
      TabIndex        =   14
      Top             =   945
      Width           =   1230
   End
   Begin VB.ListBox lstTestList 
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   105
      Style           =   1  '확인란
      TabIndex        =   0
      Top             =   2175
      Width           =   1920
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   4770
      TabIndex        =   5
      Top             =   480
      Width           =   3885
      Begin MedControls1.LisLabel lblWardId 
         Height          =   210
         Left            =   1365
         TabIndex        =   34
         Top             =   195
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   370
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   210
         Left            =   1365
         TabIndex        =   35
         Top             =   435
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   370
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblReqDt 
         Height          =   225
         Left            =   1365
         TabIndex        =   36
         Top             =   675
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   397
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblRptDt 
         Height          =   210
         Left            =   1365
         TabIndex        =   37
         Top             =   930
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   370
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보고일자 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   315
         TabIndex        =   28
         Top             =   945
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "입 원 일 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   315
         TabIndex        =   12
         Top             =   705
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "의 뢰 과 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   315
         TabIndex        =   11
         Top             =   465
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "진료병동 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   315
         TabIndex        =   10
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   150
      TabIndex        =   4
      Top             =   480
      Width           =   4545
      Begin MedControls1.LisLabel lblPtId 
         Height          =   210
         Left            =   1545
         TabIndex        =   31
         Top             =   180
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   370
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   210
         Left            =   1545
         TabIndex        =   32
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   370
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblPtSexAge 
         Height          =   225
         Left            =   1545
         TabIndex        =   33
         Top             =   660
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   397
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label lblDiagNm 
         BackColor       =   &H00D1D8D3&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1665
         TabIndex        =   44
         Top             =   945
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D1D8D3&
         Height          =   225
         Left            =   1545
         TabIndex        =   45
         Top             =   915
         Width           =   2910
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "상     병 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   300
         TabIndex        =   9
         Top             =   945
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "나이/성별 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   300
         TabIndex        =   8
         Top             =   705
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자 성명 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   300
         TabIndex        =   7
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "병록 번호 : "
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   300
         TabIndex        =   6
         Top             =   210
         Width           =   1080
      End
   End
   Begin TabDlg.SSTab tabComments 
      Height          =   3555
      Left            =   90
      TabIndex        =   30
      Top             =   5475
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   6271
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "◈ 검증/판독 소견 (Comments)"
      TabPicture(0)   =   "frmReport.frx":08BC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtCmt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "◈ 추천 (Recommendation)"
      TabPicture(1)   =   "frmReport.frx":08D8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRcmd"
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox txtCmt 
         Height          =   3225
         Left            =   30
         TabIndex        =   38
         Top             =   315
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5689
         _Version        =   393217
         BackColor       =   16776191
         ScrollBars      =   3
         TextRTF         =   $"frmReport.frx":08F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtRcmd 
         Height          =   3225
         Left            =   -74970
         TabIndex        =   39
         Top             =   315
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5689
         _Version        =   393217
         BackColor       =   16776183
         ScrollBars      =   3
         TextRTF         =   $"frmReport.frx":0999
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Crystal.CrystalReport crReport 
      Left            =   120
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "종합검증/판독 보고서"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   2325
      Left            =   2055
      TabIndex        =   19
      Top             =   2175
      Width           =   7935
      _Version        =   196608
      _ExtentX        =   13996
      _ExtentY        =   4101
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowVert    =   0   'False
      MaxCols         =   10
      MaxRows         =   20
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmReport.frx":0A3E
      TextTip         =   4
   End
   Begin VB.TextBox txtEtcResult 
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   2055
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   41
      Top             =   2175
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "임상병리과 검사 종합검증/판독 보고서"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E48372&
      Height          =   180
      Left            =   2985
      TabIndex        =   3
      Top             =   195
      Width           =   3495
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ 비정상 결과 혹은 유의한 결과를 보이는 항목"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2115
      TabIndex        =   18
      Top             =   1935
      Width           =   4050
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "☞ 선택하면 보고서          출력시 제외됩니다."
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   405
      Left            =   150
      TabIndex        =   17
      Top             =   3975
      Width           =   1905
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ 검증항목"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1935
      Width           =   990
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   90
      X2              =   9970
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   135
      X2              =   10000
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DBF2FD&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   390
      Left            =   2460
      Shape           =   4  '둥근 사각형
      Top             =   90
      Width           =   4515
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MySql As New clsLISSqlStatement   'Sql문 클래스
Private objCmtSql As New clsLISSqlReview     'Sql문 클래스
Private MyPatient As New clsLisPatient

'-- OCS Table ===================================================
Private Const F_OCSRESULT = "oram1.mdresult"          '결과관리테이블
Private Const T_SLXVERIT = "oras1.slxverit"          '결과관리테이블

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
'================================================================

Private m_DoneFg As Boolean
Private m_PtId As String
Private m_BedinDt As String
Private m_QueryFg As Boolean
Private m_SaveFg As Boolean

Private blnExpect       As Boolean


Public Property Get DoneFg() As Boolean
    DoneFg = m_DoneFg
End Property
Public Property Let DoneFg(ByVal vData As Boolean)
    m_DoneFg = vData
End Property

Public Property Get ptid() As String
    ptid = m_PtId
End Property
Public Property Let ptid(ByVal vData As String)
    m_PtId = vData
End Property

Public Property Get BedinDt() As String
    BedinDt = m_BedinDt
End Property
Public Property Let BedinDt(ByVal vData As String)
    m_BedinDt = vData
End Property

Public Property Get QueryFg() As Boolean
    QueryFg = m_QueryFg
End Property
Public Property Let QueryFg(ByVal vData As Boolean)
    m_QueryFg = vData
End Property

Private Sub chkAll_Click()
    Dim i       As Integer
    
    If chkAll.Value = 1 Then
        For i = 0 To lstTestList.ListCount - 1
            lstTestList.Selected(i) = True
        Next
    Else
        For i = 0 To lstTestList.ListCount - 1
            lstTestList.Selected(i) = False
        Next
    End If
End Sub

Private Sub chkMethod_Click(Index As Integer)
    m_SaveFg = False
    If Index = 5 Then
        If chkMethod(Index).Value = 1 Then
            txtOthers.Enabled = True
        Else
            txtOthers.Text = ""
            txtOthers.Enabled = False
        End If
    End If
End Sub

Private Sub cmdAllResult_Click()

    With frmResultReview
        .Show
        .ZOrder 0
        If .txtPtId.Text <> m_PtId Or Not .QueryFg Then
            .txtPtId.Text = m_PtId
            .BedinDt = m_BedinDt
            Call .Call_PtId_LostFocus
        End If
    End With

End Sub

Private Sub cmdCommentTemplete_Click()
    With frmTextRst
        .Left = Me.Left
        .Top = Me.Height - .Height
        .Show
        .ZOrder 0
        .DoctId = objDoctor.DoctId
        .Txtdiv = Choose(tabComments.Tab + 1, "C", "R")
        .lblTitle.Caption = tabComments.TabCaption(tabComments.Tab)
        If tabComments.Tab = 0 Then
            Set .MyCtrl = Me.txtCmt
            .txtTemplate.Text = Me.txtCmt.Text
        Else
            Set .MyCtrl = Me.txtRcmd
            .txtTemplate.Text = Me.txtRcmd.Text
        End If
        Call .LoadTemplate
    End With
End Sub

Private Sub cmdDelete_Click()
    Dim Resp As VbMsgBoxResult
    Dim SqlStmt As String
    Dim strMedDate As String
    
    Resp = MsgBox("해당 환자의 작성된 모든 보고서가 삭제됩니다. 계속 진행하시겠습니까?", vbQuestion + vbYesNo, "보고서삭제")
    If Resp = vbNo Then Exit Sub
    
On Error GoTo Err_Trap

    DBConn.BeginTrans
    
    SqlStmt = " delete from " & T_LAB502 & _
              " where " & DBW("ptid = ", lblPtId.Caption) & _
              " and   " & DBW("rptdt = ", Format(lblRptDt.Caption, CS_DateDbFormat))
    DBConn.Execute SqlStmt
    SqlStmt = " delete from " & T_LAB503 & _
              " where " & DBW("ptid = ", lblPtId.Caption) & _
              " and   " & DBW("rptdt = ", Format(lblRptDt.Caption, CS_DateDbFormat))
    DBConn.Execute SqlStmt
    SqlStmt = " delete from " & T_LAB504 & _
              " where " & DBW("ptid = ", lblPtId.Caption) & _
              " and   " & DBW("rptdt = ", Format(lblRptDt.Caption, CS_DateDbFormat))
    DBConn.Execute SqlStmt
    SqlStmt = " update " & T_LAB501 & _
              " set " & DBW("rptdt", "", 3) & DBW("rpttm", "", 3) & DBW("rptid", "", 3) & _
                        DBW("prtfg", "", 3) & DBW("donefg", "", 2) & _
              " where " & DBW("ptid = ", lblPtId.Caption) & _
              " and   " & DBW("bedindt = ", Format(lblReqDt.Caption, CS_DateDbFormat))
    DBConn.Execute SqlStmt
    
    '종결 한것도 삭제가 된다면 ocs업데이트 처리를 해 줘야 한다.
    '** 예수병원
    ' - 삭제처리
    '   1. mdexmort (수가처리) : Key(patno, meddate, ordcd = 'B0001')
    '   2. slxverit (결과처리) : Kdy(patno, meddate)
    
    strMedDate = "TO_DATE(" & DBS(Format(lblReqDt.Caption, CS_DateDbFormat)) & ", 'yyyymmdd')"
    
    SqlStmt = " delete from mdexmort " & _
              " where patno = " & DBS(lblPtId.Caption) & _
              " and   meddate = " & strMedDate & _
              " and   ordcd = " & DBS("B0001")
    DBConn.Execute SqlStmt
    
    SqlStmt = " delete from " & T_SLXVERIT & _
              " where patno = " & DBS(lblPtId.Caption) & _
              " and   meddate = " & strMedDate
    DBConn.Execute SqlStmt
    
    DBConn.CommitTrans
    
    MsgBox "정상적으로 삭제되었습니다.", vbInformation, "메세지"
    
    Call frmPtList.Call_Refresh

    Call cmdExit_Click

    Exit Sub
    
Err_Trap:
    DBConn.RollbackTrans
    
End Sub

Private Sub cmdEtcResult_Click()
    
    If cmdEtcResult.Tag = "0" Then  '기타결과
        cmdEtcResult.Tag = "1"
        cmdEtcResult.Caption = "일반결과"
        tblResult.Visible = False
        txtEtcResult.Visible = True
        txtEtcResult.ZOrder 0
        txtEtcResult.SetFocus
    Else        '일반결과
        cmdEtcResult.Tag = "0"
        cmdEtcResult.Caption = "기타결과"
        txtEtcResult.Visible = False
        tblResult.Visible = True
        tblResult.ZOrder 0
        tblResult.SetFocus
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmReport = Nothing
End Sub

Private Sub cmdPreview_Click()

    Call PrtReport(1)

End Sub

Private Sub cmdPrint_Click()
    
    Dim SqlStmt As String
    
    Call PrtReport(0)
    
    SqlStmt = " update " & T_LAB501 & " set prtfg = 'Y' " & _
              " where  " & DBW("ptid = ", gPtntId) & _
              " and    " & DBW("bedindt = ", gBedInDT)
              
    On Error GoTo Err_Trap
    
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
    
    Call frmPtList.Call_Refresh
    
    Exit Sub
    
Err_Trap:
'    Call Error_Routine
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdSave_Click()
    
    Dim SqlStmt() As String
    Dim strItems As String
    Dim strMethod As String
    Dim strDoneFg As String
    Dim strTmp As String
    Dim iSeq As Integer
    Dim i As Integer
    Dim strKey As String
    Dim blnOCS As Boolean
    
    Dim strDeptCd   As String
    Dim strMajDoct  As String
    Dim Rs As Recordset
    
    '-- 예수병원 추가변수 =================
    Dim rsTmp       As New ADODB.Recordset
    Dim strSql      As String
    Dim strAllItem  As String
    Dim strHighItem As String
    Dim strLowItem  As String
    Dim strMedDate  As String
    Dim strMedDept  As String
    Dim strChaDr    As String
    Dim strOrdDr    As String
    Dim strWardNo   As String
    Dim strRoomNo   As String
    Dim strTitle1   As String
    Dim strTitle2   As String
    Dim strTestNm   As String
    Dim strResults  As String
    Dim strUnit     As String
    Dim strHLDP     As String
    Dim strSpcNm    As String
    Dim strVfyDt    As String
    Dim strMain     As String
    Dim strAbnormal As String
    '======================================
    
    blnOCS = False
    If lblPtId.Caption = "" Then
        MsgBox "저장할 대상이 없습니다.", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    
    strItems = ""
    
    For i = 0 To lstTestList.ListCount - 1
        If Not lstTestList.Selected(i) Then
            strItems = strItems & lstTestList.List(i) & ","
        End If
    Next
    
    If Len(strItems) > 500 Then
        MsgBox "검증항목이 너무 많습니다. (" & CStr(Len(strItems) - 500) & " 문자 초과)", vbExclamation, "메세지"
        Exit Sub
    End If
    
    strMethod = ""
    For i = 0 To chkMethod.Count - 1
        strMethod = strMethod & chkMethod(i).Value
    Next
    If chkFinal.Value = 1 Then '종결보류
        strDoneFg = "1"
    Else
        strDoneFg = "2"
    End If
    
    ReDim SqlStmt(3)
    
    'lab501 save
    SqlStmt(1) = " update " & T_LAB501 & " set " & DBW("donefg = ", strDoneFg) & ", " & _
                 DBW("rptdt = ", Format(Now, CS_DateDbFormat)) & ", " & _
                 DBW("rpttm = ", Format(Now, CS_TimeDbFormat)) & ", " & _
                 DBW("rptid = ", objDoctor.DoctId) & _
                 " where " & DBW("ptid    = ", lblPtId.Caption) & _
                 " and   " & DBW("bedindt = ", m_BedinDt)
                 
    If m_DoneFg Then    '중간보고
        SqlStmt(2) = " update " & T_LAB502 & " set age = " & MyPatient.Age & ", " & _
                     " agediv = '" & MyPatient.AgeDiv & "', sex = '" & MyPatient.Sex & "', " & _
                     " items = '" & strItems & "', method = '" & strMethod & "', " & _
                     " others = '" & Trim(txtOthers.Text) & "', cmttxt = '" & Trim(txtCmt.Text) & "', " & _
                     " recmd = '" & Trim(txtRcmd.Text) & "', etcrst = '" & Trim(txtEtcResult.Text) & "' " & _
                     " where " & DBW("ptid  = ", lblPtId.Caption) & _
                     " and   " & DBW("rptdt = ", Format(lblRptDt.Caption, CS_DateDbFormat))
                     
    Else    '신규
        '** 원본 ================================================================================
'        SqlStmt(2) = " insert into " & T_LAB502 & " (ptid, rptdt, age, agediv, sex, " & _
'                     " items, method, others, cmttxt, recmd, etcrst ) values (" & DBV("ptid", lblPtId.Caption) & "," & _
'                     " '" & Format(lblRptDt.Caption, CS_DateDbFormat) & "', " & MyPatient.Age & ", " & _
'                     " '" & MyPatient.AgeDiv & "', '" & MyPatient.Sex & "', " & _
'                     " '" & strItems & "', '" & strMethod & "', " & _
'                     " '" & Trim(txtOthers.Text) & "', '" & Trim(txtCmt.Text) & "', " & _
'                     " '" & Trim(txtRcmd.Text) & "', '" & Trim(txtEtcResult.Text) & "') "
        '========================================================================================
        
        '** 예수병원 ============================================================================
        SqlStmt(2) = " insert into " & T_LAB502 & " (ptid, rptdt, age, agediv, sex, " & _
                     " items, method, others, cmttxt, recmd, etcrst ) values (" & DBV("ptid", lblPtId.Caption) & "," & _
                     " '" & Format(lblRptDt.Caption, CS_DateDbFormat) & "', " & MyPatient.Age & ", " & _
                     " '" & MyPatient.AgeDiv & "', '" & MyPatient.Sex & "', " & _
                     DBS(strItems) & "," & DBS(strMethod) & "," & _
                     DBS(txtOthers.Text) & "," & DBS(txtCmt.Text) & "," & _
                     DBS(txtRcmd.Text) & "," & DBS(txtEtcResult.Text) & ") "
        '========================================================================================
        
        m_DoneFg = True
        blnOCS = True
    End If
    SqlStmt(3) = "delete from " & T_LAB503 & _
                 " where " & DBW("ptid  = ", lblPtId.Caption) & _
                 " and   " & DBW("rptdt = ", Format(lblRptDt.Caption, CS_DateDbFormat))
    
    With tblResult
        iSeq = 0: strMain = ""
        
        strTitle1 = "◈ 비정상 결과 혹은 유의한 결과를 보이는 항목"
        strTitle2 = "검사명" & Space(10) & "결과" & Space(7) & "단위" & Space(3) & _
                   "판정" & Space(5) & "검체" & Space(6) & "보고일시"
                   
        For i = 1 To .MaxRows
            .Row = i: .Col = 1
            If .Value = 1 Then GoTo Skip    '제외
            iSeq = iSeq + 1
            strTmp = "insert into " & T_LAB503 & " (ptid, rptdt, seq, testnm, rstcd, rstunit, " & _
                     "hldiv, dpdiv, spcnm, vfydttm) values (" & DBV("ptid", lblPtId.Caption) & "," & _
                         " '" & Format(lblRptDt.Caption, CS_DateDbFormat) & "', " & iSeq & ", "
            .Col = 10: strTmp = strTmp & "'" & .Value & "', "    '검사명
            .Col = 3: strTmp = strTmp & "'" & .Value & "', "    '결과
            .Col = 4: strTmp = strTmp & "'" & .Value & "', "    '단위
            .Col = 8: strTmp = strTmp & "'" & .Value & "', "    'Low/High
            
            '-- 예수병원 추가루틴 ==========================================
'            If i <= 10 Then
                .Col = 2: strTestNm = .Value
                .Col = 3: strResults = .Value
                .Col = 4: strUnit = .Value
                .Col = 5: strHLDP = .Value
                .Col = 6: strSpcNm = .Value
                .Col = 7: strVfyDt = .Value
                
                strMain = strMain & Mid$(strTestNm, 1, 13) & Space(15 - Len(Mid$(strTestNm, 1, 13))) & _
                          Mid$(strResults, 1, 8) & Space(10 - Len(Mid$(strResults, 1, 8))) & _
                          Mid$(strUnit, 1, 5) & Space(7 - Len(Mid$(strUnit, 1, 5))) & _
                          Mid$(strHLDP, 1, 5) & Space(7 - Len(Mid$(strHLDP, 1, 5))) & _
                          Mid$(strSpcNm, 1, 7) & Space(10 - Len(Mid$(strSpcNm, 1, 7))) & _
                          Mid$(strVfyDt, 1, 14) & Space(15 - Len(Mid$(strVfyDt, 1, 14))) & Chr(13)
'            End If
            
            If .Value = "H" Then
                strHighItem = strHighItem & .Value
            ElseIf .Value = "L" Then
                strLowItem = strLowItem & .Value
            End If
            strAllItem = strAllItem & .Value
            '===============================================================
            
            .Col = 9: strTmp = strTmp & "'" & .Value & "', "    'Delta/Panic
            .Col = 6: strTmp = strTmp & "'" & .Value & "', "    '검체명
            .Col = 7: strTmp = strTmp & "'" & .Value & "') "    '보고일시
            
            
            
            ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
            SqlStmt(UBound(SqlStmt)) = strTmp
Skip:
        Next
        
        If strMain = "" Then
            strAbnormal = strTitle1 & Chr(13) & strTitle2
        Else
            strAbnormal = strTitle1 & Chr(13) & strTitle2 & Chr(13) & strMain
        End If
        
    End With
    
    If blnOCS Then
        
'        Set Rs = New Recordset
'        strTmp = " SELECT deptcd,majdoct FROM " & T_LAB501 & _
'                 "  WHERE " & DBW("ptid    = ", lblPtId.Caption) & _
'                 "    AND " & DBW("bedindt = ", m_BedinDt)
''        Rs.RsOpen , strTmp
'        Rs.Open strTmp, DBConn
'        If Rs.RecordCount > 0 Then
'            strMajDoct = Rs.Fields("majdoct").Value & ""
'            strDeptCd = Rs.Fields("deptcd").Value & ""
'        End If
''        Rs.RsClose
'        Set Rs = Nothing
'
'        '오더키
'        strKey = GetOrderKey
'        'ocs 수가내역 생성
'        strTmp = " insert into med_ocs.ipd_order_update_dmc " & _
'                 " (patient_no,order_date,clinical_dept,order_check,class_of_order,group_of_order," & _
'                 "  item_name,item_code,contents_of_order,qty_of_order,value_of_frequency,dosage_of_order, " & _
'                 "  unit_of_order,base_contents,night_day,stat_normal,remark,return_value, " & _
'                 "  duration_of_order,prn,transfer_to_pmpa,acting_check,seq_no,dr_code,input_status,write_status, " & _
'                 "  self,receipt_no,left_or_right,pmpa_code,order_key,confirm_flag,reserve_flag,base_price, " & _
'                 "  write_date,ordered_dr_code,nurse_check,order_site,flag_update_cancel,flag_apply_night_op, " & _
'                 "  pickup_time,cancel_of_frequency,pre_cost,pickup_nurse,order_ver,pc_name)  values(" & _
'                    DBS(lblPtId.Caption, 1) & ConvToDate(Format(GetSystemDate, CS_DateDbFormat)) & "," & DBS(strDeptCd, 1) & DBS("0", 1) & DBN("2", 1) & DBN("9", 1) & _
'                    DBS("임상병리검사 종합검진료", 1) & DBS("LHE114", 1) & DBN("1", 1) & DBN("1", 1) & DBN("1", 1) & DBS(Format(GetSystemDate, CS_DateDbFormat), 1) & _
'                    DBS("0", 1) & DBN("0", 1) & DBN("0", 1) & DBS("0", 1) & DBS("임상병리 종합검증", 1) & DBS("1", 1) & _
'                    DBN("1", 1) & DBS("0", 1) & DBS("0", 1) & DBS("0", 1) & DBN("0", 1) & DBS(strMajDoct, 1) & DBS("1", 1) & DBS("1", 1) & _
'                    DBS("0", 1) & DBS("0", 1) & DBS("0", 1) & DBS("B0001", 1) & DBN(strKey, 1) & DBS("1", 1) & DBS("0", 1) & DBN("0", 1) & _
'                    "SYSDATE" & "," & DBS(objDoctor.DoctId, 1) & DBS("1", 1) & DBS("CP", 1) & DBS("1", 1) & DBS("0", 1) & _
'                    "SYSDATE" & "," & DBN("0", 1) & DBS("0", 1) & DBS("LIS", 1) & DBS("LIS.1.1.1", 1) & DBS(medGetP(medGetComNm, 1, Chr(0))) & ")"
'
'
'
'
'        ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
'        SqlStmt(UBound(SqlStmt)) = strTmp
        
    '테스트 시 transfer_to_pmpa ='1' pickup_time is null
    End If
        
    '** 전주 예수 병원 종합검증료 수가 내역 생성 ============================================
    '-- OCS 처방순번 가져오기 (MDEXMORT) => MAX(OrdSeqNo) + 1
    '-- 환자정보 ==============================================
    strSql = " select * from " & T_LAB501 & _
             "  where " & DBW("ptid    = ", lblPtId.Caption) & _
             "    and " & DBW("bedindt = ", m_BedinDt)
    
    rsTmp.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If rsTmp.EOF = False Then
        strMedDept = rsTmp.Fields("deptcd").Value & ""
        strChaDr = rsTmp.Fields("majdoct").Value & ""
        strOrdDr = objDoctor.DoctId
        strWardNo = rsTmp.Fields("wardid").Value & ""
        strRoomNo = rsTmp.Fields("hosilid").Value & ""
    End If
    
    rsTmp.Close: Set rsTmp = Nothing
    '==========================================================
    
    If blnOCS Then
        Dim rsOCS       As New ADODB.Recordset
        Dim strOrdDt    As String
        Dim strOCSSeq   As String
        Dim strOrdDesc1 As String
        Dim strOrdTime  As String
        Dim strRcpAplDt As String
        Dim strResult   As String
        
        strOrdDt = "TO_DATE(" & DBS(Format(Now, "yyyymmdd")) & ", 'yyyymmdd')"
        
        '-- 순번 ==================================================
        strSql = " select max(ordseqno) as seq from mdexmort " & _
                 "  where orddate  = " & strOrdDt & _
                 "    and patno    = " & DBS(lblPtId.Caption)
                 
        rsOCS.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
        
        If rsOCS.EOF = False Then
            If rsOCS.Fields("seq").Value <> "" Then
                strOCSSeq = CInt(rsOCS.Fields("seq").Value) + 1
            Else
                strOCSSeq = "2001"
            End If
        Else
            strOCSSeq = "2001"
        End If
        
        rsOCS.Close: Set rsOCS = Nothing
        '==========================================================
        
        '-- 입원일자/진료일자
        If m_BedinDt <> "" Then
            strMedDate = "TO_DATE(" & DBS(m_BedinDt) & ", 'yyyymmdd')"
        Else
            strMedDate = "''"
        End If
        
        '-- 처방편집1
        strOrdDesc1 = "임상병리검사 종합검증료"
        
        '-- 처방시간
        strOrdTime = "TO_DATE(" & DBS(Format(Now, "yyyymmddhhmmss")) & ", 'yyyymmdd hh24:mi:ss')"
        
        '-- 원무수납적용일자(현재일자로 한다.)
        strRcpAplDt = "TO_DATE(" & DBS(Format(Now, "yyyymmdd")) & ", 'yyyymmdd')"
        
        strTmp = " insert into mdexmort " & _
                 " (patno, orddate, ordseqno, meddate, patsect, patsite, ordgrp, slipcd, " & _
                 "  ordtype, ordkind, meddept, chadr, orddr, ordcd, acptdate, execdate, " & _
                 "  readdate, rsltdate, cofmdr, wardno, roomno, ordsite, orddesc1, " & _
                 "  ordtime, rcpapldt, inputid) " & _
                 " values (" & DBS(lblPtId.Caption) & "," & strOrdDt & "," & DBN(strOCSSeq) & _
                 "," & strMedDate & "," & DBS("I") & "," & DBS("I") & "," & DBS("C1") & "," & DBS("L54") & _
                 "," & DBS("A") & "," & DBS(1) & "," & DBS(strMedDept) & "," & DBS(strChaDr) & "," & DBS(strOrdDr) & _
                 "," & DBS("B0001") & "," & strOrdTime & "," & strOrdTime & "," & strOrdTime & _
                 "," & strOrdTime & "," & DBS(objDoctor.DoctId) & "," & DBS(strWardNo) & _
                 "," & DBS(strRoomNo) & "," & DBS("06") & "," & DBS(strOrdDesc1) & "," & strOrdTime & _
                 "," & strRcpAplDt & "," & DBS(objDoctor.DoctId) & ")"
                 
        ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
        SqlStmt(UBound(SqlStmt)) = strTmp
        
    End If
    
    '** 결과전송
    strTmp = OCS_Transfer_Result(lblPtId.Caption, m_BedinDt, strMedDept, strWardNo, _
                                strRoomNo, strAllItem, strHighItem, strLowItem, _
                                Format(Now, "yyyymmddhhmmss"), objDoctor.DoctId, _
                                Format(Now, "yyyymmddhhmmss"), strAbnormal)
    
    If strTmp = "" Then
        GoTo Err_Trap
    End If
    
    ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
    SqlStmt(UBound(SqlStmt)) = strTmp
    '========================================================================================
    
    On Error GoTo Err_Trap
    
    Dim strTmpSQL   As String
    
    DBConn.BeginTrans
    For i = 1 To UBound(SqlStmt)
        'Debug.Print SqlStmt(i)
        
        If strTmpSQL = SqlStmt(i) Then
            GoTo SkipTmp
        End If
        
        DBConn.Execute SqlStmt(i)
        
        strTmpSQL = SqlStmt(i)
        
SkipTmp:
        
    Next
    
    '** 예수병원 추가 루틴 ============================================================
    If blnOCS Then
        '-- 종합 검증 판독/소견
        
'        Call DBConn.Execute(SqlOCSRESULT_INSERT(lblPtId.Caption, Format(Now, "yyyymmdd"), strOCSSeq, _
'                                "B0001", "L54", "", strResult, "", "", "", "", _
'                                Format(Now, "yyyymmddhhmmss"), Format(Now, "yyyymmddhhmmss"), _
'                                objDoctor.DoctId, Format(Now, "yyyymmddhhmmss"), objDoctor.DoctId, _
'                                Format(Now, "yyyymmddhhmmss"), "T", "", "", ""))
    End If
    '==================================================================================
    DBConn.CommitTrans
    
    MsgBox "정상적으로 저장되었습니다.", vbInformation, "메세지"
    If chkFinal.Value = 0 Then
        cmdPrint.Enabled = True
    End If
    
    Call LockRtn(strDoneFg)
    
    Call frmPtList.Call_Refresh
    cmdPreview.Enabled = True
    m_SaveFg = True
    
    Exit Sub
    
Err_Trap:
'    Call Error_Routine
    DBConn.RollbackTrans
    m_SaveFg = False
    MsgBox Err.Description, vbExclamation
End Sub

Private Function ConvToDate(ByVal argDate As String) As String
    ConvToDate = "To_Date('" & argDate & "', 'YYYYMMDD') "
End Function

Private Function GetOrderKey() As String
    Dim Rs      As Recordset
    Dim strSql  As String
    
    GetOrderKey = ""
    strSql = " SELECT med_ocs.SEQ_ORDER_KEY.NEXTVAL MAXKEY FROM DUAL "
'    Set Rs = OpenRecordSet(strSql)
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Not Rs.EOF Then GetOrderKey = Rs.Fields("MAXKEY").Value & ""
    
    Set Rs = Nothing
End Function

Private Sub cmdSupp_Click()
    With frmSupplemental
        .Top = Me.Top + (Me.Height - .Height) / 2
        .Left = Me.Left + (Me.Width - .Width) / 2
        .ptid = m_PtId
        .RptDt = Format(lblRptDt.Caption, CS_DateDbFormat)
        .BedinDt = m_BedinDt
        .Show
        .ZOrder 0
        Call .GetSuppText
    End With
End Sub

Private Sub Form_Activate()
    Me.Left = 4845
    Me.Top = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Me.Left = 4845
    Me.Top = 0
    chkFinal.Value = 1
End Sub

Public Sub StartQuery()
    
    Call GetPtInfo
    If m_DoneFg Then
        Call GetReport
        cmdPreview.Enabled = True
    Else
        lblRptDt.Caption = Format(Now, CS_DateLongFormat)
        Call GetItems
        Call GetResult
        Call LockRtn("0")
        m_SaveFg = False
    End If
    
    m_QueryFg = True
    
End Sub

Public Sub StartQuery_New(ByVal strTmp As String)
    
    Call GetPtInfo_New(strTmp)
    If m_DoneFg Then
        Call GetReport
        cmdPreview.Enabled = True
    Else
        lblRptDt.Caption = Format(Now, CS_DateLongFormat)
        Call GetItems
        Call GetResult
        Call LockRtn("0")
        m_SaveFg = False
    End If
    
    m_QueryFg = True
    
End Sub

Private Sub GetPtInfo()
    
    If m_PtId = "" Then Exit Sub
    
    With MyPatient
        If .PtntQuery(m_PtId, m_BedinDt) Then
            lblPtId.Caption = m_PtId
            lblPtNm.Caption = .PtNm
            lblPtSexAge.Caption = .Sex & " / " & .Age & "  " & .AgeDiv
            lblDiagNm.Caption = .InDiseaCd
            lblDiagNm.ToolTipText = .InDiseaCd
            lblWardId.Caption = .WardId
            lblDeptNm.Caption = .DeptNm
            lblReqDt.Caption = Format(.BedinDt, CS_DateMask)
            If Val(.DoneFg) > 0 Then
                m_DoneFg = True
                If .DoneFg = "2" Then   '보고종결
                    cmdSave.Enabled = False
                Else
                    cmdSave.Enabled = True
                End If
            Else
                m_DoneFg = False
            End If
        End If
    End With

End Sub

Private Sub GetPtInfo_New(ByVal strTmp As String)
    
    If m_PtId = "" Then Exit Sub
       
    With MyPatient
        If .PtntQuery(m_PtId, strTmp) Then
            lblPtId.Caption = m_PtId
            lblPtNm.Caption = .PtNm
            lblPtSexAge.Caption = .Sex & " / " & .Age & "  " & .AgeDiv
            lblDiagNm.Caption = .InDiseaCd
            lblDiagNm.ToolTipText = .InDiseaCd
            lblWardId.Caption = .WardId
            lblDeptNm.Caption = .DeptNm
            lblReqDt.Caption = Format(.BedinDt, CS_DateMask)
            If Val(.DoneFg) > 0 Then
                m_DoneFg = True
                If .DoneFg = "2" Then   '보고종결
                    cmdSave.Enabled = False
                Else
                    cmdSave.Enabled = True
                End If
            Else
                m_DoneFg = False
            End If
        End If
    End With

End Sub


Private Sub GetReport()
    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim strItems As String
    Dim strRptDt As String
    Dim strMajDoct As String
    Dim strDPDiv    As String
    Dim i As Integer
    
    SqlStmt = " select a.donefg,a.majdoct, b.* from " & T_LAB501 & " a, " & T_LAB502 & " b " & _
              " where " & DBW("a.ptid = ", m_PtId) & _
              " and   a.bedindt = '" & m_BedinDt & "' " & _
              " and   b.ptid = a.ptid " & _
              " and   b.rptdt = a.rptdt "
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        m_DoneFg = False
        lblRptDt.Caption = Format(Now, CS_DateLongFormat)
        Call GetItems
        Call GetResult
        Call LockRtn("0")
        m_SaveFg = False
        Exit Sub
    Else
        strMajDoct = "" & Rs.Fields("majdoct").Value
        strRptDt = "" & Rs.Fields("RptDt").Value
        lblRptDt.Caption = Format(strRptDt, CS_DateMask)
        strItems = "" & Rs.Fields("Items").Value
        txtEtcResult.Text = "" & Rs.Fields("EtcRst").Value
        For i = 1 To chkMethod.Count
            chkMethod(i - 1).Value = Val(Mid("" & Rs.Fields("Method").Value, i, 1))
        Next
        txtOthers.Text = "" & Rs.Fields("Others").Value
        txtCmt.Text = "" & Rs.Fields("cmttxt").Value
        txtRcmd.Text = "" & Rs.Fields("Recmd").Value
        lstTestList.Clear
        While (strItems <> "")
            lstTestList.AddItem medShift(strItems, ",")
        Wend
        Call LockRtn("" & Rs.Fields("DoneFg").Value)
    End If
    
'    Rs.RsClose
    
    SqlStmt = "select * from " & T_LAB503 & " where ptid = " & m_PtId & " and rptdt = '" & strRptDt & "'"
'    Set Rs = OpenRecordSet(SqlStmt)
    Set Rs = Nothing
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    With tblResult
        
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = Rs.RecordCount
        .Row = 0
        
        While (Not Rs.EOF)
            .Row = .Row + 1
            .Col = 1: .Value = 0
            .Col = 2: .Value = Trim("" & Rs.Fields("TestNm").Value)
            .Col = 3: .ForeColor = DCM_Brown           '-- 결과명(코드일 경우..)
                      .Value = Trim("" & Rs.Fields("RstCd").Value)
            .Col = 4: .Value = Trim("" & Rs.Fields("RstUnit").Value)         '-- 결과단위
            .Col = 5       '-- High / Low
                      .Value = ""
                      If Trim("" & Rs.Fields("HLDiv").Value) = "H" Then .Value = "▲": .ForeColor = DCM_LightRed
                      If Trim("" & Rs.Fields("HLDiv").Value) = "L" Then .Value = "▼": .ForeColor = DCM_LightBlue
                      If Trim("" & Rs.Fields("HLDiv").Value) = "*" Then .Value = "*": .ForeColor = vbRed
                    '## 1.1.44: 이상대(2005-05-23)
                    '   - Alpha결과 참고치를 "N"에서 "Abnormal"표시 변경
                    strDPDiv = Trim("" & Rs.Fields("DPDiv").Value)
                    strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
                    .Value = .Value & " " & strDPDiv
            .Col = 6: .Value = Trim("" & Rs.Fields("SpcNm").Value)    '검체명
            .Col = 7: .Value = Trim("" & Rs.Fields("VfyDtTm").Value)         '-- 보고일시
            .Col = 8: .Value = Trim("" & Rs.Fields("HLDiv").Value)
            .Col = 9: .Value = Trim("" & Rs.Fields("DPDiv").Value)
            .Col = 10: .Value = Trim("" & Rs.Fields("TestNm").Value)
            
            Rs.MoveNext
        
        Wend
        
        .RowHeight(-1) = 11
        .ReDraw = True
        
    End With
    
    m_SaveFg = True
'    Rs.RsClose
    Set Rs = Nothing
    
End Sub

Private Sub LockRtn(ByVal pDoneFg As String)

    Select Case pDoneFg
    Case "0", "1"   '보류/신규
        cmdSave.Enabled = True
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        chkFinal.Enabled = True
        If pDoneFg = "1" Then
            chkFinal.Value = 1
        Else
            chkFinal.Value = 0
        End If
        lstTestList.Enabled = True
        tblResult.Enabled = True
        fraMethod.Enabled = True
        txtCmt.Locked = False
        txtRcmd.Locked = False
        cmdCommentTemplete.Enabled = True
        cmdSupp.Enabled = False
        If pDoneFg = "0" Then
            m_SaveFg = False
        Else
            m_SaveFg = True
        End If
    Case "2"    '최종
        cmdSave.Enabled = True
        cmdPrint.Enabled = True
        cmdPreview.Enabled = True
        chkFinal.Value = 0
        chkFinal.Enabled = False
        lstTestList.Enabled = True  'False
        tblResult.Enabled = True    'False
        fraMethod.Enabled = True    'False
        txtCmt.Locked = False       'True
        txtRcmd.Locked = False      'True
        cmdCommentTemplete.Enabled = True   'False
        cmdSupp.Enabled = True
        m_SaveFg = True
    End Select

End Sub

Private Sub GetItems()

    Dim SqlStmt As String
    Dim tmpRs As Recordset
    
    SqlStmt = " select distinct  b.abbrnm10 as testnm from " & T_LAB001 & " b, " & T_LAB102 & " a " & _
                " where " & DBW("a.ptid = ", m_PtId) & _
                " and   " & DBW("a.orddt >= ", m_BedinDt) & _
                " and   b.testcd = a.ordcd " & _
                " and   a.stscd >= '5' " & _
                " and   b.applydt = (select max(applydt) from " & T_LAB001 & " " & _
                "                    where testcd = b.testcd) " & _
                " and  (b.detailfg = ''  or  b.detailfg is null) " & _
                " order by testnm "
'    Set tmpRs = OpenRecordSet(SqlStmt)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    lstTestList.Clear
    While (Not tmpRs.EOF)
        lstTestList.AddItem Trim("" & tmpRs.Fields("TestNm").Value)
        tmpRs.MoveNext
    Wend
    
'    tmpRs.RsClose
    Set tmpRs = Nothing
    
End Sub

Private Sub GetResult()
   Dim i As Integer, j As Integer
   Dim SqlStmt      As String
   Dim ColCnt       As Integer
   Dim strStartDt   As String
   Dim tmpTestNm    As String
   Dim tmpRs        As New Recordset
   Dim strDPDiv     As String
   Dim strOrdDiv    As String
   
   Me.Enabled = False
   QueryFg = True
   
   Screen.MousePointer = vbArrowHourglass  '13
   With frmMain
        .lblMsg.Caption = ""
        .lblMsg.Top = iMsgTop1
        .pgrBar.Value = 0
        .pgrBar.Visible = True
        .lblMsg.Caption = lblPtNm.Caption & " 님의 검사결과 내역을 조회하고 있습니다."
   End With
   DoEvents
   
   strStartDt = Format(DateAdd("d", Val(objDoctor.Daycnt) * (-1), Now), CS_DateDbFormat)
   SqlStmt = objCmtSql.SqlResultsForCmt(m_PtId, m_BedinDt, Format(Now, CS_DateDbFormat))
'   Set tmpRs = OpenRecordSet(SqlStmt)
   tmpRs.Open SqlStmt, DBConn
   
   frmMain.pgrBar.Max = tmpRs.RecordCount + 1
   
   DoEvents
   
   ReDim aryMesg(0)
'   DisplayOrders = False
   
   With tblResult
   
      .ReDraw = False
      .MaxRows = 0
      
      While (Not tmpRs.EOF)
         
         If frmMain.pgrBar.Value < frmMain.pgrBar.Max Then frmMain.pgrBar.Value = frmMain.pgrBar.Value + 1
         DoEvents
      
         .MaxRows = .MaxRows + 1
         .Row = .MaxRows
         
'         NormalFg = True
         
         .Col = 1:  .Value = 0
         .Col = 2: '.ForeColor = CR_LIGHT_BLUE      '-- 검사명
                    tmpTestNm = Mid(Trim("" & tmpRs.Fields("TestLongNm").Value), 1, 33)
                    'If Trim(tmpRs.Fields("DetailFg").Value) = "" Or _
                       Trim(tmpRs.Fields("RstDiv").Value) = "*" Then
                       .Value = tmpTestNm & " " & String(35 - Len(tmpTestNm), ".")
                    'Else
                    '   .Value = "    " & tmpTestNm & " " & String(35 - Len("  " & tmpTestNm), ".")         '-- 상세검사명
                    'End If
         .Col = 3:
                    strOrdDiv = "" & tmpRs.Fields("orddiv").Value
                    
                    .ForeColor = DCM_Brown           '-- 결과명(코드일 경우..)
                    
                    Select Case strOrdDiv
                        Case "1" '일반검사
                            If Trim("" & tmpRs.Fields("RstCdNm").Value) = "" Then
                               .TypeHAlign = TypeHAlignCenter
                               .Value = Trim("" & tmpRs.Fields("RstCd").Value)
                            Else
                               .CellType = CellTypeEdit
                               .TypeHAlign = TypeHAlignLeft
                               .Value = " " & Trim("" & tmpRs.Fields("RstCdNm").Value)
                            End If
                            If Trim("" & tmpRs.Fields("RstCd").Value) = "" Then
                                .Value = Space(3)
                            End If
                            
                        Case "2" '미생물검사 (결과에 균주명)
                            .CellType = CellTypeEdit
                            .TypeHAlign = TypeHAlignLeft
                            .Value = GetMicroNm("" & tmpRs.Fields("mnmcd").Value)
                            
                    End Select
                    
         .Col = 4:  .Value = Trim("" & tmpRs.Fields("RstUnit").Value)         '-- 결과단위
         .Col = 5       '-- High / Low
                    .Value = ""
                    If Trim("" & tmpRs.Fields("HLDiv").Value) = "H" Then .Value = "▲": .ForeColor = DCM_LightRed
                    If Trim("" & tmpRs.Fields("HLDiv").Value) = "L" Then .Value = "▼": .ForeColor = DCM_LightBlue
                    If Trim("" & tmpRs.Fields("HLDiv").Value) = "*" Then .Value = "*": .ForeColor = vbRed

                    '## 1.1.44: 이상대(2005-05-23)
                    '   - Alpha결과 참고치를 "N"에서 "Abnormal"표시 변경
                    strDPDiv = Trim("" & tmpRs.Fields("DPDiv").Value)
                    strDPDiv = IIf(strDPDiv = "N", "Abnormal", strDPDiv)
                    .Value = .Value & " " & strDPDiv
         .Col = 6:  .Value = Trim("" & tmpRs.Fields("SpcNm").Value)    '검체명
         .Col = 7:  .Value = Trim("" & tmpRs.Fields("VfyDtTm").Value)         '-- 보고일시
         .Col = 8:  .Value = Trim("" & tmpRs.Fields("HLDiv").Value)
         .Col = 9:  .Value = Trim("" & tmpRs.Fields("DPDiv").Value)
         .Col = 10: .Value = Trim("" & tmpRs.Fields("TestLongNm").Value)
         
         tmpRs.MoveNext
         
'         DisplayOrders = True
      
      Wend
      
     .Row = -1: .Col = 2: .Col2 = 3
     .BlockMode = True
     .AllowCellOverflow = True
     .BlockMode = False
       
     .RowHeight(-1) = 11
     .ReDraw = True
'      tmpRs.RsClose
      
     .ReDraw = True
   
   End With
      
    For i = 0 To chkMethod.Count - 1
        chkMethod(i).Value = Val(Mid(objDoctor.Method, i + 1, 1))
    Next
    txtOthers.Text = objDoctor.Others
        
   txtCmt.Text = ""
   txtRcmd.Text = ""
   txtEtcResult.Text = ""
   
   With frmMain
      .pgrBar.Value = frmMain.pgrBar.Max
      .lblMsg.Caption = ""
      .lblMsg.Top = iMsgTop2
      .pgrBar.Visible = False
      Call objDoctor.GetRptCount
      Call .RptStatus(objDoctor.RptCount)
   End With
      
   
NoData:
   QueryFg = False
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   Set tmpRs = Nothing
   
End Sub

Private Function GetMicroNm(ByVal pMnmCd As String) As String
    Dim strSql      As String
    Dim Rs          As New ADODB.Recordset
    
    On Error Resume Next
    
    strSql = " select text1 from " & T_LAB032 & _
             "  where cdindex = " & DBS(LC3_Microbe) & _
             "    and cdval1 = " & DBS(pMnmCd)

    Rs.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        GetMicroNm = Trim(Rs.Fields("text1").Value & "")
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim Resp As VbMsgBoxResult
    
    If Not m_SaveFg Then
        Resp = MsgBox("변경된 데이타를 저장하지 않고 종료하시겠습니까?", vbQuestion + vbYesNo, "메세지")
        If Resp = vbNo Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub lstTestList_Click()
    m_SaveFg = False
End Sub

Private Sub tblResult_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    m_SaveFg = False
End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i       As Integer
    
    If Row = 0 And Col = 1 Then
        With tblResult
            .Col = 1
            blnExpect = IIf(blnExpect, False, True)
            For i = 1 To .MaxRows
                .Row = i
                If .CellType = CellTypeCheckBox Then
                    .Value = IIf(blnExpect, 0, 1)
                End If
            Next
        End With
    End If
End Sub

Private Sub txtCmt_Change()
    
    m_SaveFg = False

End Sub

Private Sub txtEtcResult_Change()
    m_SaveFg = False
    If Trim(txtEtcResult.Text) = "" Then
        cmdEtcResult.BackColor = &HE0E0E0
    Else
        cmdEtcResult.BackColor = &HFFFBFF
    End If
End Sub


Public Sub PrtReport(ByVal pOption As Integer)
    Dim hwndPreviewWindow As Long
    Dim SqlStmt     As String
    
    Me.MousePointer = 11

'전주 예수병원 S2HIS006 : S2HIS006
'전주 예수병원 S2HIS001 : S2HIS001

On Error GoTo PRINT_ERROR
    SqlStmt = "SELECT  S2LAB505.doctid,S2LAB505.doctnm , S2LAB505.doctno, S2LAB505.certno, S2LAB502.ptid, S2LAB501.bedindt, " & _
              "        S2LAB501.wardid, S2LAB501.hosilid, S2LAB501.deptcd, S2HIS001." & F_PTNM & ", S2HIS001." & F_ADDRESS & ", " & _
              "        S2HIS006." & F_IENM & ", S2LAB502.rptdt, S2LAB502.age, S2LAB502.agediv, S2LAB502.sex, s2lab502.items, " & _
              "        S2LAB502.method, S2LAB502.others, S2LAB502.cmttxt, S2LAB502.recmd, S2LAB502.etcrst, " & _
              "        S2LAB503.testnm, S2LAB503.rstcd, S2LAB503.rstunit,S2LAB503.hldiv, S2LAB503.dpdiv, " & _
              "        S2LAB503.spcnm, S2LAB503.vfydttm "
    SqlStmt = SqlStmt & " From " & _
                      T_LAB505 & " S2LAB505, " & _
                      T_LAB501 & " S2LAB501, " & _
                      T_HIS001 & " S2HIS001, " & _
                      T_HIS006 & " S2HIS006, " & _
                      T_LAB502 & " S2LAB502, " & _
                      T_LAB503 & " S2LAB503 "
    SqlStmt = SqlStmt & " Where " & _
                                DBW("S2LAB501.ptid = ", m_PtId) & _
                      " AND " & DBW("S2LAB501.bedindt = ", m_BedinDt) & _
                      " AND S2LAB505.doctid = S2LAB501.rptid  " & _
                      " AND S2LAB501.ptid = S2HIS001." & F_PTID & _
                      " AND " & DBJ("S2HIS006." & F_ICD & " =* S2LAB501.disease") & _
                      " AND S2LAB501.ptid = S2LAB502.ptid  " & _
                      " AND S2LAB501.rptdt = S2LAB502.rptdt  " & _
                      " AND " & DBJ("S2LAB503.ptid =* S2LAB502.ptid") & _
                      " AND " & DBJ("S2LAB503.rptdt =* S2LAB502.rptdt")
                      
    SqlStmt = SqlStmt & Chr$(13) & Chr$(10) + "ORDER BY S2LAB503.seq"
    
    '## 커넥션 정보를 레지읽음
    With crReport
'        .Connect = DBConn.ConnectionString
        GetConnInfo
        .Connect = "DSN=" & medGetP(GetConnInfo, 1, ";") & ";UID=" & medGetP(GetConnInfo, 2, ";") & ";PWD=" & medGetP(GetConnInfo, 3, ";") & ";"
        
        .ReportFileName = InstallDir & "\LIS\Rpt\LabReport.rpt"
        If pOption = 1 Then
            .Destination = crptToWindow  '0 ' Window
            .WindowLeft = 0
            .WindowTop = 0
            .WindowState = crptMaximized
        Else
            .Destination = crptToPrinter
        End If
        .ParameterFields(1) = "bedindt;" & lblReqDt.Caption & ";TRUE"
        .SQLQuery = SqlStmt
        .Action = 1 ' Print
    End With
    
    Me.MousePointer = 0
    Exit Sub
    
PRINT_ERROR:
    Me.MousePointer = 0
    MsgBox Err.Description, vbExclamation
End Sub


Private Sub txtRcmd_Change()
    m_SaveFg = False
End Sub

'** OCS 결과전송 관련 프로시저 ===========================================================
Private Function OCS_Transfer_Result(ByVal pPtId As String, ByVal pMedDt As String, _
                                     ByVal pMedDept As String, ByVal pWardNo As String, _
                                     ByVal pRoomID As String, ByVal pAllItem As String, _
                                     ByVal pHighItem As String, ByVal pLowItem As String, _
                                     ByVal pTestDt As String, ByVal pTestID As String, _
                                     ByVal pRptDt As String, Optional ByVal pResult As String = "") As String
    Dim strResult       As String
    Dim strItems        As String
    Dim strItemT        As String
    Dim strItemM        As String
    Dim strMethod       As String
    Dim strMethodH      As String
    Dim strMethodB      As String
    Dim strComments     As String
    Dim strReComments   As String
    Dim strIPAdr        As String
    Dim strSupplemental As String
    Dim i               As Integer
    
    '-- ▣ 검사항목
    strItemT = "▣ 검사항목" & Chr(13)
    For i = 0 To lstTestList.ListCount - 1
        If Not lstTestList.Selected(i) Then
            strItems = strItems & lstTestList.List(i) & ","
        End If
    Next
    
    strItemM = strItemT & strItems
    
    '-- ▣ 검증방법
    strMethod = ""
    strMethodH = "▣ 검증방법 " & Chr(13)
    For i = 0 To chkMethod.Count - 1
        Select Case i
            Case 0
                If chkMethod(i).Value = 1 Then
                    strMethodB = strMethodB & "▶Calibration Verification"
                End If
                
            Case 1
                If chkMethod(i).Value = 1 Then
                    If strMethodB <> "" Then
                        strMethodB = strMethodB & Chr(13) & "▶Internal Quality Control"
                    Else
                        strMethodB = strMethodB & "▶Internal Quality Control"
                    End If
                End If
            
            Case 2
                If chkMethod(i).Value = 1 Then
                    If strMethodB <> "" Then
                        strMethodB = strMethodB & Chr(13) & "▶Delta Check Verification"
                    Else
                        strMethodB = strMethodB & "▶Delta Check Verification"
                    End If
                End If
            
            Case 3
                If chkMethod(i).Value = 1 Then
                    If strMethodB <> "" Then
                        strMethodB = strMethodB & Chr(13) & "▶Panic/Alert Value Verification"
                    Else
                        strMethodB = strMethodB & "▶Panic/Alert Value Verification"
                    End If
                End If
            
            Case 4
                If chkMethod(i).Value = 1 Then
                    If strMethodB <> "" Then
                        strMethodB = strMethodB & Chr(13) & "▶Repeat / Recheck"
                    Else
                        strMethodB = strMethodB & "▶Repeat / Recheck"
                    End If
                End If
            
            Case 5
                If chkMethod(i).Value = 1 Then
                    If strMethodB <> "" Then
                        strMethodB = strMethodB & Chr(13) & "▶Others;" & txtOthers.Text
                    Else
                        strMethodB = strMethodB & "▶Others;" & txtOthers.Text
                    End If
                End If
            
        End Select
    Next
    
    strMethod = strMethodH & strMethodB
    
    '-- ▣ 검증/판독 소견(Comments)
    strComments = ""
    If Trim(txtCmt.Text) <> "" Then
        strComments = "▣ 검증/판독 소견(Comments) " & Chr(13) & txtCmt.Text
    End If
    
    '-- ▣ 추천(Recommendation)
    strReComments = ""
    If Trim(txtRcmd.Text) <> "" Then
        strReComments = "▣ 추천(Recommendation) " & Chr(13) & txtRcmd.Text
    End If
    
    strResult = Mid$(strItemM & Chr(13) & pResult & Chr(13) & strMethod & Chr(13) & strComments & Chr(13) & strReComments, 1, 2000)
    
    '** ▣ Supplemental Report 결과 전송 추가 작업 By MGChoi 2005.03.03
    strSupplemental = GetSupplemental(pPtId, Mid(pRptDt, 1, 8))
    If strSupplemental <> "" Then
        strResult = strResult & Chr(13) & "▣ Supplemental Report" & Chr(13) & strSupplemental
    End If
    
    If Len(strResult) > 4000 Then
        GoTo ErrMsg
    End If
    
    If SqlOCSLABRESULT_FLAG(pPtId, pMedDt) = False Then
        '-- INSERT
'        OCS_Transfer_Result = SqlOCSLABRESULT_INSERT(pPtId, pMedDt, pMedDept, _
'                                pWardNo, pRoomID, strItems, pHighItem, pLowItem, "", _
'                                "", Mid$(strResult, 1, 4000), pTestDt, pTestID, pRptDt)
        
        '-- INSERT
        OCS_Transfer_Result = SqlOCSLABRESULT_INSERT(pPtId, pMedDt, pMedDept, _
                                pWardNo, pRoomID, strItems, pHighItem, pLowItem, "", _
                                "", strResult, pTestDt, pTestID, pRptDt)
                                
    Else
        '-- 전주예수병원IP Address =============
        strIPAdr = LocalIP_Address
        '=======================================
        
        '-- UPDATE
'        OCS_Transfer_Result = SqlOCSLABRESULT_UPDATE(pPtId, pMedDt, pMedDept, _
'                                pWardNo, pRoomID, strItems, pHighItem, pLowItem, "", _
'                                "", Mid$(strResult, 1, 2000), pTestID, strIPAdr, pTestDt)
                                
        '-- UPDATE
        OCS_Transfer_Result = SqlOCSLABRESULT_UPDATE(pPtId, pMedDt, pMedDept, _
                                pWardNo, pRoomID, strItems, pHighItem, pLowItem, "", _
                                "", strResult, pTestID, strIPAdr, pTestDt)
    End If
    
    Exit Function
    
ErrMsg:
    MsgBox "결과값이 초과 되었습니다. 최대 4000자 까지 입니다."
    OCS_Transfer_Result = ""
    
End Function

Private Function GetSupplemental(ByVal pPtId As String, ByVal pRptDt As String) As String
    Dim strSql  As String
    Dim Rs      As New ADODB.Recordset
    
    strSql = " select txtrst from " & T_LAB504 & _
             "  where ptid = " & DBS(pPtId) & _
             "    and rptdt = " & DBS(pRptDt) & _
             "    and mfyseq = (select max(a.mfyseq) from " & T_LAB504 & " a " & _
             "                   where a.ptid = " & DBS(pPtId) & _
             "                     and a.rptdt = " & DBS(pRptDt) & _
             "                   group by a.ptid, a.rptdt) "
    
    Rs.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        GetSupplemental = Rs.Fields("txtrst").Value & ""
    Else
        GetSupplemental = ""
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Function

'%  001. OCS측 결과테이블 RESULT INSERT (MDRESULT)
Private Function SqlOCSRESULT_INSERT(ByVal pPtId As String, ByVal pOrdDt As String, ByVal pOrdNo As Integer, _
                                    ByVal pTestCd As String, ByVal pWorkArea As String, ByVal pResult1 As String, _
                                    ByVal pResult2 As String, ByVal pFromRef As String, ByVal pToRef As String, _
                                    ByVal pRstUnit As String, ByVal pRemark As String, ByVal pRcvDt As String, _
                                    ByVal pVfyDt As String, ByVal pVfyId As String, ByVal pReadDt As String, _
                                    ByVal pReadID As String, ByVal pMfyFinalDt As String, ByVal pTestType As String, _
                                    ByVal pMfyID As String, ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    '-- Date Type을 위한 변수선언
    Dim strRcvDt  As String
    Dim strVfyDt  As String
    Dim strReadDt As String
    'Dim strMfyFDt As String
    Dim strMfyDt  As String
    
    '** 체크사항
    '   일단 (최종)변경일자와 수정일시는 동일한 일자로 한다. chngdate = editdate
    '-- 실시일자
    If pRcvDt <> "" Then
        strRcvDt = "TO_DATE(" & DBS(pRcvDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRcvDt = "''"
    End If
    
    '-- 보고일자
    If pVfyDt <> "" Then
        strVfyDt = "TO_DATE(" & DBS(pVfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strVfyDt = "''"
    End If
    
    '-- 판독일자
    If pReadDt <> "" Then
        strReadDt = "TO_DATE(" & DBS(pReadDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strReadDt = "''"
    End If
    
    '-- 수정일시
    If strMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSRESULT_INSERT = " insert into " & F_OCSRESULT & _
                          " (patno, orddate, ordseqno, examcode, slipcd, rslt1, rslt2, rsltupp, rsltlow, rsltunit, " & _
                          "  examtext, execdate, rsltdate, reptdr, readdate, readdr, chngdate, rslttype, editid, editip, editdate) " & _
                          " values (" & DBS(pPtId) & "," & "TO_DATE(" & DBS(pOrdDt) & ", 'yyyymmdd')" & "," & DBN(pOrdNo) & "," & _
                          DBS(pTestCd) & "," & DBS(pWorkArea) & "," & DBS(pResult1) & "," & _
                          DBS(pResult2) & "," & DBS(pToRef) & "," & DBS(pFromRef) & "," & _
                          DBS(pRstUnit) & "," & DBS(pRemark) & "," & strRcvDt & "," & _
                          strVfyDt & "," & DBS(pVfyId) & "," & strVfyDt & "," & DBS(pReadID) & "," & _
                          strMfyDt & "," & DBS(pTestType) & "," & DBS(pMfyID) & "," & _
                          DBS(pMfyIP) & "," & strMfyDt & ")"

End Function

'%  001. OCS측 종합검증/판독 결과 INSERT_UPDATE_FLAG (SLXVERIT) : True = UPDATE, False = INSERT
Private Function SqlOCSLABRESULT_FLAG(ByVal pPtId As String, ByVal pMedDt As String) As Boolean
    Dim Rs          As New ADODB.Recordset
    Dim strSql      As String
    Dim strMedDt    As String
    
    '-- 입원일자/진료일자
    strMedDt = "TO_DATE(" & DBS(pMedDt) & ", 'yyyymmdd')"
        
    strSql = " select * from " & T_SLXVERIT & _
             "  where patno = " & DBS(pPtId) & _
             "    and meddate = " & strMedDt
             
    Rs.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        SqlOCSLABRESULT_FLAG = True
    Else
        SqlOCSLABRESULT_FLAG = False
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Function

'%  002. OCS측 종합검증/판독 결과 INSERT (SLXVERIT)
Private Function SqlOCSLABRESULT_INSERT(ByVal pPtId As String, ByVal pMedDt As String, _
                                        ByVal pMedDept As String, ByVal pWardNo As String, _
                                        ByVal pRoomID As String, ByVal pActTest As String, _
                                        ByVal pHighTest As String, ByVal pLowTest As String, _
                                        ByVal pAbnormal As String, ByVal pTestMeth As String, _
                                        ByVal pResult As String, ByVal pTestDt As String, _
                                        ByVal pTestDoct As String, ByVal pRptDt As String) As String
    Dim strMedDt    As String
    Dim strTestDt   As String
    Dim strRptDt    As String
    
    '-- 진료일자
    strMedDt = "TO_DATE(" & DBS(pMedDt) & ", 'yyyymmdd')"
    
    '-- 검증일자
    If pTestDt <> "" Then
        strTestDt = "TO_DATE(" & DBS(pTestDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strTestDt = "''"
    End If
    
    '-- 출력일자
    If pRptDt <> "" Then
        strRptDt = "TO_DATE(" & DBS(pRptDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strRptDt = "''"
    End If
    
    SqlOCSLABRESULT_INSERT = " insert into " & T_SLXVERIT & _
                             " values (" & DBS(pPtId) & "," & strMedDt & "," & DBS(pMedDept) & _
                             "," & DBS(pWardNo) & "," & DBS(pRoomID) & "," & DBS(pActTest) & _
                             "," & DBS(pHighTest) & "," & DBS(pLowTest) & "," & DBS(pAbnormal) & _
                             "," & DBS(pTestMeth) & "," & DBS(pResult) & "," & strTestDt & _
                             "," & DBS(pTestDoct) & ", '', '', '', " & strRptDt & ")"

                             
End Function

'%  003. OCS측 종합검증/판독 결과 UPDATE (SLXVERIT)
Private Function SqlOCSLABRESULT_UPDATE(ByVal pPtId As String, ByVal pMedDt As String, _
                                        ByVal pMedDept As String, ByVal pWardNo As String, _
                                        ByVal pRoomID As String, ByVal pActTest As String, _
                                        ByVal pHighTest As String, ByVal pLowTest As String, _
                                        ByVal pAbnormal As String, ByVal pTestMeth As String, _
                                        ByVal pResult As String, ByVal pMfyID As String, _
                                        ByVal pMfyIP As String, ByVal pMfyDt As String) As String
    Dim strMedDt    As String
    Dim strTestDt   As String
    Dim strMfyDt    As String
    Dim strRptDt    As String
    
    '-- 진료일자
    strMedDt = "TO_DATE(" & DBS(pMedDt) & ", 'yyyymmdd')"
    
    '-- 수정일자
    If pMfyDt <> "" Then
        strMfyDt = "TO_DATE(" & DBS(pMfyDt) & ", 'yyyymmdd hh24:mi:ss')"
    Else
        strMfyDt = "''"
    End If
    
    SqlOCSLABRESULT_UPDATE = " update " & T_SLXVERIT & _
                             "    set donexam = " & DBS(pActTest) & _
                             ", highexam = " & DBS(pHighTest) & ", lowexam = " & DBS(pLowTest) & _
                             ", verifway = " & DBS(pTestMeth) & ", verifrslt = " & DBS(pResult) & _
                             ", editid = " & DBS(pMfyID) & ", editip = " & DBS(pMfyIP) & _
                             ", editdate = " & strMfyDt & ", prtdate = " & strMfyDt & _
                             "  where patno = " & DBS(pPtId) & _
                             "    and meddate = " & strMedDt
                             
End Function

'%  현재 작업중인 로컬 피씨의 IP Address를 확인한다.
Public Function LocalIP_Address() As String
    Dim Ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim Listing As MIB_IPADDRTABLE
    
    On Error GoTo ErrMsg
    
    GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Function
    
    ReDim bBytes(0 To Ret - 1) As Byte
    
    GetIpAddrTable bBytes(0), Ret, False
    
    CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
    
    LocalIP_Address = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
    
    Exit Function
    
ErrMsg:
    MsgBox "IP Address를 가져올 수 없습니다.", vbCritical
    
End Function

Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    CopyMemory myByte(0), longAddr, 4
    
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    
End Function
'=========================================================================================
