VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Begin VB.Form frmLabReport 
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
   Icon            =   "frmLabReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Tag             =   "ResultView2"
   Begin DRcontrol1.DrFrame fraSupp 
      Height          =   3240
      Left            =   120
      TabIndex        =   41
      Top             =   5730
      Visible         =   0   'False
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   5715
      Title           =   "결과보고일 : "
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00F2FBFB&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9405
         Style           =   1  '그래픽
         TabIndex        =   42
         Tag             =   "0"
         Top             =   90
         Width           =   270
      End
      Begin MedControls1.LisLabel lblMfyDt 
         Height          =   300
         Left            =   165
         TabIndex        =   43
         Top             =   60
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   529
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "Supplemental Report"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin RichTextLib.RichTextBox txtSupp 
         Height          =   2700
         Left            =   150
         TabIndex        =   44
         Top             =   420
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   4763
         _Version        =   393217
         BackColor       =   16054772
         ScrollBars      =   3
         TextRTF         =   $"frmLabReport.frx":038A
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
      Left            =   8625
      Style           =   1  '그래픽
      TabIndex        =   40
      Top             =   5385
      Width           =   1335
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
      Left            =   8880
      Style           =   1  '그래픽
      TabIndex        =   24
      Tag             =   "0"
      Top             =   1830
      Width           =   1095
   End
   Begin VB.Frame fraMethod 
      BackColor       =   &H00DBE6E6&
      Caption         =   "◈ 검증방법"
      Height          =   870
      Left            =   135
      TabIndex        =   15
      Top             =   4530
      Width           =   9885
      Begin VB.Frame Frame3 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Enabled         =   0   'False
         Height          =   600
         Left            =   285
         TabIndex        =   32
         Top             =   225
         Width           =   9330
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
            Left            =   6555
            TabIndex        =   39
            Top             =   300
            Width           =   2565
         End
         Begin VB.CheckBox chkMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Others;"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   5625
            TabIndex        =   38
            Top             =   300
            Width           =   900
         End
         Begin VB.CheckBox chkMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Repeat / Recheck"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   37
            Top             =   300
            Width           =   1830
         End
         Begin VB.CheckBox chkMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Panic/Alert Value Verification"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   75
            TabIndex        =   36
            Top             =   300
            Width           =   2775
         End
         Begin VB.CheckBox chkMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Delta Check Verification"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   5625
            TabIndex        =   35
            Top             =   45
            Width           =   2685
         End
         Begin VB.CheckBox chkMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Internal Quality Control"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   34
            Top             =   45
            Width           =   2220
         End
         Begin VB.CheckBox chkMethod 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Calibration Verification"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   33
            Top             =   45
            Width           =   2145
         End
      End
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
      TabIndex        =   12
      Top             =   1350
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
      Height          =   2220
      Left            =   105
      TabIndex        =   0
      Top             =   2175
      Width           =   1920
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   4770
      TabIndex        =   3
      Top             =   480
      Width           =   3885
      Begin MedControls1.LisLabel lblWardId 
         Height          =   210
         Left            =   1365
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   16
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   150
      TabIndex        =   2
      Top             =   480
      Width           =   4545
      Begin MedControls1.LisLabel lblPtId 
         Height          =   210
         Left            =   1545
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   26
         Top             =   945
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D1D8D3&
         Height          =   225
         Left            =   1545
         TabIndex        =   27
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   210
         Width           =   1080
      End
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   2325
      Left            =   2055
      TabIndex        =   14
      Top             =   2175
      Width           =   7935
      _Version        =   196608
      _ExtentX        =   13996
      _ExtentY        =   4101
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
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
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmLabReport.frx":042F
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   25
      Top             =   2175
      Visible         =   0   'False
      Width           =   7935
   End
   Begin RichTextLib.RichTextBox txtRcmd 
      Height          =   1140
      Left            =   105
      TabIndex        =   30
      Top             =   7845
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   2011
      _Version        =   393217
      BackColor       =   16776183
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmLabReport.frx":0B83
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
   Begin RichTextLib.RichTextBox txtCmt 
      Height          =   1830
      Left            =   105
      TabIndex        =   31
      Top             =   5700
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   3228
      _Version        =   393217
      BackColor       =   16776191
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmLabReport.frx":0C28
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ 추천 (Recommendation)"
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
      Left            =   150
      TabIndex        =   29
      Top             =   7650
      Width           =   2160
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "◈ 검증/판독 소견 (Comments)"
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
      Left            =   255
      TabIndex        =   28
      Top             =   5520
      Width           =   2520
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
      Left            =   3195
      TabIndex        =   1
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
      TabIndex        =   13
      Top             =   1935
      Width           =   4050
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
      TabIndex        =   11
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
      Left            =   2670
      Shape           =   4  '둥근 사각형
      Top             =   90
      Width           =   4515
   End
End
Attribute VB_Name = "frmLabReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MySql As New clsLISSqlStatement   'Sql문 클래스
Private objCmtSql As New clsLISSqlReview     'Sql문 클래스
Private MyPatient As New clsPatient ' clsLisPatient

Private m_DoneFg As Boolean
Private m_PtId As String
Private m_BedinDt As String
Private m_QueryFg As Boolean
Private m_SaveFg As Boolean

Public Property Get DoneFg() As Boolean
    DoneFg = m_DoneFg
End Property
Public Property Let DoneFg(ByVal vData As Boolean)
    m_DoneFg = vData
End Property

Public Property Get PTid() As String
    PTid = m_PtId
End Property
Public Property Let PTid(ByVal vData As String)
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
    Set frmLabReport = Nothing
End Sub

Public Sub StartQuery()
    
    Call GetPtInfo
    Call GetReport
      
    m_QueryFg = True
    
End Sub


Private Sub GetPtInfo()

    If m_PtId = "" Then Exit Sub
    
    With MyPatient
'        If .PtntQuery(m_PtId, m_BedinDt) Then
        If .GETPatient(m_PtId) Then
            lblPtId.Caption = m_PtId
            lblPtNm.Caption = .PtNm
            lblPtSexAge.Caption = .Sex & " / " & .Age & "  " & .AGEDIV
            lblDiagNm.Caption = .InDiseaCd
            lblDiagNm.ToolTipText = .InDiseaCd
            lblWardId.Caption = .WardId
            lblDeptNm.Caption = .DeptNm
            lblReqDt.Caption = Format(.BedinDt, CS_DateLongMask)
        End If
    End With

End Sub

Private Sub GetReport()
    
    Dim SqlStmt As String
    Dim RS As Recordset
    Dim strItems As String
    Dim strRptDt As String
    Dim I As Integer
    
    SqlStmt = " select a.donefg, b.* from " & T_LAB501 & " a, " & T_LAB502 & " b " & _
              " where a.ptid = '" & m_PtId & "' " & _
              " and   a.bedindt = '" & m_BedinDt & "' " & _
              " and   b.ptid = a.ptid " & _
              " and   b.rptdt = a.rptdt "
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If Not RS.EOF Then
        strRptDt = "" & RS.Fields("RptDt").Value
        lblRptDt.Caption = Format(strRptDt, CS_DateLongMask)
        strItems = "" & RS.Fields("Items").Value
        txtEtcResult.Text = "" & RS.Fields("EtcRst").Value
        For I = 1 To chkMethod.Count
            chkMethod(I - 1).Value = Val(Mid("" & RS.Fields("Method").Value, I, 1))
        Next
        txtOthers.Text = "" & RS.Fields("Others").Value
        txtCmt.Text = "" & RS.Fields("cmttxt").Value
        txtRcmd.Text = "" & RS.Fields("Recmd").Value
        lstTestList.Clear
        While (strItems <> "")
            lstTestList.AddItem medShift(strItems, ",")
        Wend
    End If
    Set RS = Nothing
    
    SqlStmt = "select * from " & T_LAB503 & " where ptid = " & m_PtId & " and rptdt = '" & strRptDt & "'"
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    With tblResult
        
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = RS.RecordCount
        .Row = 0
        
        While (Not RS.EOF)
            .Row = .Row + 1
            .Col = 1: .Value = 0
            .Col = 2: .Value = Trim("" & RS.Fields("TestNm").Value)
            .Col = 3: .ForeColor = DCM_Brown           '-- 결과명(코드일 경우..)
                      .Value = Trim("" & RS.Fields("RstCd").Value)
            .Col = 4: .Value = Trim("" & RS.Fields("RstUnit").Value)         '-- 결과단위
            .Col = 5       '-- High / Low
                      .Value = ""
                      If Trim("" & RS.Fields("HLDiv").Value) = "H" Then .Value = "▲": .ForeColor = DCM_LightRed
                      If Trim("" & RS.Fields("HLDiv").Value) = "L" Then .Value = "▼": .ForeColor = DCM_LightBlue
                      If Trim("" & RS.Fields("HLDiv").Value) = "*" Then .Value = "*": .ForeColor = vbRed
                      .Value = .Value & Trim("" & RS.Fields("DPDiv").Value)    ': .ForeColor = vbRed                '빨간색
            .Col = 6: .Value = Trim("" & RS.Fields("SpcNm").Value)    '검체명
            .Col = 7: .Value = Trim("" & RS.Fields("VfyDtTm").Value)         '-- 보고일시
            .Col = 8: .Value = Trim("" & RS.Fields("HLDiv").Value)
            .Col = 9: .Value = Trim("" & RS.Fields("DPDiv").Value)
            .Col = 10: .Value = Trim("" & RS.Fields("TestNm").Value)
            
            RS.MoveNext
        
        Wend
        
        .RowHeight(-1) = 11
        .ReDraw = True
        
    End With
    
    m_SaveFg = True
    Set RS = Nothing
    
End Sub

'-- 예수병원 추가루틴 =================================================================
' - By M.G.Choi 2005.05.30
Private Sub cmdSupp_Click()
    Dim strPtId     As String
    Dim strRptDt    As String
    
    lblMfyDt.Caption = "Supplemental Report"
    txtSupp.Text = ""
    
    strPtId = Trim(lblPtId.Caption)
    strRptDt = Format(lblRptDt.Caption, CS_DateDbFormat)
    
    Call GetSuppText(strPtId, strRptDt)
    
    With fraSupp
        .Visible = True
        .ZOrder 0
    End With
    
End Sub

Public Sub GetSuppText(ByVal pPtID As String, ByVal pRptDt As String)

    Dim SqlStmt As String
    Dim RS As Recordset
    
    SqlStmt = " select * from " & T_LAB504 & _
              " where " & DBW("ptid=", pPtID) & _
              " and   " & DBW("rptdt=", pRptDt) & _
              " order by mfyseq"
              
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    If Not RS.EOF Then
        lblMfyDt.Caption = "Supplemental Report" & "(수 정 일:" & Format("" & RS.Fields("MfyDt").Value, CS_DateLongMask) & ")"
        txtSupp.Text = Trim("" & RS.Fields("TxtRst").Value)
    Else
        lblMfyDt.Caption = "Supplemental Report"
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub cmdClose_Click()
    fraSupp.Visible = False
End Sub
'======================================================================================


Private Sub txtEtcResult_Change()
    If Trim(txtEtcResult.Text) = "" Then
        cmdEtcResult.BackColor = &HE0E0E0
    Else
        cmdEtcResult.BackColor = &HFFFBFF
    End If
End Sub
