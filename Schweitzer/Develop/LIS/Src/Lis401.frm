VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Begin VB.Form frm401ResultView 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9105
   ClientLeft      =   285
   ClientTop       =   180
   ClientWidth     =   14715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "Lis401.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   14715
   WindowState     =   2  '최대화
   Begin VB.PictureBox picPtList 
      Align           =   3  '왼쪽 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H00D7E6E6&
      DragMode        =   1  '자동
      Height          =   7725
      Left            =   0
      ScaleHeight     =   7665
      ScaleWidth      =   4245
      TabIndex        =   5
      Top             =   1380
      Width           =   4300
      Begin VB.CheckBox chkAllWard 
         BackColor       =   &H00D7E6E6&
         Caption         =   "전체병동"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2355
         TabIndex        =   71
         Top             =   150
         Width           =   1035
      End
      Begin VB.CheckBox chkVerified 
         BackColor       =   &H00D7E6E6&
         Caption         =   "금일 결과보고 대상만 검색"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00553755&
         Height          =   225
         Left            =   1785
         TabIndex        =   59
         Top             =   435
         Width           =   2460
      End
      Begin VB.Frame fraSearch 
         BackColor       =   &H00D7E6E6&
         Height          =   645
         Left            =   30
         TabIndex        =   54
         Tag             =   "136"
         Top             =   600
         Width           =   4200
         Begin VB.OptionButton optSort 
            BackColor       =   &H00D7E6E6&
            Caption         =   "&ID"
            Height          =   240
            Index           =   0
            Left            =   1995
            TabIndex        =   57
            Tag             =   "15304"
            Top             =   300
            Width           =   495
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00D7E6E6&
            Caption         =   "&Name"
            Height          =   255
            Index           =   1
            Left            =   2505
            TabIndex        =   56
            Tag             =   "15305"
            Top             =   285
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.TextBox txtSearchKey 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            MaxLength       =   10
            TabIndex        =   55
            Text            =   "테"
            Top             =   240
            Width           =   1830
         End
         Begin VB.Label lblReset 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            BackStyle       =   0  '투명
            Caption         =   "Reset"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3570
            MouseIcon       =   "Lis401.frx":08CA
            MousePointer    =   99  '사용자 정의
            TabIndex        =   58
            Top             =   285
            Width           =   495
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00808080&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  '단색
            Height          =   285
            Index           =   1
            Left            =   3465
            Shape           =   4  '둥근 사각형
            Top             =   255
            Width           =   675
         End
      End
      Begin MSComctlLib.ListView lvwPtList 
         Height          =   6450
         Left            =   15
         TabIndex        =   53
         Top             =   1230
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   11377
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16643054
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblWardId 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "병동선택"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00553755&
         Height          =   180
         Left            =   3450
         MouseIcon       =   "Lis401.frx":0BD4
         MousePointer    =   99  '사용자 정의
         TabIndex        =   60
         ToolTipText     =   "Click하시면 마감시간을 수정할 수 있습니다."
         Top             =   165
         Width           =   720
      End
      Begin VB.Label lblPtList 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Patient List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         TabIndex        =   8
         Tag             =   "106"
         Top             =   120
         Width           =   1185
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         FillColor       =   &H00E8F7F7&
         FillStyle       =   0  '단색
         Height          =   270
         Left            =   3405
         Shape           =   4  '둥근 사각형
         Top             =   120
         Width           =   795
      End
   End
   Begin VB.PictureBox picRstText 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFF7&
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   8010
      ScaleHeight     =   1515
      ScaleWidth      =   6660
      TabIndex        =   66
      Top             =   6435
      Width           =   6690
      Begin RichTextLib.RichTextBox txtRstCmt 
         Height          =   1440
         Left            =   75
         TabIndex        =   67
         Top             =   45
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   2540
         _Version        =   393217
         BackColor       =   16777207
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Lis401.frx":0EDE
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
   Begin VB.PictureBox picFootNote 
      Appearance      =   0  '평면
      BackColor       =   &H00EFFEFE&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   8010
      ScaleHeight     =   1035
      ScaleWidth      =   6660
      TabIndex        =   64
      Top             =   7995
      Width           =   6690
      Begin RichTextLib.RichTextBox txtSamCmt 
         Height          =   900
         Left            =   75
         TabIndex        =   65
         Top             =   30
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   1588
         _Version        =   393217
         BackColor       =   15728382
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Lis401.frx":126B
         MouseIcon       =   "Lis401.frx":15F8
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
   Begin FPSpread.vaSpread tblResult 
      Height          =   4230
      Left            =   8010
      TabIndex        =   63
      Top             =   2190
      Width           =   6690
      _Version        =   196608
      _ExtentX        =   11800
      _ExtentY        =   7461
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   3
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   12
      OperationMode   =   1
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis401.frx":175A
      UnitType        =   0
      UserResize      =   0
      VisibleCols     =   8
      VisibleRows     =   22
      TextTip         =   4
   End
   Begin VB.PictureBox picResult 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3F5F8&
      Height          =   825
      Left            =   8010
      ScaleHeight     =   765
      ScaleWidth      =   6720
      TabIndex        =   12
      Top             =   1380
      Width           =   6780
      Begin VB.CheckBox chkSamCmt 
         BackColor       =   &H00F3F5F8&
         Caption         =   "Sample Comment"
         ForeColor       =   &H00553755&
         Height          =   255
         Left            =   4785
         TabIndex        =   70
         Tag             =   "40205"
         Top             =   240
         Value           =   1  '확인
         Width           =   1815
      End
      Begin VB.CheckBox chkRstCmt 
         BackColor       =   &H00F3F5F8&
         Caption         =   "Result   Comment"
         ForeColor       =   &H00553755&
         Height          =   255
         Left            =   4785
         TabIndex        =   69
         Tag             =   "40204"
         Top             =   495
         Value           =   1  '확인
         Width           =   1815
      End
      Begin VB.Label lblSpecimenNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Serum"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   795
         TabIndex        =   25
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lblSpecimen 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체 : "
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
         Left            =   150
         TabIndex        =   23
         Tag             =   "157"
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblWorkArea 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "Chemistry"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DF6A3E&
         Height          =   225
         Left            =   1305
         TabIndex        =   21
         Top             =   135
         Width           =   1110
      End
      Begin VB.Label lblResults 
         BackStyle       =   0  '투명
         Caption         =   "Results  -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         TabIndex        =   19
         Tag             =   "19908"
         Top             =   120
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00DFE3E8&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   0
      ScaleHeight     =   1380
      ScaleWidth      =   14715
      TabIndex        =   6
      Top             =   0
      Width           =   14715
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   12990
         Style           =   1  '그래픽
         TabIndex        =   52
         Tag             =   "40102"
         Top             =   315
         Width           =   1410
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
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
         Height          =   420
         Left            =   12990
         Style           =   1  '그래픽
         TabIndex        =   24
         Tag             =   "128"
         Top             =   765
         Width           =   1410
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "&Report"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13395
         TabIndex        =   22
         Tag             =   "40102"
         Top             =   315
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.CheckBox chkPtList 
         BackColor       =   &H00DFE3E8&
         Caption         =   "환자검색 리스트(&S)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Tag             =   "40101"
         Top             =   1095
         Width           =   2460
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   0
         Top             =   75
         Width           =   1410
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   255
         Left            =   1335
         TabIndex        =   26
         Top             =   420
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "김미경"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblReceiverNm 
         Height          =   255
         Left            =   7575
         TabIndex        =   27
         Top             =   720
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "2"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCollectorNm 
         Height          =   255
         Left            =   7575
         TabIndex        =   32
         Top             =   405
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "2"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCollectDt 
         Height          =   255
         Left            =   10560
         TabIndex        =   33
         Top             =   420
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblVerifierNm 
         Height          =   255
         Left            =   7575
         TabIndex        =   34
         Top             =   1035
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "2"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblOrdDt 
         Height          =   255
         Left            =   10560
         TabIndex        =   35
         Top             =   105
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblReceiveDt 
         Height          =   255
         Left            =   10560
         TabIndex        =   36
         Top             =   735
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblVerifyDt 
         Height          =   255
         Left            =   10560
         TabIndex        =   37
         Top             =   1050
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "3"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   255
         Left            =   7575
         TabIndex        =   38
         Top             =   90
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "2"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   255
         Left            =   4620
         TabIndex        =   43
         Top             =   405
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "1"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblBedoutDt 
         Height          =   255
         Left            =   4620
         TabIndex        =   44
         Top             =   1050
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "1"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblBedinDt 
         Height          =   255
         Left            =   4620
         TabIndex        =   45
         Top             =   735
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "1"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   255
         Left            =   4620
         TabIndex        =   46
         Top             =   90
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   450
         BackColor       =   15988216
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
         Caption         =   "1"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   285
         Left            =   1365
         Shape           =   4  '둥근 사각형
         Top             =   60
         Width           =   1440
      End
      Begin VB.Label lblLocation1 
         BackStyle       =   0  '투명
         Caption         =   "병     실 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   3720
         TabIndex        =   50
         Tag             =   "102"
         Top             =   435
         Width           =   1110
      End
      Begin VB.Label lblDept 
         BackStyle       =   0  '투명
         Caption         =   "진 료 과 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   3720
         TabIndex        =   49
         Tag             =   "40304"
         Top             =   135
         Width           =   1110
      End
      Begin VB.Label lblADM 
         BackStyle       =   0  '투명
         Caption         =   "입 원 일 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   3705
         TabIndex        =   48
         Tag             =   "40302"
         Top             =   765
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "퇴 원 일 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   3705
         TabIndex        =   47
         Tag             =   "40302"
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label lblDoct 
         BackStyle       =   0  '투명
         Caption         =   "처 방 의 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   6630
         TabIndex        =   42
         Tag             =   "107"
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label lblCollector 
         BackStyle       =   0  '투명
         Caption         =   "채 혈 자 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   6630
         TabIndex        =   41
         Tag             =   "40103"
         Top             =   450
         Width           =   1050
      End
      Begin VB.Label lblReceiver 
         BackStyle       =   0  '투명
         Caption         =   "접 수 자 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   6630
         TabIndex        =   40
         Tag             =   "19905"
         Top             =   765
         Width           =   1050
      End
      Begin VB.Label lblVerifier 
         BackStyle       =   0  '투명
         Caption         =   "보 고 자 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   6630
         TabIndex        =   39
         Tag             =   "40111"
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2670
         TabIndex        =   30
         Top             =   810
         Width           =   120
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2340
         TabIndex        =   29
         Top             =   810
         Width           =   180
      End
      Begin VB.Label lblSex 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "여자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1455
         TabIndex        =   28
         Top             =   795
         Width           =   360
      End
      Begin VB.Label lblOrdTm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처방일시 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   9585
         TabIndex        =   16
         Tag             =   "40106"
         Top             =   135
         Width           =   840
      End
      Begin VB.Label lblColTm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "채혈일시 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   9585
         TabIndex        =   15
         Tag             =   "19902"
         Top             =   450
         Width           =   840
      End
      Begin VB.Label lblRcvTm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "접수일시 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   9585
         TabIndex        =   14
         Tag             =   "40107"
         Top             =   780
         Width           =   840
      End
      Begin VB.Label lblRptTm 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보고일시 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   9585
         TabIndex        =   13
         Tag             =   "40108"
         Top             =   1095
         Width           =   840
      End
      Begin VB.Label lblPtId 
         BackStyle       =   0  '투명
         Caption         =   "환자   ID :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   405
         TabIndex        =   11
         Tag             =   "105"
         Top             =   135
         Width           =   930
      End
      Begin VB.Label lblName 
         BackStyle       =   0  '투명
         Caption         =   "성      명 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   225
         Left            =   390
         TabIndex        =   10
         Tag             =   "103"
         Top             =   480
         Width           =   945
      End
      Begin VB.Label lblSexAge 
         BackStyle       =   0  '투명
         Caption         =   "성별/나이 :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Tag             =   "108"
         Top             =   795
         Width           =   945
      End
      Begin VB.Label Label8 
         Appearance      =   0  '평면
         BackColor       =   &H00F3F5F8&
         Caption         =   "            /"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1320
         TabIndex        =   31
         Top             =   750
         Width           =   2010
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         FillColor       =   &H00DFE3E8&
         FillStyle       =   0  '단색
         Height          =   1380
         Left            =   30
         Shape           =   4  '둥근 사각형
         Top             =   0
         Width           =   14640
      End
   End
   Begin RichTextLib.RichTextBox rtfResult 
      Height          =   7665
      Left            =   7995
      TabIndex        =   68
      Top             =   1410
      Visible         =   0   'False
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   13520
      _Version        =   393217
      BackColor       =   16777207
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9000
      TextRTF         =   $"Lis401.frx":3440
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
   Begin VB.PictureBox picOrder 
      BackColor       =   &H00F3F5F8&
      Height          =   810
      Left            =   15
      ScaleHeight     =   750
      ScaleWidth      =   7935
      TabIndex        =   7
      Top             =   1380
      Width           =   7995
      Begin VB.OptionButton optQueryKey 
         BackColor       =   &H00F3F5F8&
         Caption         =   "보고일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3660
         TabIndex        =   61
         Tag             =   "15305"
         Top             =   105
         Width           =   840
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00CCFFFF&
         Caption         =   "Re&fresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6780
         MaskColor       =   &H00C0FFFF&
         Style           =   1  '그래픽
         TabIndex        =   51
         Tag             =   "128"
         Top             =   330
         Width           =   1140
      End
      Begin VB.OptionButton optQueryKey 
         BackColor       =   &H00F3F5F8&
         Caption         =   "접수일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4530
         TabIndex        =   1
         Tag             =   "15304"
         Top             =   120
         Width           =   885
      End
      Begin VB.OptionButton optQueryKey 
         BackColor       =   &H00F3F5F8&
         Caption         =   "처방일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5445
         TabIndex        =   2
         Tag             =   "15305"
         Top             =   105
         Width           =   840
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   300
         Left            =   3660
         TabIndex        =   3
         Top             =   405
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24969219
         CurrentDate     =   36328
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   300
         Left            =   5220
         TabIndex        =   4
         Top             =   405
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24969219
         CurrentDate     =   36328
      End
      Begin VB.PictureBox picOrdDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00F3F5F8&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   45
         ScaleHeight     =   315
         ScaleWidth      =   3525
         TabIndex        =   72
         Top             =   435
         Width           =   3525
         Begin VB.CheckBox ChkDivAll 
            BackColor       =   &H00F3F5F8&
            Caption         =   "전체"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C56152&
            Height          =   225
            Left            =   60
            TabIndex        =   76
            Top             =   45
            Width           =   780
         End
         Begin VB.OptionButton optOrdDiv 
            BackColor       =   &H00FFF2EE&
            Caption         =   "임상병리"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   0
            Left            =   855
            Style           =   1  '그래픽
            TabIndex        =   75
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton optOrdDiv 
            BackColor       =   &H00EDE2ED&
            Caption         =   "해부병리"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   1
            Left            =   1695
            Style           =   1  '그래픽
            TabIndex        =   74
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton optOrdDiv 
            BackColor       =   &H00F4FDF5&
            Caption         =   "혈액은행"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   2
            Left            =   2535
            Style           =   1  '그래픽
            TabIndex        =   73
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   225
         Left            =   5055
         TabIndex        =   20
         Tag             =   "40110"
         Top             =   405
         Width           =   105
      End
      Begin VB.Label lblOrders 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Orders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         TabIndex        =   18
         Tag             =   "155"
         Top             =   90
         Width           =   735
      End
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   6870
      Left            =   15
      TabIndex        =   62
      Top             =   2190
      Width           =   7980
      _Version        =   196608
      _ExtentX        =   14076
      _ExtentY        =   12118
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   32
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   -2147483633
      ShadowText      =   0
      SpreadDesigner  =   "Lis401.frx":3814
      TextTip         =   4
   End
End
Attribute VB_Name = "frm401ResultView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'% 폼단위 전역변수 선언

'-------------------------------------
'해부병리/혈액은행 결과조회 여부
'-------------------------------------
#Const AllowAPSResultReview = True
#Const AllowBBSResultReview = True
'-------------------------------------

Private objPatient As New clsPatient     '환자 클래스
Private objSql As New clsLISSqlReview    'Sql문 클래스
Private ClearFg As Boolean
Private OrderFg As Boolean
Private ResultFg As Boolean
Private MsgFg As Boolean

Private WithEvents objMyList As clsS2DLP
Attribute objMyList.VB_VarHelpID = -1
'Private WithEvents objText As Form

Private OldRow As Long
Private OldOrdDiv As String
Private aryMesg() As String

Public PtFg As Boolean
Public QueryFg As Boolean

Private mvarDeptCd As String

Private Const lngMaxRows = 29
Private Const lngRowHeight = 11.5

Public Event LastFormUnload()
Public Event ThisFormUnload()

Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property

Private Sub chkAllWard_Click()
    If chkAllWard.Value = 0 Then
        chkVerified.Value = 0
        chkVerified.Enabled = False
    Else
        chkVerified.Enabled = True
    End If
End Sub

Private Sub ChkDivAll_Click()
    If ChkDivAll.Value = 1 Then
        optOrdDiv(0).Value = False
        optOrdDiv(1).Value = False
        optOrdDiv(2).Value = False
    Else
        optOrdDiv(0).Value = True
    End If
End Sub

Private Sub lblWardId_Click()
    
    Set objMyList = New clsS2DLP
    
    With objMyList
        .Caption = "병동 조회"
        .Tag = "WardId"
        .HeadName = "병동코드,병동명"
        Call .ListPop(, 1640, 10550, ObjLISComCode.WardId)
    End With
    
End Sub

Private Sub objMyList_SendCode(ByVal SelString As String)
    If objMyList.Tag = "WardId" Then
        lblWardId.Caption = Trim(medGetP(SelString, 1, ";"))
        lblWardId.Tag = "1"
        mvarDeptCd = lblWardId.Caption
        chkVerified.Enabled = True
        If chkVerified.Value = 1 Then Call txtSearchKey_KeyPress(vbKeyReturn)
    End If
End Sub

'Private Sub objText_Click()
'    Set objText = Nothing
'End Sub

'% 환자리스트 Display 여부
Private Sub chkPtList_Click()
    On Error GoTo Err_Trap
    If chkPtList.Value = 1 Then
        lblWardId.Caption = mvarDeptCd
        picPtList.Visible = True
        picPtList.Width = 4290
        picOrder.Left = picPtList.Width
        tblOrdSheet.Left = picOrder.Left
        picResult.Left = picPtList.Width + picOrder.Width
        'picResult.Width = picResult.Width - picPtList.Width + 50
        tblResult.Left = picResult.Left
        picRstText.Left = picResult.Left
        picFootNote.Left = picResult.Left
        txtSearchKey.SetFocus
    ElseIf chkPtList.Value = 0 Then
        picPtList.Visible = False
        picOrder.Left = 0
        tblOrdSheet.Left = picOrder.Left
        picResult.Left = picOrder.Width
        tblResult.Left = picResult.Left
        picRstText.Left = picResult.Left
        picFootNote.Left = picResult.Left
'        picResult.Width = picResult.Width + picPtList.Width + 50
    End If
    Exit Sub
Err_Trap:
End Sub

'% 텍스트 결과내역 박스 Display 여부
Private Sub chkRstCmt_Click()
    If chkRstCmt.Value = 1 And picRstText.Visible = False Then
        picRstText.Visible = True
        tblResult.Height = tblResult.Height - picRstText.Height
    ElseIf chkRstCmt.Value = 0 And picRstText.Visible = True Then
        picRstText.Visible = False
        tblResult.Height = tblResult.Height + picRstText.Height
    End If
End Sub

'% 풋노트, 검체리마크 박스 Display 여부
Private Sub chkSamCmt_Click()
    If chkSamCmt.Value = 1 And picFootNote.Visible = False Then
        picFootNote.Visible = True
        tblResult.Height = tblResult.Height - picFootNote.Height
        picRstText.Top = picRstText.Top - picFootNote.Height
    ElseIf chkSamCmt.Value = 0 And picFootNote.Visible = True Then
        picFootNote.Visible = False
        tblResult.Height = tblResult.Height + picFootNote.Height
        picRstText.Top = picRstText.Top + picFootNote.Height
    End If
End Sub

Private Sub chkVerified_Click()
    Call txtSearchKey_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdClear_Click()
'    picOrder.Width = 8000
    cmdRefresh.Left = picOrder.Width - cmdRefresh.Width - 50
    Call ClearRtn
    txtPtId.Text = ""
    On Error GoTo Err_Trap
    txtPtId.SetFocus
Err_Trap:
End Sub

'%종료
Private Sub cmdExit_Click()
    Unload Me
    Set objSql = Nothing
    Set objMyList = Nothing
    Set objPatient = Nothing
    RaiseEvent ThisFormUnload
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdRefresh_Click()
   
    '% 처방조회
    OldRow = 0
    Call dtpToDate_KeyDown(vbKeyReturn, 0)

End Sub

'% 레포트 출력
Private Sub cmdReport_Click()
    'frmPreview.Show
End Sub


'% 조회기간 입력 (From Date)
Private Sub dtpFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Trap
    If KeyCode = vbKeyReturn Then dtpToDate.SetFocus
Err_Trap:
End Sub


Private Sub dtpToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Err_Trap
    
    If KeyCode = vbKeyReturn Then
        If dtpToDate.Value < dtpFromDate.Value Then
            MsgBox "기간 입력 오류입니다. 날짜를 조정하십시오..", vbExclamation, "입력오류"
            dtpFromDate.SetFocus
            Exit Sub
        End If
        '% 처방조회
        cmdRefresh.Enabled = False
        dtpFromDate.Enabled = False
        dtpToDate.Enabled = False
        
        Call FieldClear
        Call DisplayOrders
        
        ResultFg = False
        cmdRefresh.Enabled = True
        dtpFromDate.Enabled = True
        dtpToDate.Enabled = True
        
        If OrderFg Then
            tblOrdSheet.SetFocus
        Else
            dtpFromDate.SetFocus
        End If
    End If
    Exit Sub
    
Err_Trap:
    Resume Next
End Sub

'% 환자ID, 처방일(채혈일)을 기준으로 처방내역을 검색한다.
Private Sub DisplayOrders()

    Dim i As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim RecordCnt As Integer
    Dim tmpRs As New DrRecordSet
    Dim SvOrdDt As String, SvOrdNo As String, SvDoctNm As String, SvSpcNm As String
    Dim pWorkArea As String, pAccDt As String, pAccSeq As String
    Dim strStsCd As String, strStsNm As String, lngColor As Long
    Dim iBtnFg As Long, strOrdDiv As String, strTestDiv As String
    Dim objStatus As New clsProgress
    Dim strUnit As String
    Dim strSelDiv As String
    
    QueryFg = True
    
    Call TableClear
    Call ResultClear
    
    'Status Bar Popup
    Me.Enabled = False
    MouseRunning  '13
   
   
    DoEvents
    With objStatus
        .CaptionOn = False
        .Mode = 0
        .Msg = lblPtNm.Caption & " 님의 처방내역을 검색중입니다..."
        .Min = 0
        .Max = 80
        .Value = 40

'        .Visible = True

        '.ZOrder 0
    End With
    
    DoEvents
    
    '처방일/접수일/보고일 기준
    Dim strKeyFld As String
    'strKeyFld = IIf(optQueryKey(1).Value, "orddt", "rcvdt")
    If optQueryKey(0).Value = True Then
        strKeyFld = "rcvdt"
    ElseIf optQueryKey(1).Value = True Then
        strKeyFld = "orddt"
    Else
        strKeyFld = "examdt"
    End If
        
    'pooh 수정  0-전체, 1-임상, 2-해부, 4-혈액
    If ChkDivAll.Value = 1 Then
        strSelDiv = "0"
    Else
        If optOrdDiv(0).Value = True Then
            strSelDiv = "1"
        ElseIf optOrdDiv(1).Value = True Then
            strSelDiv = "2"
        Else
            strSelDiv = "4"
        End If
    End If
        
    SqlStmt = objSql.SqlQueryNEWOrders(txtPtId.Text, strKeyFld, Format(dtpFromDate.Value, CS_DateDbFormat), _
                                   Format(dtpToDate.Value, CS_DateDbFormat), strSelDiv)
    
On Error GoTo Err_Trap
    
    'Query
    ColCnt = tmpRs.OpenCursor(DBConn, SqlStmt)
    
    If ColCnt = 0 Then GoTo Err_Trap
    
    objStatus.Max = 100
    objStatus.Value = objStatus.Max
   
    SvOrdDt = "": SvOrdNo = "": SvDoctNm = "": SvSpcNm = ""
    Erase aryMesg
   
    With tblOrdSheet
      
        .MaxRows = 0
        
        RecordCnt = 0
      
        While (tmpRs.FetchCursor(ColCnt))
            
            RecordCnt = RecordCnt + 1
            ReDim Preserve aryMesg(RecordCnt)   'Message Array ...
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            If SvOrdDt <> Trim("" & tmpRs.GetValue("OrdDate")) Then
                .Col = enREVIEW1.tcORDDT:   .Value = Trim("" & tmpRs.GetValue("OrdDate"))    '처방일
                .Col = enREVIEW1.tcORDNO:   .Value = Trim("" & tmpRs.GetValue("OrdNo"))      '처방번호
                .Col = enREVIEW1.tcSPCNM:   .Value = Trim("" & tmpRs.GetValue("SpcNm"))      '검체명
                .Col = enREVIEW1.tcDOCTNM:  .Value = Trim("" & tmpRs.GetValue("DoctNm"))     '처방의
                SvOrdDt = Trim("" & tmpRs.GetValue("OrdDate"))
                SvOrdNo = Trim("" & tmpRs.GetValue("OrdNo"))
                SvSpcNm = Trim("" & tmpRs.GetValue("SpcNm"))
                SvDoctNm = Trim("" & tmpRs.GetValue("DoctNm"))
            End If
            If SvOrdNo <> Trim("" & tmpRs.GetValue("OrdNo")) Then
                .Col = enREVIEW1.tcORDNO:   .Value = Trim("" & tmpRs.GetValue("OrdNo"))      '처방번호
                .Col = enREVIEW1.tcSPCNM:   .Value = Trim("" & tmpRs.GetValue("SpcNm"))      '검체명
                .Col = enREVIEW1.tcDOCTNM:  .Value = Trim("" & tmpRs.GetValue("DoctNm"))     '처방의
                SvOrdNo = Trim("" & tmpRs.GetValue("OrdNo"))
                SvSpcNm = Trim("" & tmpRs.GetValue("SpcNm"))
                SvDoctNm = Trim("" & tmpRs.GetValue("DoctNm"))
            End If
            If SvSpcNm <> Trim("" & tmpRs.GetValue("SpcNm")) Then
                .Col = enREVIEW1.tcSPCNM:   .Value = Trim("" & tmpRs.GetValue("SpcNm"))      '검체명
                SvSpcNm = Trim("" & tmpRs.GetValue("SpcNm"))
            End If
            If SvDoctNm <> Trim("" & tmpRs.GetValue("DoctNm")) Then
                .Col = enREVIEW1.tcDOCTNM:  .Value = Trim("" & tmpRs.GetValue("DoctNm"))     '처방의
                SvDoctNm = Trim("" & tmpRs.GetValue("DoctNm"))
            End If
            
            .Col = enREVIEW1.tcTESTNM:    .Value = Trim("" & tmpRs.GetValue("TestNm"))        '검사명
            .Col = enREVIEW1.tcSTATFG:    .Value = Choose(Val("" & tmpRs.GetValue("StatFg")) + 1, " ", "Y")     '응급여부
            .Col = enREVIEW1.tcRCVDT:     .Value = "" & Format(Format(tmpRs.GetValue("RcvDt"), CS_DateMask), "YY/MM/DD") & " " & _
                                                                      tmpRs.GetValue("RcvTm")                   '접수일시
            .Col = enREVIEW1.tcORDDATE:   .Value = "" & Format(Format(tmpRs.GetValue("OrdDt"), CS_DateMask), "YY/MM/DD") & " " & _
                                                                      tmpRs.GetValue("OrdTm")                   '처방일시
            .Col = enREVIEW1.tcORDDOCT:   .Value = Trim("" & tmpRs.GetValue("DoctNm"))        '처방의
            .Col = enREVIEW1.tcSPCNAME:   .Value = Trim("" & tmpRs.GetValue("SpcNm"))         '검체명
            .Col = enREVIEW1.tcORDNUM:    .Value = Trim("" & tmpRs.GetValue("OrdNo"))         '처방번호
            .Col = enREVIEW1.tcWORKAREA:  .Value = Trim("" & tmpRs.GetValue("WorkArea")): pWorkArea = .Value    'WorkArea
            .Col = enREVIEW1.tcACCDT:     .Value = Trim("" & tmpRs.GetValue("AccDt")):    pAccDt = .Value       'AccDt
            .Col = enREVIEW1.tcACCSEQ:    .Value = Trim("" & tmpRs.GetValue("AccSeq")):   pAccSeq = .Value      'AccSeq
            .Col = enREVIEW1.tcVFYNM:     .Value = Trim("" & tmpRs.GetValue("ExamNm"))
            .Col = enREVIEW1.tcVFYDATE:   .Value = Trim("" & tmpRs.GetValue("ExamDt")) & " " & _
                                                   Trim("" & tmpRs.GetValue("ExamTm"))                          '보고일시
            .Col = enREVIEW1.tcTESTCD:    .Value = Trim("" & tmpRs.GetValue("OrdCd"))                           '처방코드
            .Col = enREVIEW1.tcSPCCD:     .Value = Trim("" & tmpRs.GetValue("SpcCd"))                           '검체코드
            .Col = enREVIEW1.tcSPCYY:     .Value = Trim("" & tmpRs.GetValue("SpcYy"))                           '검체년도
            .Col = enREVIEW1.tcSPCNO:     .Value = Trim("" & tmpRs.GetValue("SpcNo"))                           '검체번호
            .Col = enREVIEW1.tcORDDIV:    .Value = Trim("" & tmpRs.GetValue("OrdDiv"))                          '처방구분
            .Col = enREVIEW1.tcUNITQTY:   .Value = Trim("" & tmpRs.GetValue("UnitQty")): strUnit = .Value       '수혈수량
                                            
            
            
            .Col = enREVIEW1.tcREQDATE:   .Value = Trim("" & tmpRs.GetValue("ReqDt"))         '수혈예정일
            .Col = enREVIEW1.tcREQTIME:   .Value = Trim("" & tmpRs.GetValue("ReqTm"))         '수혈예정시간
            .Col = enREVIEW1.tcWARDID:    .Value = Trim("" & tmpRs.GetValue("WardId"))        '병동
            .Col = enREVIEW1.tcHOSILID:   .Value = Trim("" & tmpRs.GetValue("HosilId"))       '호실
            .Col = 31:   .Value = Trim("" & tmpRs.GetValue("PanelFg"))       '그룹여부
            .Col = 32:   .Value = Trim("" & tmpRs.GetValue("TestDiv"))
         
            strOrdDiv = Trim("" & tmpRs.GetValue("OrdDiv"))
            strStsCd = Trim("" & tmpRs.GetValue("StsCd"))
            strTestDiv = Trim("" & tmpRs.GetValue("TestDiv"))
            
            .Col = enREVIEW1.tcSTSCD:     .Value = strStsCd     'Status
            
            .Col = enREVIEW1.tcSTSNM:
            If strOrdDiv = "B" And strStsCd = "3" Then
            '혈액은행인경우 strstscd가 3일때 검사중/완결 여부를 체크한다.(2001/07/09,kjg)
                .Value = BBS_STATUS(pWorkArea, pAccDt, pAccSeq, strUnit): .ForeColor = DCM_Gray
            
            Else
                   Call GetOrderStatus(strOrdDiv, strStsCd, strTestDiv, _
                                       strStsNm, lngColor, iBtnFg)
                   .Value = strStsNm: .ForeColor = lngColor
            End If
            
                   'D/C여부
                   If tmpRs.GetValue("DcFg") = "1" Then .Value = .Value & "*"
            
            .Col = enREVIEW1.tcTAT   '검사소요시간버튼
                   If iBtnFg = 1 Then
                       .CellType = CellTypeButton
                       .TypeButtonText = CS_QuestionMark   '"?"
                       .TypeButtonColor = DCM_LightGray     '회색
                   Else
                       .CellType = CellTypeStaticText
                       .Text = ""
                   End If
            
            '진료과 Remark(Message)
            aryMesg(RecordCnt) = "" & tmpRs.GetValue("Mesg")
         
        Wend
      
        If .MaxRows < lngMaxRows Then .MaxRows = lngMaxRows
        .RowHeight(-1) = lngRowHeight
        .Col = 1: .Row = 1: .Action = ActionActiveCell
      
       .ReDraw = True
    End With
   
Err_Trap:
    
    objStatus.Visible = False
    Set objStatus = Nothing
    
    ClearFg = False
    OrderFg = True
    OldRow = 0
   
    tmpRs.CloseCursor
    Set tmpRs = Nothing
   
    MouseDefault
    Me.Enabled = True
    QueryFg = False
    tblOrdSheet.SetFocus
   
    If RecordCnt = 0 Then
        MsgBox "이 환자는 입력하신 기간동안에 발생한 처방이 없습니다.", vbInformation, "결과조회"
        OrderFg = False
        Exit Sub
    End If
   
End Sub
Private Function BBS_STATUS(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, ByVal unitQty As String) As String
    If Val(unitQty) <= objSql.GetAssignCnt(WorkArea, AccDt, AccSeq) Then
        BBS_STATUS = "완결"
    Else
        BBS_STATUS = "검사중"
    End If
End Function
Private Sub GetOrderStatus(ByVal pOrdDiv As String, ByVal pStsCd As String, _
                           ByVal pTestDiv As String, ByRef pStsNm As String, _
                           ByRef pStsColor As Long, ByRef pBttnFg As Long)

    Select Case Trim(pStsCd)
        Case enStsCd.StsCd_LIS_Order:
             pStsNm = STS_LIS_Order:     pStsColor = DCM_Gray: pBttnFg = 1 '회색
        
        Case enStsCd.StsCd_LIS_Collection:
             pStsNm = STS_LIS_HaveSpc:   pStsColor = DCM_Gray: pBttnFg = 1 '회색
        
        Case enStsCd.StsCd_LIS_Accession:
             pStsNm = STS_LIS_Access:    pStsColor = DCM_Gray: pBttnFg = 1 '회색
        
        Case enStsCd.StsCd_LIS_InProcess:
             pStsNm = STS_LIS_Worksheet: pStsColor = DCM_Gray: pBttnFg = 1 '회색
        
        Case enStsCd.StsCd_LIS_MidRst:
             pBttnFg = 1
             If pOrdDiv = APS_ORDDIV Then
                 pStsNm = STS_LIS_Reading:   pStsColor = DCM_Gray           '회색
             Else
                 If pTestDiv = TST_MicTest Then            '미생물검사
                     pStsNm = STS_LIS_MidRst:    pStsColor = DCM_Black      '검정색
                 Else: pStsNm = STS_LIS_Partial: pStsColor = DCM_Black      '검정색
                 End If
             End If
        
        Case enStsCd.StsCd_LIS_FinRst:
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = IIf(pOrdDiv = APS_ORDDIV, STS_LIS_MidRst, _
                      IIf(pTestDiv = TST_MicTest, STS_LIS_FinRst, STS_LIS_Verify))
        
        Case enStsCd.StsCd_LIS_Modify:
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = IIf(pOrdDiv = APS_ORDDIV, STS_LIS_Verify, STS_LIS_Modify)
        
        Case "7":
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = STS_LIS_Modify
    End Select

End Sub

Private Sub Form_Activate()
    '
    MsgFg = False
    Call chkPtList_Click
On Error GoTo Err_Trap
    txtPtId.SetFocus
Err_Trap:
End Sub


Private Sub lblReset_Click()
    lvwPtList.ListItems.Clear
    txtSearchKey.Text = ""
End Sub


Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim tmpStr As String
    
    On Error GoTo Err_Trap
    
    If Item.Text = "" Then Exit Sub
    txtPtId.SetFocus
    DoEvents
    txtPtId.Text = Item.Text
    Call txtPtId_KeyPress(vbKeyReturn)
    Exit Sub
    
Err_Trap:
    Resume Next

End Sub


Private Sub optOrdDiv_Click(Index As Integer)
    optOrdDiv(0).ForeColor = &H404040
    optOrdDiv(1).ForeColor = &H404040
    optOrdDiv(2).ForeColor = &H404040
    optOrdDiv(Index).ForeColor = DCM_LightRed
    ChkDivAll.Value = 0
End Sub

Private Sub optQueryKey_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then dtpFromDate.SetFocus
End Sub


Private Sub rtfResult_DblClick()
    
    Dim sLabNo()
    Dim strTag As String
    Dim strLabNo As String
    Dim aryLabNo As Variant
    
    MouseRunning
    DoEvents
    
    strTag = rtfResult.Tag
    strLabNo = medGetP(strTag, 1, COL_DIV)
    aryLabNo = Split(strLabNo, "-")
    If aryLabNo(3) = BBS_ORDDIV Then Exit Sub
    
'    Set objText = frmAPS905
'    objText.fraTextResult.Visible = False
    frmAPS905.rtfResultText.Visible = True
    If aryLabNo(3) = APS_ORDDIV Then
        Call frmAPS905.GetResultText(aryLabNo(0), aryLabNo(1), aryLabNo(2))
    ElseIf aryLabNo(3) = LIS_ORDDIV Then
        frmAPS905.Caption = medGetP(strTag, 2, COL_DIV)
        frmAPS905.rtfResultText.TextRTF = rtfResult.TextRTF
    End If
    frmAPS905.Top = Me.Top
    frmAPS905.Left = Me.Left + 7000
    
    MouseDefault
    
    frmAPS905.Show vbModal
    DoEvents
    
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    Dim pWorkArea As String
    Dim pAccDt As String
    Dim pAccSeq As String
    Dim pTestCd As String
    Dim pSpcCd As String
    Dim pErFg As String
    Dim pTestNm As String
    Dim iNo As Integer
    
    Dim objResult As New clsLISResultReview
    Dim strTATS As String
   
    On Error GoTo Err_Trap:
   
    With tblOrdSheet
        .Row = Row
        .Col = enREVIEW1.tcWORKAREA: pWorkArea = .Value
        .Col = enREVIEW1.tcACCDT:    pAccDt = .Value
        .Col = enREVIEW1.tcACCSEQ:   pAccSeq = .Value
        .Col = enREVIEW1.tcTESTCD:   pTestCd = .Value
        .Col = enREVIEW1.tcSPCCD:    pSpcCd = .Value
        .Col = enREVIEW1.tcSTATFG:   pErFg = .Value
        .Col = enREVIEW1.tcTESTNM:   pTestNm = .Value
        
        '검사소요시간 읽어오기...
        
        strTATS = objResult.GetTAT(pTestCd, pSpcCd, pErFg)
        If pAccSeq = "" Then
            iNo = 1
        Else
            iNo = objResult.GetBuildNoForTAT(pWorkArea, pAccDt, pAccSeq)
        End If
        .Col = enREVIEW1.tcTAT
        .CellType = CellTypeEdit
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .Text = medGetP(strTATS, iNo, ":")
        MsgBox pTestNm & " 항목은 접수일시로부터 " & medGetP(strTATS, iNo, ":") & " 의 검사시간이 소요됩니다.", vbInformation, "검사소요시간"
    End With
    Exit Sub
    
Err_Trap:

End Sub

'% 처방 선택(Click)하면 해당 결과 디스플레이...
Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)

    Dim pWorkArea As String
    Dim pAccDt As String
    Dim pAccSeq As String
    Dim strOrdDiv As String
    Dim strWardId As String
    Dim strHosilId As String
    
    With tblOrdSheet
      
        If Row = 0 Then Exit Sub
        If Row > .DataRowCnt Then Exit Sub
        
        '소요시간
        If Col = enREVIEW1.tcTAT Then
            Call tblOrdSheet_ButtonClicked(Col, Row, 1)
            Exit Sub
        End If
      
        If OldRow = Row Then Exit Sub
        
        .Row = Row
        .Col = enREVIEW1.tcWORKAREA: pWorkArea = .Value
        .Col = enREVIEW1.tcACCDT:    pAccDt = .Value
        .Col = enREVIEW1.tcACCSEQ:   pAccSeq = .Value
        .Col = enREVIEW1.tcORDDIV:   strOrdDiv = .Value
        .Col = enREVIEW1.tcWARDID:   strWardId = .Value
        .Col = enREVIEW1.tcHOSILID:  strHosilId = .Value
        
        
        '병동 (처방난 시점)
        If strWardId <> "" Then
            lblLocation.Caption = strWardId & " - " & strHosilId
        Else
            lblLocation.Caption = ""
        End If
        
        If strOrdDiv = LIS_ORDDIV And (pWorkArea = "" Or pAccDt = "" Or pAccSeq = "") Then
            .Col = enREVIEW1.tcSTSCD
            If (.Value <> enStsCd.StsCd_LIS_Order) Then       '처방
                MsgBox "접수번호가 없습니다. (전산실로 연락바람 ☎" & ObjSysInfo.HelpLine & ")", vbExclamation, "오류발생"
            End If
            Exit Sub
        End If
      
        Call ResultClear
      
        If OldRow > 0 Then
            .Row = OldRow
            .Col = -1: .ForeColor = DCM_Black   '검정색
            
            .Col = enREVIEW1.tcSTSCD    '상태(처방,채혈,접수,검사중)
            If OldOrdDiv = LIS_ORDDIV And .Value = enStsCd.StsCd_LIS_Order Or .Value = enStsCd.StsCd_LIS_Collection Or _
               .Value = enStsCd.StsCd_LIS_Accession Or .Value = enStsCd.StsCd_LIS_InProcess Then
                .Col = enREVIEW1.tcSTSNM: .ForeColor = DCM_Gray            '회색
            End If
            .Col = enREVIEW1.tcSTSCD    '상태(처방,채혈,접수,검사중)
            If OldOrdDiv = APS_ORDDIV And .Value = enStsCd.StsCd_LIS_MidRst Then
                .Col = enREVIEW1.tcSTSNM: .ForeColor = DCM_Gray           '회색
            End If
        End If
         
        .Row = Row
        .Col = -1: .ForeColor = DCM_Blue        '파랑색
        OldRow = Row
        OldOrdDiv = strOrdDiv
      
        MouseRunning  '13
      
        tblResult.ReDraw = False
        
        .Col = enREVIEW1.tcSPCNAME: lblSpecimenNm.Caption = .Value      '검체
        .Col = enREVIEW1.tcORDDATE: lblOrdDt.Caption = Format(.Value, "YYYY-MM-DD HH:MM")  '처방일
        .Col = enREVIEW1.tcORDDOCT: lblDoctNm.Caption = .Value          '처방의
        .Col = enREVIEW1.tcVFYNM:   lblVerifierNm.Caption = .Value      '보고자
        .Col = enREVIEW1.tcVFYDATE: lblVerifyDt.Caption = .Value        '보고일시
      
        lblCollectorNm.Caption = ""
        lblReceiverNm.Caption = ""
      
        lblCollectDt.Caption = ""
        lblReceiveDt.Caption = ""
        
        Select Case strOrdDiv
        Case APS_ORDDIV, BBS_ORDDIV:
            rtfResult.Text = ""
            rtfResult.Tag = pWorkArea & "-" & pAccDt & "-" & pAccSeq & "-" & strOrdDiv
            If strOrdDiv = APS_ORDDIV Then
                #If AllowAPSResultReview Then
'                    picOrder.Width = 7300
'                    picResult.Width = Me.Width - picOrder.Width
'                    rtfResult.Width = picResult.Width
                    cmdRefresh.Left = picOrder.Width - cmdRefresh.Width - 50
                    DoEvents
                    Call DisplayAPSResult(pWorkArea, pAccDt, Val(pAccSeq))
                #End If
            Else
                #If AllowBBSResultReview Then
                    rtfResult.Visible = True
                    rtfResult.ZOrder 0
                    DoEvents
'                    picOrder.Width = 8000
                    cmdRefresh.Left = picOrder.Width - cmdRefresh.Width - 50
                    DoEvents
                    Call DisplayBBSResult(pWorkArea, pAccDt, Val(pAccSeq), Row)
                #End If
            End If
        Case LIS_ORDDIV:
            rtfResult.Tag = pWorkArea & "-" & pAccDt & "-" & pAccSeq & "-" & strOrdDiv
            rtfResult.Visible = False
            DoEvents
            tblResult.ReDraw = False
'            picOrder.Width = 8000
            cmdRefresh.Left = picOrder.Width - cmdRefresh.Width - 50
            DoEvents
            Call DisplayLISResult(pWorkArea, pAccDt, Val(pAccSeq))
            tblResult.ReDraw = True
        
        Case POC_ORDDIV:
            .Col = enREVIEW1.tcORDDATE:     pAccDt = Format(.Value, CS_DateDbFormat)
            Call DisplayPOCResult(txtPtId.Text, pAccDt)
        
        End Select
        
        tblResult.TopRow = 1
        ResultFg = True
      
        tblResult.ReDraw = True
        tblResult.Refresh
   
        chkPtList.Value = 0
        DoEvents
   
        MouseDefault
   
    End With
    
End Sub

'% Lab No.를 기준으로 검색한 결과내역을 테이블에 Display한다.
Private Sub DisplayBBSResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As Integer, ByVal iRow As Long)

    #If AllowBBSResultReview Then
        
        Dim strTransResult As String
        Dim strUnitQty     As String
        Dim strReqDtTm     As String
        Dim strReason      As String
        Dim strOrdDt       As String
        Dim strOrdNo       As String
        Dim lngAssignCnt   As Long
        Dim lngDeliveryCnt As Long
        
    
        Dim ObjABO         As New clsABO
        Dim objTransReason As New clsQueryOrder
        Dim rs             As DrRecordSet
        
        Dim strTmp         As String
        Dim strTmpBlood    As String
        Dim strJudge       As String
        Dim TF             As Boolean
        
        With tblOrdSheet
            .Row = iRow
            .Col = enREVIEW1.tcUNITQTY: strUnitQty = .Value
            .Col = enREVIEW1.tcREQDATE: strReqDtTm = Format(.Value, CS_DateMask)
            .Col = enREVIEW1.tcREQTIME: strReqDtTm = strReqDtTm & " " & Format(Mid(.Value, 1, 4), CS_TimeShortMask)
            .Col = enREVIEW1.tcORDDATE: strOrdDt = Format(.Value, CS_DateDbFormat)
            .Col = enREVIEW1.tcORDNUM:  strOrdNo = .Value
        
        End With
        
        strReason = objTransReason.GetTransReason(txtPtId.Text, strOrdDt, strOrdNo)
        Set objTransReason = Nothing
        
        strTmp = objSql.GetDeliveryCnt(pWorkArea, pAccDt, CStr(pAccSeq))
        
        If strTmp <> "" Then
            lngAssignCnt = Val(medGetP(strTmp, 1, COL_DIV)) - Val(medGetP(strTmp, 2, COL_DIV))
            lngDeliveryCnt = Val(medGetP(strTmp, 3, COL_DIV))
        End If
    
    '    Set objABOinit.Database = DbConn
    
        ObjABO.PtId = txtPtId.Text  '혈액형을 구하자.
        ObjABO.GetABO
        
        Set rs = objTransReason.DonorInformation(txtPtId.Text)
        
        
        
        With rtfResult
            .Visible = False
            .Text = vbCrLf & Space(13) & "◈ 수혈 진행상황 ◈" & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "▶ 혈 액 형 : " & ObjABO.ABO & ObjABO.Rh & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "▶ 예 정 일 : " & strReqDtTm & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "▶ 수혈사유 : " & strReason & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "▶ 수    량 : " & strUnitQty & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "▶ 결과등록 : " & lngAssignCnt & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "▶ 출고수량 : " & lngDeliveryCnt & vbCrLf & vbCrLf
            If Not rs.EOF Then
                Do Until rs.EOF
                    Select Case rs.Fields("okdiv3").Value & ""
                        Case "1":  strJudge = "적격"
                        Case "0":  strJudge = "부적격"
                        Case Else: strJudge = "미등록"
                    End Select
                    
                    strTmpBlood = rs.Fields("donornm").Value & "" & "(" & rs.Fields("tmpid").Value & "" & "," & strJudge & ")" & vbCrLf
                    If TF = False Then
                        .Text = .Text & Space(3) & "▶ 헌 혈 자 : " & strTmpBlood & vbCrLf '& vbCrLf
                    Else
                        .Text = .Text & Space(3) & Space(13) & strTmpBlood & vbCrLf
                    End If
                    TF = True
                    rs.MoveNext
                Loop
            End If
            
            .SelStart = 15: .SelLength = Len(.Text)
            .SelFontName = "굴림체"
            .SelFontSize = 13
            .SelBold = True
            
            .SelStart = 30: .SelLength = Len(.Text)
            .SelFontName = "돋움체"
            .SelFontSize = 10
            .SelBold = True
            '.SelColor = &H553755 &HE48372 '약간 파란색
            
            Call HighlightText(rtfResult, "◈ 수혈 진행상황 ◈", True, , &H4A4189)
            Call HighlightText(rtfResult, "▶ 혈 액 형 :", False, , &H553755)
            Call HighlightText(rtfResult, ObjABO.ABO & ObjABO.Rh, False, , &H7477EF, 15)  '약간 붉은색
            Call HighlightText(rtfResult, "▶ 예 정 일 :", False, , &H553755)
            Call HighlightText(rtfResult, strReqDtTm, False, , &HE48372)
            Call HighlightText(rtfResult, "▶ 수혈사유 :", False, , &H553755)
            Call HighlightText(rtfResult, "▶ 수    량 :", False, , &H553755)
            Call HighlightText(rtfResult, "▶ 결과등록 :", False, , &H553755)
            Call HighlightText(rtfResult, "▶ 출고수량 :", False, , &H553755)
            If TF = True Then
                Call HighlightText(rtfResult, "▶ 헌 혈 자 :", False, , &H553755)
            End If
            .Visible = True
        
        End With
        
        Set rs = Nothing
        Set ObjABO = Nothing
    
    #End If
    
End Sub


Private Sub DisplayPOCResult(ByVal pPtId As String, ByVal pVfyDt As String)

    Dim i As Integer, j As Integer
    Dim objResult As New clsLISResultReview
    Dim ResultBuffer As String
    
    
    With objResult
        ResultBuffer = .POCResultQuery(pPtId, pVfyDt)
        For i = 1 To .RstRow
            tblResult.Row = i   '+ .OffSet
            For j = 1 To 8
                tblResult.Col = j
                tblResult.ForeColor = .Get_ForeColor(j, i)
            Next
        Next
    End With
    
    '결과내역 Display
    tblResult.Row = 1
    tblResult.Row2 = tblResult.MaxRows
    tblResult.Col = 2
    tblResult.Col2 = tblResult.MaxCols
    tblResult.BlockMode = True
    tblResult.AllowCellOverflow = True
    tblResult.Clip = ResultBuffer
    tblResult.BlockMode = False
    
End Sub


'% Lab No.를 기준으로 검색한 결과내역을 테이블에 Display한다.
Private Sub DisplayLISResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
   
    Dim i As Integer, j As Integer
    Dim objResult As New clsLISResultReview
    Dim ResultBuffer As String
    Dim RstTxtBuffer As String
    Dim SamTxtBuffer As String
    Dim strTestDiv As String
   
    With objResult
      
        'Set .MyDb = DBConn
        
        Call .ResultQuery(pWorkArea, pAccDt, pAccSeq)
        
        lblWorkArea.Caption = .GetWorkAreaNm(pWorkArea)     'WorkAra Name
        lblCollectorNm.Caption = .GetDoctNm(.ColId)         '채혈자(간호사)
        If Trim(lblCollectorNm.Caption) = "" Then
            lblCollectorNm.Caption = .GetEmpNm(.ColId)      '채혈자(병리사)
        End If
        lblReceiverNm.Caption = .GetEmpNm(.RcvId)           '접수자
        
        lblCollectDt.Caption = .ColDtTm                     '채혈일시
        lblReceiveDt.Caption = .RcvDtTm                     '접수일시
        
        lblDeptNm.Caption = .DeptNm
        lblBedinDt.Caption = .BedIndt
        
        '상태가 처방/채혈/접수/검사중이면 Exit
        tblOrdSheet.Col = enREVIEW1.tcSTSCD
        If tblOrdSheet.Value = enStsCd.StsCd_LIS_Order Or tblOrdSheet.Value = enStsCd.StsCd_LIS_Collection Or _
            tblOrdSheet.Value = enStsCd.StsCd_LIS_Accession Or tblOrdSheet.Value = enStsCd.StsCd_LIS_InProcess Then Exit Sub
        If .ResultCnt = 0 Then Exit Sub
      
        ' 일반검사 - High / Low 컬럼 ForeColor 설정
        For i = 1 To .RstRow
            tblResult.Row = i   '+ .OffSet
            For j = 1 To 8
                tblResult.Col = j
                'If .Get_ForeColor(j, i) <> 0 Then tblResult.ForeColor = .Get_ForeColor(j, i)
                tblResult.ForeColor = .Get_ForeColor(j, i)
            Next
        Next
        'End If
      
        '결과내역 Display
        tblResult.Row = 1
        tblResult.Row2 = tblResult.MaxRows
        tblResult.Col = 2
        tblResult.Col2 = tblResult.MaxCols
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.Clip = .ResultClipText '& .SenClipText 'ResultBuffer
        tblResult.BlockMode = False
      
        '미생물 감수성 결과의 경우 항생제명 순으로 Sort / Align Left
        If .SortFg Then
            For i = 1 To .SensiCount
                tblResult.SortBy = SortByRow
                tblResult.SortKey(1) = 2  '항생제명
                tblResult.SortKeyOrder(1) = SortKeyOrderAscending
                tblResult.Col = -1
                tblResult.Row = .AntiSortStartRow(i)   '+ .OffSet
                tblResult.Row2 = .AntiSortEndRow(i)    '+ .OffSet
                tblResult.Action = ActionSort
                tblResult.Row = .SortStartRow - 1 '+ .OffSet
                tblResult.Col = 2
                tblResult.FontUnderline = True
            Next
        Else
            tblResult.Col = 6
            tblResult.Row = -1
            tblResult.ForeColor = DCM_LightRed
            tblResult.FontBold = True
        End If
        If .TestDiv = TST_MicTest Then
            '미생물 결과 : 균명컬럼 Align Left
            tblResult.Row = -1
            tblResult.Col = -1
            tblResult.BlockMode = True
            tblResult.AllowCellOverflow = True
            tblResult.TypeHAlign = TypeHAlignLeft
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 17
            'tblResult.ColWidth(3) = 60
            For i = 1 To 5
                If .MicFg(i) Then
                    tblResult.ColWidth(i + 2) = 9
                Else
                    tblResult.ColWidth(i + 2) = 4
                End If
            Next
            tblResult.ColWidth(8) = 20
            tblResult.Col = 3: tblResult.Col2 = 7
            tblResult.Row = -1
            tblResult.BlockMode = True
            tblResult.FontBold = False
            tblResult.BlockMode = False
        Else
            '일반결과 : 결과컬럼 Align Center
            tblResult.Row = 1: tblResult.Row2 = tblResult.MaxRows
            tblResult.Col = 3: tblResult.Col2 = 7
            tblResult.BlockMode = True
            tblResult.TypeHAlign = TypeHAlignCenter
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 13
            tblResult.ColWidth(3) = 9
            tblResult.ColWidth(4) = 9
            tblResult.ColWidth(5) = 3
            tblResult.ColWidth(6) = 5
            tblResult.ColWidth(7) = 13
        End If
      
        '텍스트결과 Display
        If .TextFg Then
            txtRstCmt.TextRTF = .RstTextBuffer      'RstTxtBuffer
            chkRstCmt.Value = 1
            chkRstCmt.Enabled = True
            Call HighlightText(txtRstCmt, "<< 검사 소견 >>", True, , &H4A4189)
            Call HighlightText(txtRstCmt, "<< Supplemental Report >>", False, , &H4A4189)
        Else
            chkRstCmt.Value = 0
            chkRstCmt.Enabled = False
        End If
       
        '특수검사 결과 Display
        If .SpecialFg Then
            rtfResult.TextRTF = .SpeTextBuffer
            rtfResult.Tag = rtfResult.Tag & COL_DIV & .SpeRstTitle
            tblOrdSheet.Row = OldRow
            tblOrdSheet.Col = 32: strTestDiv = tblOrdSheet.Value
            If strTestDiv = CStr(enTestDiv.TST_SpeTest) Then Call rtfResult_DblClick
        End If
        
        
        '검체리마크 & 풋노트 Display
        If .CommentFg Then
            txtSamCmt.Text = .SamTextBuffer
'            txtSamCmt1.Text = .SamTextBuffer
            chkSamCmt.Value = 1
            chkSamCmt.Enabled = True
            Call HighlightText(txtSamCmt, "<< Remark >>", True)
            Call HighlightText(txtSamCmt, "<< Foot Note >>", False)
'            Call HighlightText(txtSamCmt1, "<< Remark >>", True)
'            Call HighlightText(txtSamCmt1, "<< Foot Note >>", False)
        Else
            chkSamCmt.Value = 0
            chkSamCmt.Enabled = False
        End If
        
    End With
    
    Set objResult = Nothing
   
End Sub

'% Lab No.를 기준으로 검색한 결과내역을 테이블에 Display한다.
Private Sub DisplayAPSResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)

#If AllowAPSResultReview Then

    Dim i As Integer, j As Integer
    Dim ResultBuffer As String
    Dim RstTxtBuffer As String
    Dim SamTxtBuffer As String
    Dim strWAccDt As String
    Dim strAccSeq As String
    Dim rs As New DrRecordSet
    Dim objResult As New clsAPSResult
    Dim strRsEntryType  As String
    
    With objResult
    
        strWAccDt = Trim(pWorkArea) & Trim(Mid(pAccDt, 3, 2))
        strAccSeq = Trim(Format(pAccSeq, "00000#"))
        
        Call .LoadResult(strWAccDt, strAccSeq, , False, False)
        
        strRsEntryType = .RstEntryType
        
        If strRsEntryType = "" Then Exit Sub
        
        If .stscd < "6" Then Exit Sub   ' 판독
        
'        Call .LoadResult(strWAccDt, strAccSeq, strRsEntryType)
'
'        ObjLISComCode.PTHDOCT.Exists (.PTHDOCT)
'        If ObjLISComCode.PTHDOCT.Exists(.PTHDOCT) = True Then
'            ObjLISComCode.PTHDOCT.KeyChange .PTHDOCT
'            lblVerifierNm.Caption = ObjLISComCode.PTHDOCT.Fields("pthdoctnm")   '확인자
'        Else
'            lblVerifierNm.Caption = ""
'        End If
'
'        lblDeptNm.Caption = .DeptCdNm
'
'        '결과 조회
'        Call LoadResultText(.WorkArea, .AccDt, .AccSeq)
'        rtfResult.Visible = True
'        DoEvents
        Call rtfResult_DblClick
        DoEvents

    End With

#End If
    
End Sub

Private Sub LoadResultText(ByVal pWorkArea As String, ByVal pAccDt As String, _
                           ByVal pAccSeq As String)

    #If AllowAPSResultReview Then
        Dim objText As clsAPSScreenResult
    
        Set objText = New clsAPSScreenResult
        'objText.setDbConn DBConn
        
        Call objText.LoadScreenResult(pWorkArea, pAccDt, pAccSeq, rtfResult)
    
        Set objText = Nothing
    #End If
End Sub


'% 폼 로드
Private Sub Form_Load()
    'Me.Show
    txtSearchKey.Text = ""
    chkPtList.Value = 0
    picPtList.Visible = False
    OrderFg = False
    ResultFg = False
    ClearFg = True
    PtFg = False
    optSort(1).Value = True
    OldRow = 0
    medInitLvwHead lvwPtList, "환자ID,환자성명,주민등록번호,생년월일,성별/나이", _
                       "50,50,800,300,100"
   
    If gUsingInWardMenu Then
        dtpFromDate.Value = DateAdd("d", -7, Now)
        optQueryKey(2).Value = True
    Else
        dtpFromDate.Value = DateAdd("d", -4, Now)
        optQueryKey(1).Value = True
    End If
    dtpToDate.Value = Now
'    picOrder.Width = 8000
'    picResult.Width = Me.Width - picOrder.Width
'    rtfResult.Width = picResult.Width
    Call ClearRtn
    ChkDivAll.Value = 1
    'Set objPatient.MyOraSE = OraSe
    Set objPatient.objDb = DBConn
    
    If gUsingInWardMenu Then
        ChkDivAll.Value = 1
        Call ChkDivAll_Click
    Else
        ChkDivAll.Value = 0
        Call ChkDivAll_Click
        Select Case ObjSysInfo.ProjectId
        Case "LIS": optOrdDiv(0).Value = True
        Case "APS": optOrdDiv(1).Value = True
        Case "BBS": optOrdDiv(2).Value = True
        End Select
    End If

#If AllowAPSResultReview Then
    If ObjLISComCode.PTHDOCT.RecordCount = 0 Then Call ObjLISComCode.LoadPthDoct
#End If

End Sub


'% 정렬 기준 선택
Private Sub optSort_Click(Index As Integer)
    If Not picPtList.Visible Then Exit Sub
    If txtSearchKey.Text <> "" Then
        Call txtSearchKey_KeyPress(vbKeyReturn)
    End If
    txtSearchKey.SetFocus
End Sub




'% 처방테이블 Set Focus
Private Sub tblOrdSheet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Trap
    If OrderFg Then tblOrdSheet.SetFocus
Err_Trap:
End Sub

'처방내역 테이블에 ToolTip 보여주기...
Private Sub tblOrdSheet_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    Dim tmpToolTip As String
    Dim tmpPanelFg As String
    Dim strSql As String
    
    Dim strWorkArea As String
    Dim strAccDt As String
    Dim strAccSeq As String
    Dim strReqdt  As String
    
    If Not OrderFg Then Exit Sub
   
    If Row <= 0 Then Exit Sub
    tmpToolTip = vbCrLf
   
    With tblOrdSheet
        .Row = Row
        
        .Col = 3: If Trim(.Value) = "" Then Exit Sub
        
        .Col = enREVIEW1.tcREQDATE:   '.Value = Trim("" & tmpRs.GetValue("ReqDt"))         '수혈예정일
        strReqdt = Format(.Value, "####-##-##")
        
        .Col = enREVIEW1.tcREQTIME:  ' .Value = Trim("" & tmpRs.GetValue("ReqTm"))         '수혈예정시간
        strReqdt = strReqdt & "  " & Format(.Value, "0#:##:##")
        
        .Col = 9:   tmpToolTip = tmpToolTip & "  처방일시 : " & .Value & vbCrLf  '처방일시
        .Col = 13:  tmpToolTip = tmpToolTip & "  처방번호 : " & .Value & vbCrLf   '처방번호
        .Col = 3:   tmpToolTip = tmpToolTip & "  검 사 명 : " & .Value & vbCrLf  '검사명
        .Col = 4:   tmpToolTip = tmpToolTip & "  검    체 : " & .Value & vbCrLf  '검체
        .Col = 11:  tmpToolTip = tmpToolTip & "  처 방 의 : " & .Value & vbCrLf '처방의
                    
        
        
        
        .Col = 14:  strWorkArea = .Value
        .Col = 15:  strAccDt = .Value
        .Col = 16:  strAccSeq = .Value
        
        .Col = 31:  tmpPanelFg = .Value
        If tmpPanelFg = PN_Group Then
            
            Dim objRs As DrRecordSet
            Dim lngSeq As String
            
            lngSeq = 0
            strSql = objSql.SqlMultiTest(strWorkArea, strAccDt, Val(strAccSeq))
            Set objRs = OpenRecordSet(strSql)
            If Not objRs.EOF Then
                tmpToolTip = tmpToolTip & "  접수번호 : " & vbCrLf
                While Not objRs.EOF
                    lngSeq = lngSeq + 1
                    tmpToolTip = tmpToolTip & "      복수검체 " & CStr(lngSeq) & " : " & objRs.Fields("WorkArea").Value & "-"
                    tmpToolTip = tmpToolTip & Mid("" & objRs.Fields("AccDt").Value, 3) & "-"
                    tmpToolTip = tmpToolTip & objRs.Fields("AccSeq").Value & vbCrLf
                    objRs.MoveNext
                Wend
                objRs.RsClose
                Set objRs = Nothing
            Else
                tmpToolTip = tmpToolTip & "  접수번호 : " & strWorkArea & "-"
                tmpToolTip = tmpToolTip & Mid(strAccDt, 3) & "-"
                tmpToolTip = tmpToolTip & strAccSeq & vbCrLf
            End If
        Else
            tmpToolTip = tmpToolTip & "  접수번호 : " & strWorkArea & "-"
            tmpToolTip = tmpToolTip & Mid(strAccDt, 3) & "-"
            tmpToolTip = tmpToolTip & strAccSeq & vbCrLf
        End If
        If UBound(aryMesg) >= Row Then
            If aryMesg(Row) <> "" Then tmpToolTip = tmpToolTip & vbCrLf & "  " & aryMesg(Row) & vbCrLf
        End If
        tmpToolTip = tmpToolTip & "  희망채혈일시 :" & strReqdt & vbCrLf
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 4000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
   
End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    tblResult.Row = Row
    tblResult.Col = Col
    If tblResult.Value = "☞ RESULT" Then
        If Trim(rtfResult.Text) <> "" Then
            Call rtfResult_DblClick
        End If
    End If
End Sub

'% 결과테이블 Set Focus
Private Sub tblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Trap
    If ResultFg Then tblResult.SetFocus
Err_Trap:
End Sub

'결과내역 테이블에 ToolTip 보여주기...
Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    Dim tmpToolTip As String
    Dim svWorkArea As String
    Dim svAccDt As String
    Dim svAccSeq As String
    Dim strSql As String
    Dim rsMod As DrRecordSet
    
    If Not ResultFg Then Exit Sub
    
    tmpToolTip = vbCrLf
   
    With tblResult
        .Row = Row
        .Col = 2:
                 If .Value = "" Then
                    ShowTip = False
                    Exit Sub
                 End If
        .Col = 8:  tmpToolTip = tmpToolTip & "  " & .Value & vbCrLf   '처방명(Long)
        .Col = 9:
                If .Value <> "" Then
                    tmpToolTip = tmpToolTip & vbCrLf & "  최근결과 : " & .Value & vbCrLf   '최근결과
                    .Col = 10
                    tmpToolTip = tmpToolTip & "  보고일시 : " & .Value & vbCrLf  '최근결과일
                End If
        .Col = 12: 'If .Value <> "1" Then GoTo Skip
                 
                Dim strModRst As String
                 
        .Col = 13:
                strSql = objSql.SqlGetOldResult(svWorkArea, svAccDt, svAccSeq, .Value)
                Set rsMod = OpenRecordSet(strSql)
                If Not rsMod.EOF Then
                   tmpToolTip = tmpToolTip & vbCrLf & "  [ 수정전 결과 ]  " '& vbCrLf
                
                   While (Not rsMod.EOF)
                      'strModRst = "  " & rsMod.Fields("AbbrNm5").Value & Space(5)
                      'strModRst = Mid(strModRst, 1, 16) & ":  " & rsMod.Fields("RstCd").Value & Space(10)
                      strModRst = Trim("" & rsMod.Fields("RstCd").Value) & Space(3)
                      strModRst = strModRst & Format("" & rsMod.Fields("vfydt").Value, "####-##-##") & Space(2)
                      strModRst = strModRst & "by " & rsMod.Fields("EmpNm").Value & vbCrLf
                      tmpToolTip = tmpToolTip & strModRst & Space(19)
                      rsMod.MoveNext
                   Wend
                End If
                rsMod.RsClose
                Set rsMod = Nothing
      
      
Skip:
        MultiLine = 1
        If Trim(Replace(tmpToolTip, vbCrLf, "", 1, -1, vbBinaryCompare)) = "" Then
            ShowTip = False
            Exit Sub
        End If
        TipText = tmpToolTip
        TipWidth = 5500
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
   
End Sub

'% 환자ID가 변경되면 화면Clear
Private Sub txtPtId_Change()
    If Not ClearFg Then
        Call ClearRtn
    End If
End Sub

'% 환자 ID
Private Sub txtPtId_GotFocus()
    With txtPtId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% 환자정보 검색
Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtPtId_LostFocus()
    Dim strWardId As String
    
    If Not gUsingInWardMenu Then

        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        If Screen.ActiveControl Is Nothing Then Exit Sub
        
        If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
        If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
        If Screen.ActiveControl.Name = chkPtList.Name Then Exit Sub
        If Screen.ActiveControl.Name = chkVerified.Name Then Exit Sub
        If Screen.ActiveControl.Name = txtSearchKey.Name Then Exit Sub
    
    End If
    
    If MsgFg Then Exit Sub
      
    On Error GoTo Err_Trap
    If txtPtId.Text = "" Then
        'txtPtId.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    With objPatient
        If Trim(txtPtId.Text) <> "" And .PtntQuery(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblSex.Caption = .SexNm
            lblAge.Caption = .Age
            lblAgeDiv.Caption = .AgeDiv
            lblDeptNm.Caption = .DeptNm
            strWardId = .WardId
            If strWardId <> "" Then
                If .RoomId <> "" Then strWardId = strWardId & "-" & .RoomId
                If gUsingInWardMenu Then
                    dtpFromDate.Value = DateAdd("d", -2, Now)
                    'optQueryKey(2).Value = True
                End If
            Else
                If gUsingInWardMenu Then
                    dtpFromDate.Value = DateAdd("d", -10, Now)
                    'optQueryKey(2).Value = True
                End If
            End If
            lblLocation.Caption = strWardId
            lblBedinDt.Caption = Format(.BedIndt, CS_DateMask)
            lblBedoutDt.Caption = Format(.BedOutDt, CS_DateMask)
            PtFg = True
            ClearFg = False
        Else
            If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
            txtPtId.SetFocus
            MsgFg = False
            PtFg = False
            Call txtPtId_GotFocus
            Exit Sub
        End If
    End With

On Error GoTo Err_Trap
    If ActiveControl.Name <> cmdRefresh.Name Then dtpFromDate.SetFocus
    Exit Sub
Err_Trap:
    Resume Next
End Sub

'% 텍스트결과 박스 더블클릭 - Larger Box Popup
Private Sub txtRstCmt_DblClick()
'    Set objText = frmAPS905
    With frmAPS905
        .rtfResultText.Visible = True
        .rtfResultText.TextRTF = txtRstCmt.Text
        Call HighlightText(.rtfResultText, "<< 검사 소견 >>", True, , &H4A4189)
        Call HighlightText(.rtfResultText, "<< Supplemental Report >>", False, , &H4A4189)
        .Show vbModal
    End With
End Sub


'% 풋노트 박스 더블클릭 - Larger Box Popup
Private Sub txtSamCmt_DblClick()
'    Set objText = frmAPS905
    With frmAPS905
        .rtfResultText.Visible = False
'        .fraTextResult.Visible = True
        .rtfResultText.Text = txtRstCmt.Text & vbCrLf & vbCrLf
        .rtfResultText.Text = txtSamCmt.Text
'        .fraTextResult.ZOrder 0
        Call HighlightText(.rtfResultText, "<< 검사 소견 >>", True, , &H4A4189)
        Call HighlightText(.rtfResultText, "<< Supplemental Report >>", False, , &H4A4189)
        .Show vbModal
    End With
End Sub


'% 환자 검색 (ID 또는 성명으로...)
Private Sub txtSearchKey_GotFocus()

    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'% 환자ID 또는 성명으로 검색 리스트 작성.
Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
    
    Dim objPtInfo As New clsHosComSQLStmt
    Dim DrRs As New DrRecordSet
    Dim itmx As ListItem
    Dim lngSearch As Long
    Dim ColCnt As Long
    Dim RowCnt As Long
    
    'Set objPtInfo.objDb = dbConn
    If KeyAscii = vbKeyReturn Then
        lngSearch = IIf(optSort(0).Value, 1, 2)  'True:환자ID, False:환자명
        If lngSearch = 1 And Not IsNumeric(txtSearchKey.Text) Then Exit Sub
        If chkVerified.Value = 0 Then
            If lngSearch = 2 And Len(txtSearchKey.Text) < 2 Then
                MsgBox "2문자 이상 입력하신후 검색하십시오.", vbInformation, "환자검색"
                txtSearchKey.SetFocus
                Exit Sub
            End If
            ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey))
        Else
            ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey, _
                          mvarDeptCd, Format(DBConn.GetSysDate, CS_DateDbFormat)))
        End If
        lvwPtList.ListItems.Clear
        If ColCnt > 0 Then
            RowCnt = 0
            With lvwPtList
                Do While (DrRs.FetchCursor(ColCnt))
                    RowCnt = RowCnt + 1
                    Set itmx = .ListItems.Add(, , "" & DrRs.GetValue("ptid"))
                    itmx.SubItems(1) = "" & DrRs.GetValue("ptnm")
                    itmx.SubItems(2) = "" & DrRs.GetValue("SSN")
                    itmx.SubItems(3) = "" & DrRs.GetValue("DOB")
                    If Not IsDate(itmx.SubItems(3)) Then
                        itmx.SubItems(3) = Mid(itmx.SubItems(3), 1, 4) & "-01-01"
                    End If
                    If IsNumeric(Mid("" & DrRs.GetValue("ssn"), 8, 1)) Then
                        itmx.SubItems(4) = IIf((Mid("" & DrRs.GetValue("ssn"), 8, 1) Mod 2) = 1, "남", "여")
                    Else
                        itmx.SubItems(4) = "모름"
                    End If
                    If IsDate(itmx.SubItems(3)) Then
                        itmx.SubItems(4) = itmx.SubItems(4) & " / " & DateDiff("yyyy", itmx.SubItems(3), Now)
                    Else
                        itmx.SubItems(4) = itmx.SubItems(4) & " / ? "
                    End If
                    If RowCnt > 1000 Then Exit Do
                Loop
            End With
        Else
            MsgBox "조건에 맞는 자료가 없습니다. 확인후 검색하세요", vbInformation + vbOKOnly, Me.Caption
        End If
        DrRs.CloseCursor:     Set DrRs = Nothing
    
    End If
    
    Set objPtInfo = Nothing
    
End Sub



'% Clear 루틴
Private Sub ClearRtn()
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDeptNm.Caption = ""
    lblLocation.Caption = ""
    lblBedinDt.Caption = ""
    lblBedoutDt.Caption = ""
    rtfResult.Visible = False
    Call FieldClear
    Call TableClear
    ClearFg = True
    OrderFg = False
    MsgFg = False
    QueryFg = False
    OldRow = 0
End Sub

Private Sub FieldClear()

    lblDoctNm.Caption = ""
    lblCollectorNm.Caption = ""
    lblReceiverNm.Caption = ""
    lblVerifierNm.Caption = ""
    lblOrdDt.Caption = ""
    lblCollectDt.Caption = ""
    lblReceiveDt.Caption = ""
    lblVerifyDt.Caption = ""
    txtSamCmt.Text = ""
    txtRstCmt.Text = ""
'    txtSamCmt1.Text = ""
'    txtRstCmt1.Text = ""
    lblWorkArea.Caption = ""
    lblSpecimenNm.Caption = ""

End Sub

Private Sub TableClear()
    tblOrdSheet.MaxRows = 0
    tblOrdSheet.MaxRows = 100
    
    tblResult.MaxRows = 0
    tblResult.MaxRows = 100
End Sub

'% 결과 Part Clear
Private Sub ResultClear()
   
'    ResultBuffer = ""
'    RstTxtBuffer = ""
'    SamTxtBuffer = ""
    
    txtRstCmt.Text = ""
    txtSamCmt.Text = ""
'    chkRstCmt.Enabled = False
'    chkSamCmt.Enabled = False
'    txtRstCmt1.Text = ""
'    txtSamCmt1.Text = ""
       
    lblWorkArea.Caption = ""
    lblSpecimenNm.Caption = ""
   
    rtfResult.Tag = ""
    rtfResult.Text = ""
    ResultFg = False
   
    With tblResult
        '결과테이블 Clear
        .Row = -1:  .Col = -1
        .BlockMode = True
        .FontBold = False
        .Action = ActionClearText
        .BlockMode = False
        '검사명/결과 컬럼 Bold
        .Row = -1: .Col = 2: .Col2 = 3
        .BlockMode = True
        .FontBold = True
        .BlockMode = False
        'High/Low field font 지정
        .Row = -1: .Col = 5: .Col2 = 5
        .BlockMode = True
        .FontName = "돋움"
        .BlockMode = False
        .RowsFrozen = 0
    End With
   

End Sub

Public Sub Call_ToDate_LostFocus()

    If Not gUsingInWardMenu Then
    
        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        If ActiveControl.Name = cmdExit.Name Then Exit Sub
        If ActiveControl.Name = cmdClear.Name Then Exit Sub
        If ActiveControl.Name = chkPtList.Name Then Exit Sub
        
    End If
    
    Call dtpToDate_KeyDown(vbKeyReturn, 0)
   
End Sub


Public Sub Call_PtId_KeyPress()

   Call txtPtId_LostFocus

End Sub

