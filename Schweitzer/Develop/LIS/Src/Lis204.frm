VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm204WSDataEntry 
   BackColor       =   &H00DBE6E6&
   Caption         =   "워크쉬트별 결과등록"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14535
   Icon            =   "Lis204.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14535
   WindowState     =   2  '최대화
   Begin VB.TextBox txtBatchRst 
      Appearance      =   0  '평면
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11235
      MaxLength       =   15
      TabIndex        =   56
      Tag             =   "opt"
      Top             =   6285
      Width           =   1785
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00ACCDD0&
      Caption         =   "적용"
      Enabled         =   0   'False
      Height          =   330
      Left            =   13035
      Style           =   1  '그래픽
      TabIndex        =   55
      Top             =   6285
      Width           =   690
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "취소"
      Height          =   330
      Left            =   13740
      Style           =   1  '그래픽
      TabIndex        =   54
      Top             =   6285
      Width           =   690
   End
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H00DBE6E6&
      Caption         =   "특  수"
      Height          =   285
      Left            =   12735
      Style           =   1  '그래픽
      TabIndex        =   49
      Top             =   500
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdMicro 
      BackColor       =   &H00DBE6E6&
      Caption         =   "미생물"
      Height          =   285
      Left            =   13575
      Style           =   1  '그래픽
      TabIndex        =   48
      Top             =   500
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdRmk 
      BackColor       =   &H008080FF&
      Caption         =   "처방비고"
      Height          =   285
      Left            =   11835
      Style           =   1  '그래픽
      TabIndex        =   47
      Top             =   500
      Visible         =   0   'False
      Width           =   900
   End
   Begin MedControls1.LisLabel lblDisease 
      Height          =   270
      Left            =   8475
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   810
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   476
      BackColor       =   16777215
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "확인(&S)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   35
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   34
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   33
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.ComboBox cboRelTest 
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8475
      Style           =   2  '드롭다운 목록
      TabIndex        =   27
      Top             =   1095
      Width           =   5955
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFF7FC&
      Caption         =   "(&N) >>"
      Height          =   450
      Left            =   8505
      Style           =   1  '그래픽
      TabIndex        =   26
      Top             =   45
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFF7FC&
      Caption         =   "<< (&P)"
      Height          =   450
      Left            =   7170
      Style           =   1  '그래픽
      TabIndex        =   25
      Top             =   45
      Width           =   1320
   End
   Begin VB.PictureBox picRst 
      Height          =   4200
      Left            =   3060
      ScaleHeight     =   4140
      ScaleWidth      =   11340
      TabIndex        =   16
      Top             =   1980
      Width           =   11400
      Begin MSComctlLib.ProgressBar prgRst 
         Height          =   240
         Left            =   0
         TabIndex        =   21
         ToolTipText     =   "자료를 가져오고 있읍니다."
         Top             =   3900
         Visible         =   0   'False
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread ssRst 
         CausesValidation=   0   'False
         Height          =   3900
         Left            =   0
         TabIndex        =   8
         Tag             =   "20001"
         Top             =   0
         Width           =   11325
         _Version        =   196608
         _ExtentX        =   19976
         _ExtentY        =   6879
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   15857140
         GridColor       =   13290186
         MaxCols         =   19
         MaxRows         =   14
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis204.frx":038A
         VisibleCols     =   10
         VisibleRows     =   13
         TextTip         =   2
      End
      Begin VB.Label lblSpreadLoading 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "잠시 기다려 주세요. 결과 데이터를 로딩하고 있읍니다."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2970
         TabIndex        =   20
         Top             =   1890
         Width           =   4605
      End
   End
   Begin VB.Frame fraText 
      BackColor       =   &H00DBE6E6&
      Caption         =   " Text Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   8835
      TabIndex        =   18
      Tag             =   "20002"
      Top             =   6690
      Width           =   5625
      Begin VB.CommandButton cmdTextTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
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
         Left            =   5230
         Picture         =   "Lis204.frx":0C01
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   1455
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   1500
         Left            =   75
         TabIndex        =   12
         Top             =   270
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   2646
         _Version        =   393217
         BackColor       =   15663102
         Enabled         =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis204.frx":1133
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvwPatient 
      Height          =   555
      Left            =   3045
      TabIndex        =   7
      Tag             =   "20113"
      Top             =   1410
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   979
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   15857140
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwWS 
      Height          =   6765
      Left            =   60
      TabIndex        =   6
      Top             =   1410
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   11933
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15857140
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Comment by Accession No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   3075
      TabIndex        =   17
      Tag             =   "20003"
      Top             =   6690
      Width           =   5760
      Begin VB.CommandButton cmdRemarkTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
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
         Left            =   5400
         Picture         =   "Lis204.frx":13A6
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   1440
         Width           =   315
      End
      Begin VB.CommandButton cmdCommentTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
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
         Left            =   5400
         Picture         =   "Lis204.frx":18D8
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   825
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfComment 
         Height          =   900
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   1588
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis204.frx":1E0A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfRemark 
         Height          =   360
         Left            =   90
         TabIndex        =   23
         Top             =   1410
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   635
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"Lis204.frx":203C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   1155
         Width           =   1545
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   345
      Left            =   60
      TabIndex        =   31
      Top             =   8175
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
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
      Alignment       =   1
      Caption         =   "워크쉬트 건수"
      LeftGab         =   0
   End
   Begin VB.Frame fraCul 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '없음
      Height          =   555
      Left            =   8085
      TabIndex        =   36
      Top             =   8520
      Width           =   2535
      Begin VB.CheckBox chkCul 
         BackColor       =   &H00DBE6E6&
         Caption         =   "부분"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   38
         Top             =   90
         Width           =   960
      End
      Begin VB.CommandButton cmdCul 
         BackColor       =   &H00F4F0F2&
         Caption         =   "누적결과조회"
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1080
         Style           =   1  '그래픽
         TabIndex        =   37
         Tag             =   "135"
         Top             =   15
         Width           =   1320
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   1
      Left            =   7185
      TabIndex        =   39
      Top             =   525
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "연 락 처"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   2
      Left            =   7185
      TabIndex        =   40
      Top             =   810
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "상 병 명"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   3
      Left            =   7185
      TabIndex        =   41
      Top             =   1095
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "관련검사 결과"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblTelno 
      Height          =   270
      Left            =   8475
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   525
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   476
      BackColor       =   16777215
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.Frame fraWS 
      BackColor       =   &H00DBE6E6&
      Height          =   1455
      Left            =   75
      TabIndex        =   15
      Top             =   -45
      Width           =   7065
      Begin VB.CheckBox chkStatFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "응급검체 우선"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5535
         TabIndex        =   4
         Top             =   495
         Width           =   1455
      End
      Begin VB.CommandButton cmdWSList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2205
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   360
         Width           =   285
      End
      Begin MedControls1.LisLabel lblWorkCdNm 
         Height          =   330
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         BackColor       =   16252927
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FFF7FC&
         Caption         =   "&Query"
         Height          =   510
         Left            =   5520
         MaskColor       =   &H00808080&
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   780
         Width           =   1320
      End
      Begin MSMask.MaskEdBox mskFrWorkNo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   3600
         TabIndex        =   2
         Top             =   840
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   15857140
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "####"
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   4155
         TabIndex        =   9
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtWorkCd 
         Appearance      =   0  '평면
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   990
         TabIndex        =   0
         Top             =   360
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dptWorkDt 
         Height          =   330
         Left            =   990
         TabIndex        =   1
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy'-'MM'-'dd"
         Format          =   83558403
         CurrentDate     =   36287
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   5235
         TabIndex        =   10
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox mskToWorkNo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   4680
         TabIndex        =   3
         Top             =   840
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   15857140
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "####"
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   60
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "WS 코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   825
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "작업일자"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   2625
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   840
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "작업번호"
         Appearance      =   0
      End
      Begin VB.Line Line1 
         X1              =   4500
         X2              =   4575
         Y1              =   975
         Y2              =   975
      End
   End
   Begin VB.ListBox lstWSCode 
      Appearance      =   0  '평면
      BackColor       =   &H00FFF9F7&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   1050
      TabIndex        =   28
      Top             =   630
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Frame fraMesg 
      BackColor       =   &H00DBE6E6&
      Height          =   2655
      Left            =   10290
      TabIndex        =   50
      Top             =   1005
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtMesg 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   52
         Top             =   390
         Width           =   4050
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00DBE6E6&
         Caption         =   "확인"
         Height          =   420
         Left            =   2940
         Style           =   1  '그래픽
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1095
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   2
         Left            =   15
         TabIndex        =   53
         Top             =   90
         Width           =   4050
         _ExtentX        =   7144
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
         Caption         =   "처방 비고사항 조회"
         Appearance      =   0
         LeftGab         =   200
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   330
      Index           =   5
      Left            =   10290
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   6285
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   582
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "배치결과"
      Appearance      =   0
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackColor       =   &H00DDF0F5&
      BackStyle       =   0  '투명
      Caption         =   "오류가 발생했다."
      ForeColor       =   &H00313D46&
      Height          =   180
      Left            =   165
      TabIndex        =   32
      Top             =   8670
      Width           =   1380
   End
   Begin VB.Label lblWSCnt 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '단일 고정
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      TabIndex        =   19
      Tag             =   "20304"
      Top             =   8175
      Width           =   1020
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF9F7&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   8580
      Width           =   7815
   End
End
Attribute VB_Name = "frm204WSDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private insForm                 As Form
Private gintTemplete            As Integer

Private WithEvents clsTemplete   As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1
Private WithEvents objCuM       As frmTmpCumulative
Attribute objCuM.VB_VarHelpID = -1
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1

Private objLab032       As clsComcode032
Private objLab301       As clsWSList
Private objPtInfo       As clsPatientInfo

Private blnFirst        As Boolean
Private blnDayCount     As Boolean

Private gstrPtAddInfo   As String
Private gblnNewObj      As Boolean
Private gblnModify      As Boolean
Private gstrModifyData  As String

Private IndexPointer    As Integer   'List View 의 index pointer
Private MsgFg           As Boolean
Private LeaveCellFg As Boolean

Private strCombo        As String
Private blnRstChange    As Boolean

Private Sub cmdApply_Click()
    Dim strBatchRst As String
    Dim i           As Long
    
    If Trim(txtBatchRst.Text) = "" Then
        txtBatchRst.SetFocus
        Exit Sub
    End If
    
    '** 중요사항
    ' * 컬럼 중 기존에 사용치 않고 있는 14는 배치결과 등록 시 사용한다.
    strBatchRst = Trim(txtBatchRst.Text)
    i = 1
    With ssRst
        If .DataRowCnt = 0 Then
            Exit Sub
        End If
        
        For i = 1 To .DataRowCnt
            .Row = i
            
            .Col = 1
            If .BackColor = vbGrayText Then
                GoTo Skip
            End If
            
            .Col = 2:
            If Trim(.Value) = "" Then
                .Value = strBatchRst
                .ForeColor = DCM_LightRed
                
                '-- Batch Result
                .Col = .MaxCols: .Value = strBatchRst
                
                '-- Batch Result Check Flag
                .Col = 14: .Value = 1
            End If
            
Skip:

        Next
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Dim i       As Long
    
    i = 1
    With ssRst
        If .DataRowCnt = 0 Then
            Exit Sub
        End If
        
        For i = 1 To .DataRowCnt
            .Row = i
            
            '-- Batch Result
            .Col = 14:
            If Trim(.Value) = 1 Then
                .Value = ""
                
                '-- Batch Result
                .Col = 2: .Value = ""
                
                '-- Batch Result Check Flag
                .Col = 14: .Value = ""
            End If
        Next
    End With
    
End Sub

Private Sub cmdCul_Click()
    Dim objTestCd   As New clsDictionary
    Dim sPtid       As String
    Dim ii          As Integer
    
    Me.MousePointer = vbHourglass
    
    Set objCuM = New frmTmpCumulative

    objTestCd.Clear
    objTestCd.FieldInialize "testcd", "spccd"

    objTestCd.Sort = False
    For ii = 1 To ssRst.MaxRows
        ssRst.Row = ii
        ssRst.Col = 1
        With objPtInfo.Result.Item(ii)
            If chkCul.Value = 0 Then
                If objTestCd.Exists("testcd") = False Then
                    objTestCd.AddNew .TestCd, .SpcCd
                End If
            Else
                If ssRst.ForeColor = DCM_LightRed Then
                    If objTestCd.Exists("testcd") = False Then
                        objTestCd.AddNew .TestCd, .SpcCd
                    End If
                End If
            End If
            sPtid = objPtInfo.PtId
        End With
    Next ii
    objTestCd.Sort = True

    With objCuM
        .Top = Me.Top + 2000
        .Left = Me.Left + 200
        .MousePointer = vbDefault
        .Caption = "환자ID: " & sPtid & " 누적결과"
        Call .DisplayItem(objTestCd, sPtid)
        DoEvents
        
        .WindowState = 0
        .Show vbModal
        DoEvents
    End With

    Set objTestCd = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdSpecial_Click()
    frmRealTestShow.DrFrame1.ZOrder 0
    frmRealTestShow.LisLabel7(0).Caption = "특수검사 관련검사 리스트"
    frmRealTestShow.Show
    Call frmRealTestShow.SpecialTest(lvwPatient.ListItems.Item(1).Text, lvwPatient.ListItems.Item(1).SubItems(1), cboRelTest, "1")
End Sub
Private Sub cmdMicro_Click()
    frmRealTestShow.DrFrame2.ZOrder 0
    frmRealTestShow.LisLabel7(0).Caption = "미생물 관련검사 리스트"
    frmRealTestShow.Show
    Call frmRealTestShow.SpecialTest(lvwPatient.ListItems.Item(1).Text, lvwPatient.ListItems.Item(1).SubItems(1), cboRelTest, "2")
End Sub

Private Sub chkStatFg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdClear_Click()
    Call ClearData
End Sub

Private Sub cmdExit_Click()

    Dim intYesNo As VbMsgBoxResult
   
    If gblnModify = True Then
        If DataFetch <> gstrModifyData Then
            intYesNo = MsgBox("자료가 수정되었습니다." & vbNewLine & "수정된 자료를 저장하시겠슴니까?", vbYesNo, "결과등록")
            If intYesNo = vbYes Then Call cmdSave_Click    '데이타 저장
        End If
        gblnModify = False: gstrModifyData = ""
    End If

    Set clsTemplete = Nothing
    Set objLab301 = Nothing
    Set objPtInfo = Nothing
    Unload Me
    Set frm204WSDataEntry = Nothing
    
End Sub

Private Sub cmdNext_Click()
    
    Dim objLvwItem As MSComctlLib.ListItem

    If lvwWS.ListItems.Count > IndexPointer Then
        Set objLvwItem = lvwWS.ListItems.Item(IndexPointer + 1)
        lvwWS.ListItems(IndexPointer + 1).EnsureVisible
        lvwWS_ItemClick objLvwItem
    End If

End Sub

Private Sub cmdPrevious_Click()
    
    Dim objLvwItem As MSComctlLib.ListItem

    If IndexPointer > 1 Then
        Set objLvwItem = lvwWS.ListItems.Item(IndexPointer - 1)
        lvwWS.ListItems(IndexPointer - 1).EnsureVisible
        lvwWS_ItemClick objLvwItem
    End If

End Sub

Private Sub cmdQuery_Click()
    
    Dim objLvwItem As MSComctlLib.ListItem
    Dim i As Integer
   '
    If txtWorkCd.Text = "" Then Exit Sub
   '
    MouseRunning

    Set objLab301 = New clsWSList
    With objLab301
        .LoadTable txtWorkCd.Text, _
                     DateStr(dptWorkDt.Value), mskFrWorkNo.ClipText, mskToWorkNo.ClipText
        medDataLoadLvw lvwWS, vbNewLine, vbTab, .GetStrWSList
        DoEvents

        For i = 1 To lvwWS.ListItems.Count
            Set objLvwItem = lvwWS.ListItems.Item(i)
            If objLvwItem.SubItems(3) <> "" Then objLvwItem.ForeColor = vbRed
        Next

        With lvwWS
            If chkStatFg.Value = 1 Then
                .SortKey = 3
                .Sorted = True
            Else
                '.SortKey = 0
                .Sorted = False
            End If
            .FlatScrollBar = True
        End With

        If .RecordCount > 0 Then
            EditData
            DisplayCount
            lblErr.Caption = ""
            IndexPointer = 1
            Set objLvwItem = lvwWS.ListItems.Item(1)
            lvwWS_ItemClick objLvwItem
        Else
            MsgBox "해당 데이타가 없습니다."
            ClearData
        End If
    End With
    '
    MouseDefault

End Sub

Private Sub cmdRemarkTemplete_Click()
   
'    Dim SqlStmt As String
'
'    Set objCodeList = Nothing
'    Set objCodeList = New clsPopUpList
'
'    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark)
    
    Dim SqlStmt As String
    Dim RS      As Recordset
    Dim strWorkArea As String
    
    
    Set objCodeList = Nothing
    Set objCodeList = New clsPopUpList
    strWorkArea = medGetP(lvwPatient.ListItems.Item(1).Text, 1, "-")
    
    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE " & DBW("cdindex = ", LC4_Remark) & " and " & DBW("field1=", strWorkArea)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    If RS.EOF Then
        SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE " & DBW("cdindex = ", LC4_Remark)
    End If
    Set RS = Nothing
    
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Remark"
        .ColumnHeaderText = "Code;Remark"
'        .HideColumnHeaders = True
        .ColumnHeaderWidth = "840.189;5309.858"
        .FormHeight = 3105
        .FormWidth = 6605
        .HideSearchTool = True
        .SelectByClick = True
        .Tag = "Remark"
        .LoadPopUp SqlStmt
        
'        .ListCaption = "Remark"
'        .ListColHeader = "Code" & vbTab & "Remark"
'        .Top = Me.cmdRemarkTemplete.Top + 5700
'        .Left = Me.cmdRemarkTemplete.Left + 2000
'        .Width = 6250
'        .Height = 3000
'        .Tag = "Remark"
'        .CaptionOn = True
'        .MultiSel = False
'        .PopupList SqlStmt, 2
'        .ListAdd vbTab & "< 없 음 > ", 2, 1
   End With

End Sub

Private Sub cmdSave_Click()
    
    Dim ii As Long
    Dim blnDBSuccess As Boolean
    Dim objLvwItem As MSComctlLib.ListItem
    Dim intLvwCount As Integer
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    '
    With objPtInfo
        .FootNote = rtfComment.Text
        .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
    End With
    '/*
    For ii = 1 To ssRst.MaxRows
        With objPtInfo.Result.Item(ii)
            ssRst.Row = ii
            ssRst.Col = objPtInfo.SSCol("RESULT")
'            If ssRst.Value = CS_EqpError Then
            If UCase(ssRst.Value) = UCase(CS_EqpError) Then
                ssRst.Action = ActionActiveCell
                Exit Sub
            End If
            'If .TxtType = "2" Then
            If .TxtType = "2" And .RstDiv = "R" Then
                If .TextRst = "" Or ssRst.Value = "" Then
                'If (ssRst.Value <> "" AND .TextRst = "") _
                            Or (ssRst.Value = "" AND .TextRst <> "") Then
                            '검사는 일반결과와 텍스트 결과를 같이 입력요. 결과보류 처리.
                    ssRst.Col = objPtInfo.SSCol("EC")
                    ssRst.Value = 1
                End If
            End If
        End With
    Next ii
    '
    blnDBSuccess = objPtInfo.DataEntry 'objPtInfo                  '결과등록을 수행한다.
    If blnDBSuccess = False Then
        MsgBox objPtInfo.ErrNo & " - " & objPtInfo.ErrText, vbCritical + vbOKOnly, "결과등록 ERROR"
        Exit Sub
    Else
        
        If P_RealPrinter = True Then
            '결과지 응급실 수술실 보내기
'            DoEvents
            With lvwWS
                strWorkArea = medGetP(.SelectedItem.ListSubItems(1).Text, 1, "-")
                strAccDt = Mid(Format(GetSystemDate, "YYYY"), 1, 2) & medGetP(.SelectedItem.ListSubItems(1).Text, 2, "-")
                strAccSeq = medGetP(.SelectedItem.ListSubItems(1).Text, 3, "-")
        
                Call PrintEROP24(strWorkArea, strAccDt, strAccSeq)
            End With
'            DoEvents
        End If
        
        lblErr.Caption = "자료가 정상적으로 보관되었읍니다."
    End If
   '
    If Not objPtInfo.WSDataExists(txtWorkCd.Text) Then
        ssRst.MaxRows = 0
        lvwPatient.ListItems.Clear
        rtfText.Text = ""
        rtfComment.Text = ""
        rtfRemark.Text = ""
        With lvwWS
            intLvwCount = .ListItems.Count
            For ii = 1 To .ListItems.Count
                If .ListItems.Item(ii).Selected = True Then
                    .ListItems.Remove (ii)
                    Exit For
                End If
            Next ii
            If intLvwCount = .ListItems.Count Then
                For ii = 1 To .ListItems.Count
                    If .ListItems.Item(1).SubItems(1) = objPtInfo.AccNo Then
                        .ListItems.Remove (ii)
                        Exit For
                    End If
                Next ii
            End If
        End With
        IndexPointer = IndexPointer - 1
        If lvwWS.ListItems.Count = IndexPointer Then IndexPointer = IndexPointer - 1
    End If
    '
    If lvwWS.ListItems.Count = IndexPointer Then
        Set objLvwItem = lvwWS.ListItems(IndexPointer)
        objLvwItem.SubItems(2) = " "
        IndexPointer = 0
    End If
    If lvwWS.ListItems.Count = 0 Then
        ClearData
    Else
        gblnModify = False
        Call cmdNext_Click
        'Set objLvwItem = lvwWS.ListItems.Item(1)
        'lvwWS_ItemClick objLvwItem
    End If
   '
End Sub

Private Sub PrintEROP24(ByVal LastWorkArea As String, ByVal LastAccDt As String, ByVal LastAccSeq As String)
    Dim RS              As Recordset
    Dim objReport       As clsBatchReport
    Dim objSQL          As clsLISSqlReport
    Dim objDisease      As S2LIS_ReportLib.clsDisease
    Dim picESign        As Object
    Dim strSQL          As String
    Dim strEmpId        As String
    Dim strAge          As String
    Dim strWardID       As String
    
    Set objReport = New clsBatchReport
    Set objSQL = New clsLISSqlReport
    Set objDisease = New S2LIS_ReportLib.clsDisease
    
    '오라클 기준으로 설정
    strSQL = " SELECT a.ptid,a.workarea,a.accdt,a.accseq,a.stscd,a.vfydt,a.vfytm, " & _
             "        d." & F_PTNM & " as ptnm, d." & F_DOB & " as dob, d." & F_SEX & " as sex, " & _
             "        c.deptcd, c.wardid, c.majdoct " & _
             "   FROM " & T_HIS001 & " d, " & T_LAB101 & " c, " & T_LAB102 & " b, " & T_LAB201 & " a " & _
             "  WHERE " & DBW("a.workarea", LastWorkArea, 2) & _
             "    AND " & DBW("a.accdt", LastAccDt, 2) & _
             "    AND " & DBW("a.accseq", LastAccSeq, 2) & _
             "    AND a.reqtotcnt=a.reqinputcnt " & _
             "    AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
             "    AND b.ptid=c.ptid AND b.orddt=c.orddt AND b.ordno=c.ordno " & _
             "    AND c.deptcd in ('ER','24') AND b.ptid = d." & F_PTID
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.RecordCount > 0 Then '
        
        With objReport
            .PtId = RS.Fields("ptid").Value & ""
            .ptnm = RS.Fields("ptnm").Value & ""
            .PtSex = RS.Fields("sex").Value & ""
            strAge = RS.Fields("dob").Value & ""
            If Len(strAge) = 6 Then strAge = strAge & "01"
            .PtAge = DateDiff("yyyy", Format(strAge, CS_DateMask), GetSystemDate)
                
            .FromDt = RS.Fields("vfydt").Value & ""
            .ToDt = RS.Fields("vfydt").Value & ""
            
            If Trim(RS.Fields("bussdiv").Value & "") = "1" Then
                .Dept = RS.Fields("deptcd").Value & ""
            Else
                .Dept = RS.Fields("wardid").Value & ""
            End If
            
            .VfyDt = RS.Fields("vfydt").Value & ""
'            If objLisComCode.DeptCd.Exists(Rs.Fields("deptcd").Value & "") Then
'                objLisComCode.DeptCd.KeyChange (Rs.Fields("deptcd").Value & "")
                .DeptNm = GetDeptNm(RS.Fields("deptcd").Value & "") 'objLisComCode.DeptCd.Fields("deptnm")
'            End If
            
            strWardID = Trim(RS.Fields("wardid").Value & "")
            
            .WardId = strWardID
            
'            If objLisComCode.WardId.Exists(strWardID) = True Then
'                objLisComCode.WardId.KeyChange (strWardID)
                objReport.Ward = GetWardNm(strWardID) 'objLisComCode.WardId.Fields("wardnm")
'            End If
            .Doct = GetEmpNm(RS.Fields("majdoct").Value & "")
            .VfyNM = ObjMyUser.EmpLngNm
            .MdfDt = ""
            objDisease.PtId = RS.Fields("ptid").Value
            .ICD = objDisease.Disease
            Call .ReportForOneERPatient(RS.Fields("ptid").Value & "", RS.Fields("vfydt").Value & "", _
                                   RS.Fields("workarea").Value & "", RS.Fields("accdt").Value & "", _
                                   RS.Fields("accseq").Value & "", picESign, RS.Fields("vfydt").Value & "", _
                                   RS.Fields("vfydt").Value & "")
        End With
    End If
    
    If P_PrinterChkFg = True Then
        strSQL = " update " & T_LAB302 & _
                 "    set " & _
                            DBW("rptfg", "Y", 3) & _
                            DBW("rptdt", Format(GetSystemDate, "YYYYMMDD"), 3) & _
                            DBW("rptid", ObjSysInfo.EmpId, 2) & _
                 "  WHERE " & DBW("a.workarea", LastWorkArea, 2) & _
                 "    AND " & DBW("a.accdt", LastAccDt, 2) & _
                 "    AND " & DBW("a.accseq", LastAccSeq, 2)
        
        DBConn.Execute strSQL
    End If
    
NoData:
    Set RS = Nothing
    Set objSQL = Nothing
    Set objReport = Nothing
    Set objDisease = Nothing
End Sub

'Private Function GetEmpNm(ByVal vEmpID As String) As String
'    Dim objData As New clsBasisData
'
'    GetEmpNm = objData.GetEmpNm(vEmpID)
'    Set objData = Nothing
'End Function

'Private Function GetDeptNm(ByVal vDeptCd As String) As String
'    Dim objData As New clsBasisData
'
'    GetDeptNm = objData.GetDeptNm(vDeptCd)
'    Set objData = Nothing
'End Function

'Private Function GetWardNm(ByVal vWardId As String) As String
'    Dim objData As New clsBasisData
'
'    GetWardNm = objData.GetWardNm(vWardId)
'    Set objData = Nothing
'End Function

Private Sub cmdTextTemplete_Click()
    If rtfText.Enabled = False Then Exit Sub
    Call CallTemplete(2, 0)
End Sub

Private Sub cmdCommentTemplete_Click()
    If ssRst.MaxRows < 1 Then Exit Sub
    Call CallTemplete(3, 0)
End Sub

Private Sub cmdWSList_Click()

    If lstWSCode.ListCount = 0 Then
        MsgBox "등록된 Worksheet 코드가 없습니다.", vbExclamation, "메세지"
        Exit Sub
    End If
    lstWSCode.Visible = True
    lstWSCode.ZOrder 0
    Call medCodeHelp(0, lstWSCode, txtWorkCd.Text, txtWorkCd, dptWorkDt)

End Sub

Private Sub dptWorkDt_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub dptWorkDt_Validate(Cancel As Boolean)

    Dim SqlStmt As String
    Dim tmpRs As Recordset

    mskFrWorkNo.Text = "1___"
    If txtWorkCd.Text <> "" Then
        SqlStmt = "SELECT max(workseq) as MaxSeq FROM " & T_LAB301 & _
                      " WHERE  " & DBW("workcd =", txtWorkCd.Text) & " " & _
                      " AND  " & DBW("workdt =", Format(dptWorkDt.Value, CS_DateDbFormat)) & " "
        Set tmpRs = New Recordset
        tmpRs.Open SqlStmt, DBConn
        
        If tmpRs.EOF Then GoTo NoData

        If tmpRs.Fields("MaxSeq").Value <> "" Then
            mskToWorkNo.Text = tmpRs.Fields("MaxSeq").Value & String(4 - Len(tmpRs.Fields("MaxSeq").Value), "_")
        Else
            mskToWorkNo.Text = "1___"
        End If

NoData:
        Set tmpRs = Nothing
    End If

End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
    '
    If blnFirst = False Then
        Call LoadLvwHead
        blnFirst = True
        ClearData
    End If
    '
    If objLab301 Is Nothing Then
        Set objLab301 = New clsWSList
        objLab301.LoadTable , , "", ""
    End If
    '누적결과및 관련검사(미생물/특수조회여부)
    If P_RealTestMicSpecial = True Then fraCul.Visible = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If lstWSCode.Visible Then lstWSCode.Visible = False
End Sub

Private Sub Form_Load()
    '
    Me.Show
    Call cmdClear_Click
    blnFirst = False
    gblnModify = False
    'Set OraSE = CreateObject("OracleInProcServer.XOraSession")
    'Set OraDB = OraSE.OpenDatabase(DatabaseName$, Connect$, 0&)
    dptWorkDt.Value = Date
    dptWorkDt.MaxDate = Date
    '
    prgRst.Align = vbAlignBottom
    prgRst.Visible = False
    ssRst.RowHeight(-1) = 12.5
    '
    cboRelTest.Clear
    cboRelTest.AddItem "관련 검사의 최근결과"
    cboRelTest.ListIndex = 0
    '
    Set objPtInfo = New clsPatientInfo
    Call objPtInfo.LoadWorksheetCode(ObjSysInfo.BuildingCd, lstWSCode)
    KeyPreview = True
End Sub

Private Sub clsTemplete_CopyTemplete()
    '
    If ssRst.MaxRows < 1 Then Exit Sub
    With objPtInfo
        Select Case gintTemplete
            Case 1:
                If clsTemplete.rtfText.Text <> "" Then
                    rtfRemark.Text = clsTemplete.rtfText.Text
                    .RmkCd = frm230TempSearch.lblCode.Caption
                    .RmkNm = rtfRemark.Text
                Else
                    rtfRemark.Text = ""
                    .RmkCd = ""
                    .RmkNm = ""
                End If
            Case 2:
                rtfText.Text = clsTemplete.rtfText.Text
                .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
                rtfText.SetFocus
            Case 3:
                rtfComment.Text = clsTemplete.rtfText.Text
                .FootNote = rtfComment.Text
                rtfComment.SetFocus
        End Select
    End With
    Set clsTemplete = Nothing

End Sub

Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
    
    Dim strTitle As String
    Dim strWorkArea As String
    
    Set clsTemplete = New frm230TempSearch
    strTitle = Choose(pintPrg, "Remark", "Text Result", "Foot Note")
    strWorkArea = medGetP(lvwPatient.ListItems.Item(1).Text, 1, "-")
    With clsTemplete
        .qField1 = strWorkArea
        .Show
        If pintMode = 0 Then
            .lblName.Caption = "Edit " & strTitle
        Else
            .lblName.Caption = "Modify " & strTitle
        End If
        .Caption = strTitle & " " & "Templete Editor"
        .lblInfo.Caption = pintMode & "$" & pintPrg
        Select Case pintPrg
            Case 1:
                .lblCode.Caption = objPtInfo.RmkCd
                .rtfText = rtfRemark.Text
            Case 2:
                .rtfText = rtfText.Text
            Case 3:
                .rtfText = rtfComment.Text
        End Select
    End With
    gintTemplete = pintPrg
    
End Sub

Private Sub LoadLvwHead()
    
    Dim colHead As ColumnHeader
    Dim intMode As Integer
   
    '국가별 설정 모드
    intMode = 1         'Korea
    'intMode = 2         'English
    If intMode = 1 Then
'        medInitLvwHead lvwPatient, "접수번호,환자ID,환자성명,성/나이,생년월일,병상,주치의,검체,접수일자", _
'                                   "550,-100,0,-450,-20,70,-50,0"
        medInitLvwHead lvwPatient, "접수번호,환자ID,환자성명,성/나이,생년월일,병상,주치의,검체,접수일자,비고(외부QC)", _
                                   "550,-100,0,-450,-20,70,-50,50,0"
        medInitLvwHead lvwWS, "No,접수번호,,응급여부 ", _
                              "-150,1200,-340,-50"
    Else
        medInitLvwHead lvwPatient, "Accession#,Patient ID,Patient Name,Sex/Age,Location,Physician", _
                                   "200,-100,200,-400,0,100,0"
        medInitLvwHead lvwWS, "No,Accession No,,StatFg ", _
                              "-300,700,0,100"
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsTemplete = Nothing
    Set objCodeList = Nothing
    Set objCuM = Nothing
    Set objPtInfo = Nothing
    Set objLab032 = Nothing
    Set objLab301 = Nothing
    Call ICSPatientMark
End Sub

Private Sub lstWSCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn And lstWSCode.ListIndex >= 0 Then
        txtWorkCd.Text = Trim(Mid(lstWSCode.Text, 1, _
                 InStr(1, lstWSCode.Text, vbTab) - 1))
        lblWorkCdNm.Caption = medGetP(lstWSCode.Text, 2, vbTab)
        lstWSCode.Visible = False
        dptWorkDt.SetFocus
    End If

End Sub

Private Sub lstWSCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call lstWSCode_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstWSCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lstWSCode.SetFocus
End Sub

Private Sub mskFrWorkNo_GotFocus()
    FocusMe Me.mskFrWorkNo
End Sub

Private Sub mskFrWorkNo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub mskToWorkNo_GotFocus()
    FocusMe Me.mskToWorkNo
End Sub

Private Sub mskToWorkNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

'Private Sub objCodeList_ListClick(ByVal SelList As String)
'
'    Dim strTmp As String
'   '
'    If Not IsNull(SelList) And SelList <> "" Then
'        Select Case objCodeList.Tag
'            Case "Remark":
'                objPtInfo.RmkCd = medGetP(SelList, 1, vbTab)
'                If Trim(objPtInfo.RmkCd) <> "" Then
'                    objPtInfo.RmkNm = medGetP(SelList, 2, vbTab)
'                Else
'                    objPtInfo.RmkNm = ""
'                End If
'                rtfRemark.Text = objPtInfo.RmkNm
'        End Select
'    End If
'    Set objCodeList = Nothing
'   '
'End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    Dim strTmp As String
   '
'    If Not IsNull(SelList) And SelList <> "" Then
        Select Case objCodeList.Tag
            Case "Remark":
                objPtInfo.RmkCd = medGetP(pSelectedItem, 1, ";")
                If Trim(objPtInfo.RmkCd) <> "" Then
                    objPtInfo.RmkNm = medGetP(pSelectedItem, 2, ";")
                Else
                    objPtInfo.RmkNm = ""
                End If
                rtfRemark.Text = objPtInfo.RmkNm
        End Select
'    End If
    Set objCodeList = Nothing
End Sub

Private Sub rtfText_LostFocus()
    '
    objPtInfo.Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
    '
End Sub

Private Sub ssRst_EditChange(ByVal Col As Long, ByVal Row As Long)
    ssRst.Row = Row
    ssRst.Col = objPtInfo.SSCol("MAXCOL")
    ssRst.Value = ""
End Sub

Private Sub txtBatchRst_GotFocus()
    With txtBatchRst
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtBatchRst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdApply.SetFocus
    End If
End Sub

Private Sub txtWorkCd_Change()

    If txtWorkCd.Text = "" Then lblWorkCdNm.Caption = ""

End Sub

Private Sub txtWorkCd_GotFocus()
    '
    FocusMe Me.txtWorkCd
    '
End Sub

Private Sub txtWorkCd_KeyPress(KeyAscii As Integer)
    
    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = vbKeyEscape Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call lstWSCode_KeyDown(vbKeyReturn, 0)
        lstWSCode.Visible = False
        Exit Sub
    End If

    lstWSCode.Visible = True
    lstWSCode.ZOrder 0
    Call medCodeHelp(KeyAscii, lstWSCode, txtWorkCd.Text, txtWorkCd, dptWorkDt)

End Sub

Private Sub txtWorkCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstWSCode.ListCount = 0 Then Exit Sub
    If KeyCode = vbKeyDown Then
        lstWSCode.Visible = True
        If lstWSCode.ListIndex < lstWSCode.ListCount - 1 Then lstWSCode.ListIndex = lstWSCode.ListIndex + 1
        lstWSCode.ZOrder 0
        lstWSCode.SetFocus
    End If
End Sub



Private Sub txtWorkCd_Validate(Cancel As Boolean)
    '
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub

    If txtWorkCd.Text = "" Then Exit Sub

    IndexPointer = 0
    lblWorkCdNm.Caption = ""
    If Trim(txtWorkCd.Text) = "" Then
        Cancel = True
        Exit Sub
    End If
    '
    If objLab301.IsWorkCd(txtWorkCd.Text) = False Then
        MsgBox "코드 입력 Error!", vbCritical
        Cancel = True
        FocusMe Me.txtWorkCd
        Exit Sub
    End If
    '
    Set objLab032 = New clsComcode032
    With objLab032
        .LoadTable LC3_WorkSheetName, , txtWorkCd.Text
        .MoveFirst
        If .RecordCount > 0 Then
            If Trim(ObjSysInfo.BuildingCd) <> Trim(.Field2) Then
                MsgBox "현재 건물에서는 사용할 수 없는 코드입니다.", vbCritical
                Cancel = True
                FocusMe Me.txtWorkCd
                Exit Sub
            End If
            lblWorkCdNm.Caption = .Field1
        End If
    End With
    Set objLab032 = Nothing
    lstWSCode.Visible = False
   '
End Sub


Private Sub lvwWS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim strLvw As String
    Dim strCurrentData As String
    Dim intYesNo As VbMsgBoxResult
    Dim objLvwItem As MSComctlLib.ListItem
    Dim ii As Integer
    '
    If gblnModify = True Then
        objPtInfo.FootNote = rtfComment.Text
        objPtInfo.Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
        If DataFetch <> gstrModifyData Then
            intYesNo = MsgBox("자료가 수정되었읍니다." & vbNewLine & "수정된 자료를 저장하시겠습니까?", _
                                vbYesNo, "결과등록")
            If intYesNo = vbYes Then Call cmdSave_Click    '데이타 저장
        End If
        gblnModify = False: gstrModifyData = ""
    End If
    '
    lvwPatient.ListItems.Clear
    ssRst.MaxRows = 0
    rtfText.Text = ""
    rtfComment.Text = ""
    rtfRemark.Text = ""
    cboRelTest.Clear
    CmdTemplete False
    DoEvents
    '
    If objPtInfo Is Nothing Then
        Set objPtInfo = New clsPatientInfo
    Else
        Set objPtInfo = Nothing
        Set objPtInfo = New clsPatientInfo
    End If
    '
    If IndexPointer > 0 Then
        Set objLvwItem = lvwWS.ListItems(IndexPointer)
        objLvwItem.SubItems(2) = " "
    End If

    IndexPointer = Item.Index
    Item.SubItems(2) = "◀"
    Item.Selected = True
    If IndexPointer = lvwWS.ListItems.Count Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    If IndexPointer = 1 Then
        cmdPrevious.Enabled = False
    Else
        cmdPrevious.Enabled = True
    End If
    strLvw = LvwClickData(Item)
    PtResultLoad medGetP(strLvw, 2, vbTab)
    '
    '
    '/* 새로운 워크쉬트별 결과를 조회해서 들어갈때 처음 로드된 상태에서 수정하고
    '  빠져나갈때 데이터가 변했는지 확인하기 위해 gblnModify,gstrModifyData를 이용.
    '  gblnModify = True : 데이터 수정시작,gstrModifyData : 수정전 데이터
    DoEvents
    If ssRst.MaxRows > 0 Then
        gblnModify = True
        gstrModifyData = DataFetch()
        ssRst.SetFocus
    End If

   '
End Sub


Private Function DataFetch() As String
    
    Dim ii As Integer
    
    DataFetch = ""
    With ssRst
        .Col = objPtInfo.SSCol("RESULT"): .COL2 = objPtInfo.SSCol("EC")
        .Row = 1: .Row2 = .MaxRows
        DataFetch = .Clip & "$"
    End With
    With objPtInfo
        DataFetch = DataFetch & .FootNote & "$" & .RmkCd & "$"
        For ii = 1 To ssRst.MaxRows
            DataFetch = DataFetch & .Result.Item(ii).TextRst
        Next ii
    End With
    
End Function

Private Sub ClearData()
    gblnModify = False
    txtWorkCd.Text = ""
    lblWorkCdNm.Caption = ""
    lblErr.Caption = ""
    lblDisease.Caption = ""
    lblTelno.Caption = ""
    dptWorkDt.Value = Date
    fraWS.Enabled = True
    If blnFirst = True Then
        txtWorkCd.SetFocus
    End If
    '
    dptWorkDt.Value = Date
    mskFrWorkNo.Text = "1___"
    mskToWorkNo.Text = "1___"
    ssRst.MaxRows = 0
    ssRst.Enabled = False
    txtWorkCd.BackColor = vbWhite
    dptWorkDt.Enabled = True
    cmdQuery.Enabled = True
    cmdSave.Enabled = False
    CmdTemplete False
    '
    cboRelTest.Clear
    lvwWS.ListItems.Clear
    lvwPatient.ListItems.Clear
    mskFrWorkNo.BackColor = vbWhite
    mskToWorkNo.BackColor = vbWhite
    lvwWS.BackColor = DCM_LightGray
    lvwPatient.BackColor = DCM_LightGray
    rtfComment.BackColor = DCM_LightGray
    rtfText.BackColor = DCM_LightGray
    '
    fraComment.Enabled = False
    fraText.Enabled = False
    '
    lblWSCnt.Caption = "0"
    rtfComment.Text = ""
    rtfText.Text = ""
    rtfRemark.Text = ""
    IndexPointer = 0
    cmdRmk.Visible = False
    fraMesg.Visible = False
    MsgFg = False
    LeaveCellFg = False
    cmdApply.Enabled = False
    txtBatchRst.Text = ""
End Sub

Private Sub EditData()
    '
    ssRst.Enabled = True
    '
    txtWorkCd.BackColor = DCM_LightGray
    dptWorkDt.Enabled = False
    cmdQuery.Enabled = False
    cmdSave.Enabled = True
    '
    fraComment.Enabled = True
    fraText.Enabled = True
    '
    fraWS.Enabled = False
    mskFrWorkNo.BackColor = DCM_LightGray
    mskToWorkNo.BackColor = DCM_LightGray
    lvwWS.BackColor = vbWhite
    lvwPatient.BackColor = vbWhite
    rtfComment.BackColor = &HF1F5F4     'vbWhite
    rtfText.BackColor = &HEEFFFE    'vbWhite
    '
End Sub

Private Sub DisplayCount()
    lblWSCnt.Caption = lvwWS.ListItems.Count
End Sub

Private Sub UpDown1_UpClick()
    mskFrWorkNo.Text = FormatUnder(mskFrWorkNo.ClipText, "+")
    mskFrWorkNo.SetFocus
End Sub

Private Sub UpDown1_DownClick()
    mskFrWorkNo.Text = FormatUnder(mskFrWorkNo.ClipText, "-")
    mskFrWorkNo.SetFocus
End Sub

Private Sub UpDown2_DownClick()
    mskToWorkNo.Text = FormatUnder(mskToWorkNo.ClipText, "-")
    If Val(mskToWorkNo.ClipText) = 1 Then
        mskToWorkNo.Text = "9999"
    End If
    mskToWorkNo.SetFocus
End Sub

Private Sub UpDown2_UpClick()
    mskToWorkNo.Text = FormatUnder(mskToWorkNo.ClipText, "+")
    mskToWorkNo.SetFocus
End Sub

Private Function FormatUnder(ByRef strval As String, _
                             ByVal strSign As String) As String
    
    Dim intLen As Integer
    Dim ii As Integer
    
    If strSign = "+" Then
        FormatUnder = FormatUnder & CStr(Val(strval) + 1)
        strval = Val(strval) + 1
    Else
        FormatUnder = FormatUnder & CStr(Val(strval) - 1)
        strval = Val(strval) - 1
    End If
    '
    intLen = 4 - Len(strval)
    For ii = 1 To intLen
        FormatUnder = "_" & FormatUnder
    Next
    
    If Val(strval) < 1 Then
        FormatUnder = "1___"
    End If
    
End Function

Private Sub PtResultLoad(ByVal strAccNo As String)
'
    Dim intLvwCount As Integer
    Dim ii As Integer
    Dim objLvwItem As MSComctlLib.ListItem

    lvwPatient.ListItems.Clear
    
    MouseRunning
    
    Set objPtInfo.prgBar = prgRst
    objPtInfo.PrgBarInit
    ssRst.Visible = False
    
    If fraMesg.Visible Then fraMesg.Visible = False
    If cmdRmk.Visible Then cmdRmk.Visible = False
    
    With objPtInfo
        .PtType = RESULT_BY_WORKSHEET                 '/* 결과등록 유형, 반드시 셋팅 해야 됨./
        .AccNo = strAccNo      '/* 접수번호, 반드시 셋팅 해야 됨./
        
        .LoadTable txtWorkCd.Text, ObjMyUser.EmpId
        
        If .TestCount > 0 Then
            CmdTemplete True
            If lvwPatient.Enabled = False Then
               lvwPatient.Enabled = True
            End If
            medDataLoadLvw lvwPatient, vbNewLine, vbTab, .GetStringPtInfo
            
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.PtId = lvwPatient.ListItems(1).SubItems(1)
            lblDisease.Caption = objDisease.Disease
            lblDisease.ToolTipText = objDisease.Disease
            Set objDisease = Nothing
            '========================================================================================
            '감염관리
            Call ICSPatientMark(lvwPatient.ListItems(1).SubItems(1), enICSNum.LIS_ALL)
            '병동/진료과 연락처(환자ID,CONTROL)
            Call GetPtTelInfo(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq, lblTelno)
            '========================================================================================
            rtfRemark.Text = .RmkNm
            rtfComment.Text = .FootNote
            If objPtInfo.Result.Item(1).TxtType <> "0" Then
                rtfText.Text = objPtInfo.Result.Item(1).TextRst
                rtfText.Enabled = True
                rtfText.BackColor = &HEEFFFE    'vbWhite
                cmdTextTemplete.Enabled = True
            Else
                rtfText.Enabled = False
                rtfText.BackColor = DCM_LightGray
                cmdTextTemplete.Enabled = False
            End If
            .GetResultSpread ssRst, RESULT_BY_DEFAULT

            '관련검사의 결과 ...
            Dim MyResult As New clsLISResultReview
            Dim RS       As Recordset
            Dim SSQL     As String
            
            Call MyResult.GetRelTest(cboRelTest, strAccNo)
            
            '------------------------추가 사항----------------------------
            strCombo = ""
            For ii = 0 To cboRelTest.ListCount - 1
                strCombo = strCombo & cboRelTest.List(ii) & COL_DIV
            Next
            If strCombo <> "" Then strCombo = Mid(strCombo, 1, Len(strCombo) - 1)
            Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(1).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
            
            '처방리마크 조회(있는지 여부만 조회)
            SSQL = MyResult.GetOrderRemark(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq)
            Set RS = New Recordset
            RS.Open SSQL, DBConn
            If Not RS.EOF Then cmdRmk.Visible = True
            
            cboRelTest.ListIndex = 0
            Set RS = Nothing
            Set MyResult = Nothing
            
            '------------------------추가 사항----------------------------
            cmdApply.Enabled = True
        Else
            MsgBox "해당 Worksheet 항목의 결과가 모두 확인됬습니다.", vbCritical + vbOKOnly, "결과등록 Message"
            lblErr.Caption = "해당 Worksheet 항목의 결과가 모두 확인됬습니다."
            ssRst.MaxRows = 0
            lvwPatient.ListItems.Clear
            rtfText.Text = ""
            rtfComment.Text = ""
            rtfRemark.Text = ""
            With lvwWS
                intLvwCount = .ListItems.Count
                For ii = 1 To .ListItems.Count
                    If .ListItems.Item(ii).Selected = True Then
                        .ListItems.Remove (ii)
                        Exit For
                    End If
                Next ii
                If intLvwCount = .ListItems.Count Then
                    For ii = 1 To .ListItems.Count
                        If .ListItems.Item(1).SubItems(1) = objPtInfo.AccNo Then
                            .ListItems.Remove (ii)
                            Exit For
                        End If
                    Next ii
                End If
            End With
            IndexPointer = IndexPointer - 1
            If lvwWS.ListItems.Count = IndexPointer Then IndexPointer = IndexPointer - 1
   '
            If lvwWS.ListItems.Count = IndexPointer Then
                Set objLvwItem = lvwWS.ListItems(IndexPointer)
                objLvwItem.SubItems(2) = " "
                IndexPointer = 0
            End If
            
            If lvwWS.ListItems.Count = 0 Then
                ClearData
            Else
                gblnModify = False
                Call cmdNext_Click
                'Set objLvwItem = lvwWS.ListItems.Item(1)
                'lvwWS_ItemClick objLvwItem
            End If
            
            cmdApply.Enabled = False
            
        End If
    End With
    
    With ssRst
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 4: .ForeColor = DCM_LightRed: .FontBold = True
        Next
    End With
    
'    Dim i       As Integer
'
'    With ssRst
'        For i = 1 To .MaxRows
''            Call ssRst_LeaveCell(2, i, 2, i, False)
'            Call objPtInfo.NumValCheck(i)
'        Next
'    End With
    
    Dim i       As Integer

    With ssRst
        For i = 1 To .DataRowCnt
            .Col = 2: .Row = i
            If IsNumeric(.Text) Then
                Call objPtInfo.NumValCheck(i)
            Else
                Call ssRst_LeaveCell(2, i, 2, i, False)
            End If
        Next
        .Col = 2: .Row = 1
        If IsNumeric(.Text) Then
            Call objPtInfo.NumValCheck(1)
        Else
            Call ssRst_LeaveCell(2, 1, 2, 1, False)
        End If
    End With
    
    ssRst.Visible = True
    
    MouseDefault
    
    objPtInfo.PrgBarClear
    DoEvents
   '
End Sub

Private Sub ssRst_Click(ByVal Col As Long, ByVal Row As Long)
   '
    Dim i As Long
    
    If Row = 0 And Col = 3 Then
        With ssRst
            .Col = 3
            For i = 1 To .MaxRows
                .Row = i
                If .CellType = CellTypeCheckBox Then .Value = 0
            Next
        End With
    End If
    If Row <= 0 Then Exit Sub
    SpDispRtfText
   
   
    '부분누적결과
    If Row = 0 Then Exit Sub
    If Not P_RealTestMicSpecial Then Exit Sub
    
    If Col = 1 Then
        ssRst.Row = Row:        ssRst.Col = Col
        If objPtInfo.Result.Item(Row).RstDiv = "*" Then
            If ssRst.ForeColor = vbWhite Then
                ssRst.ForeColor = DCM_LightRed
            Else
                ssRst.ForeColor = vbWhite
            End If
        Else
            If ssRst.ForeColor = DCM_MidBlue Then
                ssRst.ForeColor = DCM_LightRed
            Else
                ssRst.ForeColor = DCM_MidBlue
            End If
        End If
    End If
    
    chkCul.Value = 0
    For i = 1 To ssRst.DataRowCnt
        ssRst.Row = i: ssRst.Col = 1
        If ssRst.ForeColor = DCM_LightRed Then
            chkCul.Value = 1
        End If
    Next

End Sub

Private Sub ssRst_GotFocus()
    If MsgFg Then Exit Sub
    If LeaveCellFg Then Exit Sub

    With ssRst
        If .MaxRows = 0 Then Exit Sub
        .Row = 1
        .Col = objPtInfo.SSCol("RESULT")
        .Action = ActionActiveCell
        .EditEnterAction = EditEnterActionDown
    End With
End Sub

Private Sub ssRst_KeyUp(KeyCode As Integer, Shift As Integer)
   '
    If KeyCode = 38 Or KeyCode = 40 Then
        SpDispRtfText
    ElseIf KeyCode = vbKeyF2 Then
        Call ssRst_RightClick(1, ssRst.ActiveCol, ssRst.ActiveRow, 100, 100)
    End If
  '
End Sub

Private Sub ssRst_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
  '
    If ClickType <> 1 Then Exit Sub
    
    If MsgFg Then Exit Sub
    If Row <= 0 Then Exit Sub
    objPtInfo.SsTop = picRst.Top
    objPtInfo.SsLeft = picRst.Left
    ssRst.Row = Row
    ssRst.Col = Col
    ssRst.Action = ActionActiveCell
    objPtInfo.MfyFg = False
    MsgFg = True
    Call objPtInfo.PopUp(, Col)
    MsgFg = False
      '
End Sub
'Private Sub ssRst_LostFocus()
'    Dim strTmp          As String
'    Dim strTmp1         As String
'    Dim strUTmp         As String
'    Dim strRstVal       As String
'
'    Dim strResultVal    As String
'    Dim strResultChk    As String
'    Dim strTestCd       As String
'
'    If ssRst.ActiveRow < 1 Then Exit Sub
'
'    ssRst.Row = ssRst.ActiveRow
'    ssRst.Col = objPtInfo.SSCol("RESULT")
'    strTestCd = objPtInfo.Result.Item(ssRst.ActiveRow).TestCd
'    strTmp = UCase(ssRst.Value)
'    strUTmp = ssRst.Value
'
'    ssRst.Col = objPtInfo.SSCol("MAXCOL"): strTmp1 = ssRst.Value
'    strRstVal = Trim(medGetP(objPtInfo.GetRstCdValString(strTestCd, strTmp1), 1, COL_DIV))
'
'    If strTmp = strRstVal Or strUTmp = strRstVal Then
'        blnRstChange = True
'        Exit Sub
'    End If
'
'    strResultVal = objPtInfo.GetRstCdValString(strTestCd, strTmp)
'    strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
'    strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
'
'    If strTmp <> strResultVal Then
'    '결과코드값이 있다.
'        ssRst.Col = objPtInfo.SSCol("RESULT"): ssRst.Value = strResultVal
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"): ssRst.Value = strTmp
'        If strResultChk <> "" Then
'            objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = ""
'            objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = ""
'            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = ""
'            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = ""
'            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = ""
'        End If
'
'        Select Case strResultChk
'            Case "*"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = "N"
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
''                    objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = "N"
''                    ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
''                                                            ssRst.FontBold = True
''                                                            ssRst.ForeColor = DCM_LightBlue
''                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
''                                                            ssRst.FontBold = True
''                                                            ssRst.ForeColor = DCM_LightBlue
'            Case "L"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = strResultChk
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'            Case "H"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = strResultChk
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'        End Select
'        blnRstChange = True
'    Else
'    '결과코드값이 없다
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"):  ssRst.Value = strTmp
'        ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = ""
'        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = ""
'        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = ""
'        objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = ""
'        objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = ""
'    End If
'End Sub

Private Sub ssRst_Advance(ByVal AdvanceNext As Boolean)
    Dim strCodeValue    As String
    Dim strRstType      As String
    Dim strErr          As String
    Dim strTestCd       As String
    Dim strResultVal    As String
    Dim strResultChk    As String
    Dim lngMaxCol       As String
    Dim lngResultCol    As String
    
    Dim Col             As Long
    Dim Row             As Long
   '
    Row = ssRst.ActiveRow
    If Row < 0 Then Exit Sub
    lngResultCol = objPtInfo.SSCol("RESULT")
    lngMaxCol = objPtInfo.SSCol("MAXCOL")
    
    On Error GoTo ErrLevaeCell:
    '
    Col = ssRst.ActiveCol
    If Col = lngResultCol Then
        objPtInfo.ResultCheck
        strRstType = objPtInfo.Result.Item(Row).RstType
        If strRstType = "N" Then
            strErr = objPtInfo.Result.Item(Row).AvalVal
            If objPtInfo.IsAvalVal = False Then
                If strErr <> "0" Then
                    strErr = "유효숫자 입력 오류. (" & objPtInfo.Result.Item(Row).AvalVal & "자리)"
                Else
                    strErr = "유효숫자 입력 오류. (정수형만 입력)"
                End If
                GoTo ErrLevaeCell
            Else
                Call objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            If objPtInfo.IsAlphaCd = False Then
                strErr = "ALPHA 결과코드 입력 오류!"
                GoTo ErrLevaeCell
            End If
        ElseIf strRstType = "R" Then
            If objPtInfo.IsRateCd = False Then
                strErr = "비율결과 입력 오류!"
                GoTo ErrLevaeCell
            Else
               lblErr.Caption = ""
            End If
        ElseIf strRstType = "F" Then
            If objPtInfo.IsFreeResult = False Then
                strErr = "FREE결과 입력 오류! (10자리이내)"
                GoTo ErrLevaeCell
            Else
                Call objPtInfo.NumValCheck
                lblErr.Caption = ""
            End If
        End If
    End If
        
    strTestCd = objPtInfo.Result.Item(Row).TestCd

    If Col = lngResultCol Then
        ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue = "" Then
            ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        End If
'        ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue <> "" Then
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
        
            If strResultVal <> ssRst.Value Then
                ssRst.Row = Row: ssRst.Col = lngResultCol:  ssRst.Value = strResultVal
                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDiv"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
            End If
            
        Else
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
            
            If strResultVal <> strCodeValue Then
                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDiv"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                If strRstType = "F" Then
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                        ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                    Else
                        ssRst.Col = lngResultCol:   ssRst.Value = ""
                        ssRst.Col = lngMaxCol:      ssRst.Value = ""
                    End If
                Else
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                End If
            End If
        End If
    End If
    
    LeaveCellFg = False
    Exit Sub
   '
ErrLevaeCell:
    With ssRst
        .Row = Row: .Col = objPtInfo.SSCol("RESULT"): .Value = ""
    End With
    objPtInfo.ResultCheck
    
    MsgFg = True
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    MsgFg = False
    
    On Error Resume Next
    ssRst.SetFocus
End Sub

Private Sub ssRst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim strCodeValue    As String       '입력값
    Dim strRstType      As String       '결과타입
    Dim strErr          As String       '에러메세지
    Dim strTestCd       As String       '결과등록 검사코드
    Dim strResultVal    As String       '결과값
    Dim strResultChk    As String       '결과코드입력값 체크
    Dim lngResultCol    As Long         '결과입력 Col
    Dim lngMaxCol       As Long         '결과저장 Col
    
    strResultVal = "": strResultChk = ""
    lngMaxCol = objPtInfo.SSCol("MAXCOL")
    lngResultCol = objPtInfo.SSCol("RESULT")
    
    If Row < 1 Then Exit Sub
    If MsgFg Then Exit Sub

    On Error GoTo ErrLevaeCell
    
    If NewRow > 0 Then Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(NewRow).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
    If Row = ssRst.MaxRows Then
        'Advance 이벤트에서 포커스가 스프레드에서 다른컨트롤로 넘어갈시
        'LeaveCell이벤트의 뼁뼁이를 방지하기 위해서 exit sub를 씀
        '허나, ESR이 아닌 다른 아이템에 대해서는 항목이 하나일때 EXIT SUb를 빼면
        '참고치 체크가 안된다.
        blnRstChange = False
        If lngResultCol <> Col Then blnRstChange = True
        If blnRstChange = True Then Exit Sub
'        If lngResultCol = Col Then Call ssRst_LostFocus
'
'        If UCase(Me.ActiveControl.Name) = "SSRST" Then Exit Sub
        If blnRstChange = True Then Exit Sub
'    Else
'        If lngResultCol <> Col Then blnRstChange = True
'        If lngResultCol = Col Then Call ssRst_LostFocus
'        If blnRstChange = True Then Exit Sub
    End If
    
    lblErr.Caption = ""
    If Col = lngResultCol Then
        Call objPtInfo.ResultCheck
        strRstType = objPtInfo.Result.Item(Row).RstType
        If strRstType = "N" Then
            strErr = objPtInfo.Result.Item(Row).AvalVal
            If objPtInfo.IsAvalVal = False Then
                If strErr <> "0" Then
                    strErr = "유효숫자 입력 오류. (" & objPtInfo.Result.Item(Row).AvalVal & "자리)"
                Else
                    strErr = "유효숫자 입력 오류. (정수형만 입력)"
                End If
                GoTo ErrLevaeCell
            Else
                objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            If objPtInfo.IsAlphaCd = False Then
               strErr = "결과 입력 오류!"
               GoTo ErrLevaeCell
            End If
        ElseIf strRstType = "R" Then
            If objPtInfo.IsRateCd = False Then
               strErr = "비율결과 입력 오류!"
               GoTo ErrLevaeCell
            End If
        ElseIf strRstType = "F" Then
            If objPtInfo.IsFreeResult = False Then
               strErr = "FREE결과 입력 오류! (10자리이내)"
               GoTo ErrLevaeCell
            End If
            objPtInfo.NumValCheck
        End If
        ssRst.EditEnterAction = EditEnterActionDown
    End If
   '
    Call SpDispRtfText(NewRow)
    
    strTestCd = objPtInfo.Result.Item(Row).TestCd
    If Col = lngResultCol Then
        ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue = "" Then
            ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        End If
'        ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue <> "" Then
            '저장 Col에 값이 있을경우(popup이용)
'            ssRst.Col = lngMaxCol:          ssRst.Value = strCodeValue
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)       '결과값
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))          '결과체크값
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))          '결과값

            ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
            ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'            If strResultChk <> "" Then
'                objPtInfo.Result.Item(Row).DPDiv = ""
'                objPtInfo.Result.Item(Row).HLDiv = ""
'            End If
            Select Case strResultChk
                Case "*"
                        objPtInfo.Result.Item(Row).HLDiv = "N"
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                Case "L"
                        objPtInfo.Result.Item(Row).HLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                Case "H"
                        objPtInfo.Result.Item(Row).HLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
            End Select
'            If strResultVal <> ssRst.Value Then
'                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
'                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
'                Select Case strResultChk
'                    Case "*"
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                    Case "L"
'                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                    Case "H"
'                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightRed
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightRed
'                End Select
'            Else
'                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
'            End If
        Else
            '저장Col에 값이 없을경우(직접입력)
            ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)       '결과값
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))          '결과체크값
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))          '결과값
            If strResultVal <> strCodeValue Then
                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).DPDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = strResultChk
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = strResultChk
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "▼Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High▲"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                If strRstType = "F" Then
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                        ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                    Else
                        ssRst.Col = lngResultCol:   ssRst.Value = ""
                        ssRst.Col = lngMaxCol:      ssRst.Value = ""
                    End If
                Else
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                End If
            End If
        End If
    End If
    
    LeaveCellFg = False
    Exit Sub
   '
ErrLevaeCell:
    With ssRst
        .Row = Row: .Col = objPtInfo.SSCol("RESULT"): .Value = ""
    End With
    objPtInfo.ResultCheck
    
    MsgFg = True
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    MsgFg = False
    
    LeaveCellFg = True
    
    Cancel = True
    
    On Error Resume Next
    ssRst.SetFocus
   '
End Sub

Private Sub ssRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
   '
    If Row < 1 Then Exit Sub
    
    ssRst.Row = Row
    ssRst.Col = 11
    objPtInfo.SpToolTip Row, Col, MultiLine, ShowTip, TipText, TipWidth
    ssRst.TextTip = TextTipFloatingFocusOnly
   '
End Sub

Private Sub SpDispRtfText(Optional Row As Long = 0)
   '
    If Row < 0 Then Exit Sub
    If Row = 0 Then
        ssRst.Row = ssRst.ActiveRow
    Else
        ssRst.Row = Row
    End If
    
    ssRst.Col = objPtInfo.SSCol("TXT")
    With objPtInfo.Result.Item(ssRst.Row)
        If ssRst.CellType = CellTypePicture Or ssRst.Text = "T" Then
            If .TxtType <> "0" Then
                rtfText.Text = .TextRst
                rtfText.Enabled = True
                cmdTextTemplete.Enabled = True
                rtfText.BackColor = &HEEFFFE    'vbWhite
            Else
                rtfText.Text = ""
                rtfText.Enabled = False
                cmdTextTemplete.Enabled = False
                rtfText.BackColor = DCM_LightGray
            End If
        Else
            rtfText.Text = ""
            rtfText.Enabled = False
            cmdTextTemplete.Enabled = False
            rtfText.BackColor = DCM_LightGray
        End If
    End With
   '
End Sub

Private Sub CmdTemplete(ByVal blnVisible As Boolean)
   '
    cmdTextTemplete.Enabled = blnVisible
    cmdRemarkTemplete.Enabled = blnVisible
    cmdCommentTemplete.Enabled = blnVisible
   '
End Sub

Private Sub cmdRmk_Click()
    Dim objSQL   As clsLISResultReview
    Dim RS       As Recordset
    Dim aryTmp() As String
    Dim strTmp   As String
    Dim SSQL     As String
    Dim ii       As Integer
    
    txtMesg.Text = ""
    Set objSQL = New clsLISResultReview
    SSQL = objSQL.GetOrderRemark(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    
    If Not RS.EOF Then
        strTmp = "처방일자 : " & Format(RS.Fields("orddt").Value & "", "####-##-##") & vbCRLF
        strTmp = strTmp & "처방번호 : " & RS.Fields("ordno").Value & "" & vbCRLF
        strTmp = strTmp & "처방비고  " & vbCRLF
        txtMesg.Text = strTmp
        strTmp = ""
        aryTmp = Split(RS.Fields("mesg").Value & "", vbCRLF)
        For ii = LBound(aryTmp) To UBound(aryTmp)
            strTmp = strTmp & " " & aryTmp(ii) & vbCRLF
        Next
        txtMesg.Text = txtMesg.Text & strTmp
        fraMesg.Visible = True
        fraMesg.ZOrder 0
    End If
    
    Set RS = Nothing
    Set objSQL = Nothing
End Sub

Private Sub cmdOk_Click()
    fraMesg.Visible = False
    ssRst.SetFocus
End Sub
