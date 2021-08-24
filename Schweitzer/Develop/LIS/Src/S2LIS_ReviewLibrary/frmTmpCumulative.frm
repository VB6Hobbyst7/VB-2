VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTmpCumulative 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "누적결과 조회"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14520
   Icon            =   "frmTmpCumulative.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   14520
   Begin FPSpread.vaSpread tblResult 
      Height          =   4140
      Left            =   15
      TabIndex        =   28
      Top             =   540
      Width           =   14445
      _Version        =   196608
      _ExtentX        =   25479
      _ExtentY        =   7303
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
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
      GrayAreaBackColor=   14411494
      MaxCols         =   17
      MaxRows         =   50
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14870761
      ShadowDark      =   14870761
      SpreadDesigner  =   "frmTmpCumulative.frx":144A
      TextTip         =   4
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   315
      Left            =   8070
      TabIndex        =   42
      Top             =   90
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
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
      Format          =   64028672
      CurrentDate     =   36567
   End
   Begin MSComCtl2.DTPicker dtpFromDt 
      Height          =   315
      Left            =   5115
      TabIndex        =   32
      Top             =   90
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   556
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
      Format          =   64028672
      CurrentDate     =   36567
   End
   Begin VB.ListBox lstCumList 
      BackColor       =   &H00EEF4F4&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1140
      Left            =   3060
      TabIndex        =   39
      Top             =   360
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00FFF9F4&
      Caption         =   "조회(&Q)"
      Height          =   450
      Left            =   12000
      Style           =   1  '그래픽
      TabIndex        =   33
      Tag             =   "133"
      Top             =   45
      Width           =   1230
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   435
      Left            =   10770
      Style           =   1  '그래픽
      TabIndex        =   31
      Tag             =   "132"
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Clear"
      Height          =   435
      Left            =   13215
      Style           =   1  '그래픽
      TabIndex        =   30
      Tag             =   "124"
      Top             =   1665
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdItemAdd 
      BackColor       =   &H00F4F0F2&
      Caption         =   "항목추가"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4965
      Style           =   1  '그래픽
      TabIndex        =   23
      Top             =   1635
      Visible         =   0   'False
      Width           =   1080
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
      Left            =   1005
      MaxLength       =   10
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.TextBox txtCumCd 
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
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   19
      Top             =   45
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.OptionButton optCumCd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Default"
      Height          =   330
      Index           =   0
      Left            =   3060
      Style           =   1  '그래픽
      TabIndex        =   18
      Top             =   1650
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.OptionButton optCumCd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "과별"
      Height          =   330
      Index           =   1
      Left            =   3855
      Style           =   1  '그래픽
      TabIndex        =   17
      Top             =   1650
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   465
      Left            =   13215
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "132"
      Top             =   690
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MedControls1.LisLabel lblPtDob 
      Height          =   270
      Left            =   1110
      TabIndex        =   0
      Top             =   690
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   476
      BackColor       =   13359320
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
      Alignment       =   1
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblPtNm 
      Height          =   270
      Left            =   1005
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   476
      BackColor       =   13359320
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
      Alignment       =   1
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin RichTextLib.RichTextBox txtSamCmt 
      Height          =   735
      Left            =   7005
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   15658734
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmTmpCumulative.frx":21CB
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
   Begin MedControls1.LisLabel lblDeptNm 
      Height          =   270
      Left            =   1005
      TabIndex        =   40
      Top             =   1365
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   476
      BackColor       =   13359320
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
      Alignment       =   1
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblWardId 
      Height          =   270
      Left            =   1005
      TabIndex        =   41
      Top             =   1695
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   476
      BackColor       =   13359320
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
      Alignment       =   1
      Caption         =   ""
      Appearance      =   0
      LeftGab         =   100
   End
   Begin VB.CheckBox chkGraph 
      BackColor       =   &H00DBE6E6&
      Caption         =   "그래프(&G)"
      ForeColor       =   &H00475765&
      Height          =   270
      Left            =   45
      TabIndex        =   37
      Tag             =   "40201"
      Top             =   180
      Width           =   1260
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E4F3F8&
      Caption         =   "<< (&P)"
      Height          =   465
      Index           =   0
      Left            =   1305
      Style           =   1  '그래픽
      TabIndex        =   36
      Top             =   60
      Width           =   1380
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E4F3F8&
      Caption         =   "(&N) >>"
      Height          =   465
      Index           =   1
      Left            =   2730
      Style           =   1  '그래픽
      TabIndex        =   35
      Top             =   60
      Width           =   1380
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   435
      Left            =   13260
      Style           =   1  '그래픽
      TabIndex        =   29
      Tag             =   "128"
      Top             =   45
      Width           =   1215
   End
   Begin VB.CommandButton cmdCumItem 
      BackColor       =   &H00F4F0F2&
      Caption         =   "누적코드등록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4680
      Style           =   1  '그래픽
      TabIndex        =   27
      Top             =   60
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.ListBox lstDtTm 
      Height          =   2580
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   2100
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.ListBox lstRemark 
      Height          =   2580
      Left            =   2685
      Sorted          =   -1  'True
      TabIndex        =   24
      Top             =   2100
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.CommandButton cmdPrintGraph 
      BackColor       =   &H00DBF2FD&
      Caption         =   "Print"
      Height          =   315
      Left            =   13605
      Style           =   1  '그래픽
      TabIndex        =   21
      Top             =   4695
      Width           =   885
   End
   Begin VB.ListBox lstSpcList 
      BackColor       =   &H00F1F5F4&
      Height          =   2040
      Left            =   6315
      TabIndex        =   4
      Top             =   2010
      Visible         =   0   'False
      Width           =   3480
   End
   Begin FPSpread.vaSpread tblExcel 
      Height          =   750
      Left            =   7215
      TabIndex        =   2
      Top             =   4695
      Visible         =   0   'False
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   1323
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
      SpreadDesigner  =   "frmTmpCumulative.frx":2270
   End
   Begin MSComDlg.CommonDialog cmdDlg 
      Left            =   7020
      Top             =   2145
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MedControls1.LisLabel lblMsg 
      Height          =   495
      Left            =   3870
      TabIndex        =   22
      Top             =   1425
      Visible         =   0   'False
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   873
      BackColor       =   16252927
      ForeColor       =   14641726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "누적코드를 로딩중입니다. 잠시만 기다리세요...."
      Appearance      =   0
      LeftGab         =   0
   End
   Begin ChartfxLibCtl.ChartFX grpResult 
      Height          =   2145
      Left            =   15
      TabIndex        =   25
      Top             =   4695
      Width           =   14460
      _cx             =   1710384034
      _cy             =   1710362312
      Build           =   7
      TypeMask        =   -1884749823
      Style           =   -67125249
      LeftGap         =   60
      RightGap        =   50
      TopGap          =   40
      BottomGap       =   31
      WallWidth       =   8
      View3DDepth     =   60
      TypeEx          =   32
      StyleEx         =   0
      DblClk          =   0
      RigClk          =   0
      MarkerShape     =   5
      MarkerSize      =   2
      Axis(0).MinorStep=   -2
      Axis(0).Max     =   6
      Axis(0).Decimals=   1
      Axis(0).TickMark=   -32767
      Axis(1).Min     =   0
      Axis(1).Max     =   100
      Axis(1).Decimals=   0
      Axis(1).Style   =   10344
      Axis(1).GridColor=   0
      Axis(2).Step    =   1
      Axis(2).MinorStep=   1
      Axis(2).Min     =   0
      Axis(2).Max     =   100
      Axis(2).Style   =   14368
      Axis(2).PixPerUnit=   0
      RGBBk           =   14870761
      RGB2DBk         =   16777215
      RGB3DBk         =   14870761
      nColors         =   1
      Colors          =   "frmTmpCumulative.frx":25BA
      TopFontMask     =   268435456
      BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BottomFontMask  =   268435456
      BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PointFontMask   =   268435456
      BeginProperty PointFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      nPts            =   25
      nSer            =   1
      NumPoint        =   25
      NumSer          =   1
      _Data_          =   "frmTmpCumulative.frx":25E2
   End
   Begin DRcontrol1.DrFrame fraTextResult 
      Height          =   8040
      Left            =   1980
      TabIndex        =   43
      Top             =   1080
      Visible         =   0   'False
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   14182
      Title           =   ""
      DelLine         =   0
      BackColor       =   13753559
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         Height          =   270
         Left            =   8280
         TabIndex        =   44
         Top             =   135
         Width           =   285
      End
      Begin RichTextLib.RichTextBox txtSamCmt1 
         Height          =   2265
         Left            =   -1275
         TabIndex        =   45
         Top             =   5460
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   3995
         _Version        =   393217
         BackColor       =   16252927
         TextRTF         =   $"frmTmpCumulative.frx":277B
      End
      Begin RichTextLib.RichTextBox txtRstCmt1 
         Height          =   4815
         Left            =   165
         TabIndex        =   46
         Top             =   3000
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   8493
         _Version        =   393217
         BackColor       =   16710910
         TextRTF         =   $"frmTmpCumulative.frx":2818
         MouseIcon       =   "frmTmpCumulative.frx":28B5
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   225
         Left            =   0
         TabIndex        =   64
         Tag             =   "40110"
         Top             =   0
         Width           =   450
      End
      Begin VB.Label lblRstCmt1 
         BackStyle       =   0  '투명
         Caption         =   "Result Comment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   225
         TabIndex        =   48
         Tag             =   "40204"
         Top             =   2790
         Width           =   2205
      End
      Begin VB.Label lblSamCmt1 
         BackStyle       =   0  '투명
         Caption         =   "Sample Comment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   47
         Tag             =   "40205"
         Top             =   180
         Width           =   2370
      End
   End
   Begin RichTextLib.RichTextBox txtRstCmt 
      Height          =   1515
      Left            =   3975
      TabIndex        =   38
      Top             =   3780
      Visible         =   0   'False
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   2672
      _Version        =   393217
      BackColor       =   16252927
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmTmpCumulative.frx":2A17
      MouseIcon       =   "frmTmpCumulative.frx":2ABC
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
   Begin DRcontrol1.DrFrame fraAddItem 
      Height          =   5355
      Left            =   5490
      TabIndex        =   5
      Top             =   750
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   9446
      Title           =   "검사항목 추가"
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdSpcList 
         BackColor       =   &H00D1DCD7&
         Caption         =   "▼"
         Height          =   300
         Left            =   1800
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   945
         Width           =   285
      End
      Begin VB.TextBox txtSpcCd 
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
         Height          =   315
         Left            =   825
         TabIndex        =   12
         Top             =   930
         Width           =   945
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00F4F0F2&
         Caption         =   "지움"
         Height          =   345
         Left            =   2010
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   4830
         Width           =   750
      End
      Begin VB.ListBox lstItemList 
         BackColor       =   &H00F4FEED&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   165
         Style           =   1  '확인란
         TabIndex        =   10
         Top             =   1335
         Width           =   4140
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F4F0F2&
         Caption         =   "취소"
         Height          =   345
         Left            =   3585
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   4830
         Width           =   765
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00F4F0F2&
         Caption         =   "확인"
         Height          =   345
         Left            =   2790
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   4830
         Width           =   765
      End
      Begin VB.ListBox lstSelList 
         BackColor       =   &H00EBEBEB&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   165
         TabIndex        =   7
         Top             =   2670
         Width           =   4140
      End
      Begin VB.ComboBox cboWorkArea 
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
         Left            =   840
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   570
         Width           =   3480
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   2100
         TabIndex        =   14
         Top             =   930
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         BackColor       =   14737632
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
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 체"
         Height          =   180
         Left            =   255
         TabIndex        =   16
         Tag             =   "40202"
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "Work Area"
         Height          =   330
         Left            =   255
         TabIndex        =   15
         Tag             =   "40202"
         Top             =   540
         Width           =   465
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblPtId 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "환자  I D : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   105
      TabIndex        =   63
      Tag             =   "105"
      Top             =   90
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblName 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "성      명 : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   90
      TabIndex        =   62
      Tag             =   "103"
      Top             =   435
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblSexAge 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "성별/연령 : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   0
      TabIndex        =   61
      Tag             =   "108"
      Top             =   750
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lblPtSex 
      BackStyle       =   0  '투명
      Height          =   180
      Left            =   1260
      TabIndex        =   60
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblPtAge 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   180
      Left            =   2355
      TabIndex        =   59
      Top             =   750
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label lblDOB 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "생년월일 : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   90
      TabIndex        =   58
      Tag             =   "101"
      Top             =   1095
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblRptNm 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "누적코드 :"
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   3105
      TabIndex        =   57
      Tag             =   "40202"
      Top             =   75
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblSamCmt 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "Remark : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   6180
      TabIndex        =   56
      Tag             =   "40205"
      Top             =   450
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblRstCmt 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사소견 : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   6105
      TabIndex        =   55
      Tag             =   "40204"
      Top             =   1200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblFrDt 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "시 작 일 : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   6150
      TabIndex        =   54
      Tag             =   "154"
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblLocation 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "병      실 : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   90
      TabIndex        =   53
      Tag             =   "102"
      Top             =   1755
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "진료부서 : "
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   90
      TabIndex        =   52
      Tag             =   "102"
      Top             =   1425
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblAgeDiv 
      Alignment       =   1  '오른쪽 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   180
      Left            =   2670
      TabIndex        =   51
      Top             =   750
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      Height          =   285
      Left            =   990
      Shape           =   4  '둥근 사각형
      Top             =   15
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   285
      Left            =   4065
      Shape           =   4  '둥근 사각형
      Top             =   30
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Label Label8 
      Appearance      =   0  '평면
      BackColor       =   &H00CBD8D8&
      Caption         =   "               /"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   960
      TabIndex        =   49
      Top             =   1065
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "~"
      Height          =   180
      Left            =   7815
      TabIndex        =   50
      Tag             =   "40110"
      Top             =   150
      Width           =   135
   End
End
Attribute VB_Name = "frmTmpCumulative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------
'누적결과 조회
'--------------------------

Public Event Click()


Private MyPatient As New clsPatient     '환자 클래스
Private MySql As New clsLISSqlStatement      'Sql문 클래스
Private objRstSQL As New clsLISSqlReview       'Sql문 클래스
Private objResult As New clsLISResultReview

Private Type MyItem
    TestCd As String
    PanelFg As String
    TestDiv As String
    SpcCd As String
    WorkArea As String
    TestNm As String
    SpcNm As String
    RefVal As String
End Type

Private MyItem() As MyItem
Private mItemCount As Integer
Private iPageNo As Integer
Private iPageCnt As Integer
Private OldRow As Integer
Private OldColor As Long

Const iColPerPage = 10

Private mCumCol As Collection
Private mDeptCd As String

Public PtFg As Boolean
Public QueryFg As Boolean

Private ClearFg As Boolean
Private blnNewFg  As Boolean
Private blnChanged As Boolean
Private MsgFg As Boolean


'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuCopy As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_COPY& = 1

Public Property Get DeptCd() As String
    DeptCd = mDeptCd
End Property

Public Property Let DeptCd(ByVal vNewValue As String)
    mDeptCd = vNewValue
End Property


Private Sub cboWorkArea_Click()
    Call objResult.LoadItem(lstItemList, medGetP(cboWorkArea.Text, 1, " "), txtSpcCd.Text)
End Sub

Private Sub chkGraph_Click()
    If chkGraph.Value = 1 Then
        'If Not grpResult.Visible Then
            grpResult.Visible = True
            cmdPrintGraph.Visible = True
            tblResult.Height = grpResult.Top - tblResult.Top
            If OldRow > 0 Then
                Call tblResult_Click(2, OldRow)
                tblResult.TopRow = OldRow
            Else
                Call tblResult_Click(2, 1)
                tblResult.TopRow = 1
            End If
        'End If
    Else
        'If grpResult.Visible Then
            grpResult.Visible = False
            cmdPrintGraph.Visible = False
            tblResult.Height = grpResult.Top - tblResult.Top + grpResult.Height + 20
        'End If
    End If
End Sub

Private Sub cmdCancel_Click()
    txtCumCd.SetFocus
    fraAddItem.Visible = False
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn
    'txtPtId.Text = ""
    txtCumCd.Text = ""
    Call SetStartDt
'    txtPtID.SetFocus
End Sub

Private Sub cmdClose_Click()
    fraTextResult.Visible = False
End Sub

Private Sub cmdCumItem_Click()
    lblMsg.Caption = "누적코드를 로딩중입니다. 잠시만 기다리세요...."
    lblMsg.Visible = True
    lblMsg.ZOrder 0
    DoEvents
    frm4021CumCdSet.DeptCd = gDeptCd    'mDeptCd
    frm4021CumCdSet.IsManager = False
    frm4021CumCdSet.Show 1
    lblMsg.Visible = False
End Sub

Private Sub cmdExcel_Click()
    Dim strPath     As String
    Dim strTmp      As String
    Dim strFileNm   As String
    
    If tblResult.DataRowCnt = 0 Then Exit Sub
    
    strFileNm = medGetP(lstCumList.List(lstCumList.ListIndex), 2, vbTab)
    
    With tblResult
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblExcel.MaxRows = .MaxRows + 1
        tblExcel.MaxCols = .MaxCols
        tblExcel.Row = 1: tblExcel.Row2 = tblExcel.MaxRows
        tblExcel.Col = 1: tblExcel.Col2 = tblExcel.MaxCols
        tblExcel.BlockMode = True
        tblExcel.Clip = strTmp
        tblExcel.BlockMode = False
    End With
    
    cmdDlg.InitDir = "C:\My Documents"
    cmdDlg.Filter = "ExCelFile(*.XLS)|*.XLS"
    cmdDlg.FileName = strFileNm & " 누적결과"
    cmdDlg.ShowSave
    

    tblExcel.SaveTabFile (cmdDlg.FileName)
End Sub

Private Sub cmdExit_Click()
    RaiseEvent Click
    Unload Me
End Sub

Private Sub cmdItemAdd_Click()
    

    lblMsg.Caption = "검사항목 리스트를 로드하고 있습니다. 잠시만 기다리세요...."
    lblMsg.Visible = True
    DoEvents
    cboWorkArea.ListIndex = -1
'    If mItemCount > 0 Then
'        txtSpcCd.Text = MyItem(mItemCount).SpcCd
'        tblResult.Row = 0: tblResult.Col = 2
'        lblSpcNm.Caption = tblResult.Text
'    Else
    txtSpcCd.Text = ""
    lblSpcNm.Caption = ""
'    End If
    If lstSpcList.ListCount = 0 Then Call LoadSpcList(lstSpcList)
    Call cmdReset_Click
    lstSelList.Clear
    lstSpcList.Visible = False
    fraAddItem.Visible = True
    fraAddItem.ZOrder 0
    lblMsg.Visible = False
End Sub

Private Sub cmdNext_Click(Index As Integer)
    
    Select Case Index
    Case 0:
        iPageNo = iPageNo - 1
        If iPageCnt > 1 Then cmdNext(1).Enabled = True
    Case 1:
        iPageNo = iPageNo + 1
        If iPageCnt > 1 Then cmdNext(0).Enabled = True
    End Select
    Call DisplayOnePage(iPageNo)
    If chkGraph.Value = 1 Then Call ShowGraph(OldRow)
    
    If iPageNo = 1 Then
        cmdNext(0).Enabled = False
        If iPageCnt > 1 Then cmdNext(1).Enabled = True
    End If
    If iPageNo = iPageCnt Then
        cmdNext(1).Enabled = False
        If iPageCnt > 1 Then cmdNext(0).Enabled = True
    End If
'    tblResult.SetFocus
    
End Sub

Private Sub cmdOK_Click()

    Dim I As Integer
    Dim SqlStmt As String
    Dim rsRef As Recordset
    Dim RefF As Double, RefT As Double
    Dim objMyRst As New clsLISSqlReview
    Dim strDt As String
       
    strDt = Format(Now, CS_DateDbFormat)
    
    tblResult.ReDraw = False
    
    For I = 0 To lstSelList.ListCount - 1
        
        mItemCount = mItemCount + 1
        ReDim Preserve MyItem(mItemCount)
        With MyItem(mItemCount)
            .TestCd = medGetP(lstSelList.List(I), 1, " ")
            .TestNm = Trim(Mid(medGetP(lstSelList.List(I), 1, vbTab), 10))
            .TestDiv = medGetP(lstSelList.List(I), 2, vbTab)
            .WorkArea = medGetP(lstSelList.List(I), 3, vbTab)
            .PanelFg = medGetP(lstSelList.List(I), 4, vbTab)
            .SpcCd = medGetP(lstSelList.List(I), 5, vbTab)
            .SpcNm = medGetP(lstSelList.List(I), 6, vbTab)
            SqlStmt = objMyRst.SqlGetReference(.TestCd, .SpcCd, strDt, "B", _
                                            DateDiff("y", Format(MyPatient.Dob, CS_DateMask), Now))
            Set rsRef = Nothing
            Set rsRef = New Recordset
            rsRef.Open SqlStmt, DBConn
            
            If rsRef.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
               SqlStmt = objMyRst.SqlGetReference(.TestCd, .SpcCd, strDt, MyPatient.Sex, _
                                            DateDiff("y", Format(MyPatient.Dob, CS_DateMask), Now))
                Set rsRef = Nothing
                Set rsRef = New Recordset
                rsRef.Open SqlStmt, DBConn
            End If
            If rsRef.EOF Then
               .RefVal = ""
            Else
               RefF = Val("" & rsRef.Fields("RefValFrom").Value)
               RefT = Val("" & rsRef.Fields("RefValTo").Value)
               .RefVal = Trim("" & rsRef.Fields("RefCd").Value)
               If RefF <> 0 Or RefT <> 0 Then .RefVal = RefF & " - " & RefT
            End If
            Set rsRef = Nothing
            
            tblResult.MaxRows = mItemCount
            tblResult.Row = mItemCount
            tblResult.Col = 1: tblResult.Value = .TestCd
            tblResult.Col = 2: tblResult.Value = .TestNm
            tblResult.Col = 3: tblResult.Value = .SpcCd
            tblResult.Col = 4: tblResult.Value = .SpcNm
            tblResult.Col = 16: tblResult.Value = .RefVal
        End With
    Next
    
    Call SetTable
    tblResult.ReDraw = True
    txtCumCd.SetFocus
    fraAddItem.Visible = False
                
End Sub


Private Sub cmdPrint_Click()

    With tblResult
        
        .PrintMarginTop = 100
        .PrintJobName = "누적결과레포트 출력"
        
        .PrintAbortMsg = "누적결과지를 출력중입니다. "

        .PrintOrientation = PrintOrientationLandscape
        If Printer.PaperSize = vbPRPSA4 Then
            .PrintMarginLeft = 1700
            .PrintMarginRight = 100
            .PrintMarginTop = 800
            .PrintMarginBottom = 800
        Else
            .PrintMarginTop = 300
            .PrintMarginBottom = 500
            .PrintMarginLeft = 250
            .PrintMarginRight = 100
        End If
        .PrintColor = False
        .PrintFirstPageNumber = 1
       
        .PrintHeader = "/n/n/l/fb1 " & "♧ 누적결과 - " & txtPtId.Text & "  " & lblPtNm.Caption & "   " & _
                                        lblPtSex.Caption & "/" & lblPtAge.Caption & " " & lblAgeDiv.Caption & " /c/fb1/n/n"
        
        .PrintFooter = "/c/p/fb1"
        
        .PrintGrid = False
        .PrintShadows = False
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintPageEnd = 2
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        '.PrintGrid = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll
         
        .Action = ActionPrint
        .GridSolid = True
    End With
    
    'If chkGraph.Value = 1 Then Call PrintGraph

End Sub

Private Sub cmdPrintGraph_Click()
    Call PrintGraph
End Sub

Private Sub cmdQuery_Click()
    
    'Dim objPrgBar As New clsprogress
    Dim objPrgBar As New jProgressBar.clsProgress
    
    With objPrgBar
        .Container = Me
        .Left = tblResult.Left
        .Top = tblResult.Top
        .Width = tblResult.Width
        .Height = 280
        .Message = lblPtNm.Caption & " 님의 결과내역을 검색중입니다..."
        .Max = 100
        .Value = 1
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = tblResult.Width
'        .XPos = tblResult.Left
'        .YPos = tblResult.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = lblPtNm.Caption & " 님의 결과내역을 검색중입니다..."
'        .Value = 1
    End With
    
'    With objPrgBar
'        .Mode = 0
'        .CaptionOn = False
'        .Msg = "해당 환자의 결과내역을 검색중입니다."
'        .Max = mItemCount * 10 + 1
'        .Min = 0
'        .Value = 0
'        .Visible = True
'    End With
    DoEvents
    
    MouseRunning
    
'    txtPtId.Locked = True
'    txtPtId.BackColor = &HE0E0E0
    txtCumCd.Locked = False
    txtCumCd.BackColor = vbWhite
    lstCumList.Enabled = True
    optCumCd(0).Enabled = True
    optCumCd(1).Enabled = True
    cmdItemAdd.Enabled = True
    dtpFromDt.Enabled = True
    cmdQuery.Enabled = True
    cmdPrint.Enabled = True
    cmdExcel.Enabled = True

    Call DisplayResult(objPrgBar)
    
    MouseDefault
    'objPrgBar.Visible = False
    Set objPrgBar = Nothing
    
    If lstDtTm.ListCount <= 0 Then
        MsgBox "해당 환자의 누적결과가 없습니다.", vbInformation, "메세지"
        Call cmdClear_Click
        Exit Sub
    End If
    
End Sub

Private Sub cmdSpcList_Click()

    lstSpcList.Visible = True
    lstSpcList.ZOrder 0

End Sub

Private Sub cmdReset_Click()
    
    Dim I As Integer
    
    For I = 0 To lstItemList.ListCount - 1
        lstItemList.Selected(I) = False
    Next
    lstSelList.Clear

End Sub

Private Sub dtpFromDt_Click()
    
    If dtpFromDt.Value > Now Then
        MsgBox "시작일이 현재날짜보다 큽니다. 다시 설정하십시오.", vbExclamation, "메세지"
        dtpFromDt.SetFocus
    End If
    
End Sub

Private Sub dtpFromDt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If cmdQuery.Enabled Then cmdQuery.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Trim(gPatientId) <> "" Then txtPtId.Text = gPatientId
On Error GoTo Err_Trap
    txtPtId.SetFocus
    If Trim(txtPtId.Text) <> "" Then SendKeys "{TAB}"
Err_Trap:

End Sub

Private Sub Form_Load()
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
'    Me.Show
'    DoEvents
    PtFg = False
    Call ClearRtn
    Call LoadCumList(lstCumList, "")
    Call objResult.LoadWorkArea(cboWorkArea)
    Call SetStartDt
    'Call GrpDraw
End Sub

Private Sub lstCumList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call lstCumList_MouseDown(1, 0, 0, 0)
        dtpFromDt.SetFocus
    End If
End Sub

Private Sub lstCumList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        tblResult.MaxRows = 0
        txtCumCd.Text = Trim(medGetP(lstCumList.Text, 3, vbTab))
        MouseRunning
'        Call DisplayItem(Trim(txtCumCd.Text))
        Call cmdQuery_Click
        MouseDefault
        blnNewFg = False
        blnChanged = False
    End If
    
End Sub

Private Sub lstItemList_Click()
    Dim I As Integer
    I = medListFind(lstSelList, lstItemList.Text)
    If lstItemList.Selected(lstItemList.ListIndex) Then
        If I < 0 Then
            If ItemCheck(medGetP(lstItemList.Text, 1, " ")) Then lstSelList.AddItem lstItemList.Text
        End If
    Else
        If I >= 0 Then lstSelList.RemoveItem (I)
    End If
End Sub

Private Function ItemCheck(ByVal pTestCd As String) As Boolean
    Dim I As Integer
    ItemCheck = True
    For I = 1 To mItemCount
        If MyItem(I).TestCd = pTestCd Then
            ItemCheck = False
            Exit For
        End If
    Next
End Function

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_COPY
            Dim I As Long
            Dim strClip As String
            
            With tblResult
                .Row = OldRow
                .Col = 2: strClip = .Value
                .Col = 4: strClip = strClip & " ; " & .Value & " : "
                For I = 5 To 14
                    .Row = OldRow
                    .Col = I
                    If I >= .SelBlockCol And I <= .SelBlockCol2 Then
                        If Trim(.Value) <> "" Then
                            strClip = strClip & .Value
                            .Row = 0
                            .Col = I: strClip = strClip & "(" & Mid(.Value, 4, 5) & ")" & Space(2)
                        End If
                    End If
                Next
                .Row = OldRow
                .Col = 15: strClip = strClip & Space(3) & "단위(" & .Value & ")"
                .Col = 16: strClip = strClip & Space(3) & "기준치(" & .Value & ")"
            End With
            
            Clipboard.Clear
            Clipboard.SetText strClip
    End Select
End Sub

'Private Sub mnuCopy_Click()
'
'    Dim I As Long
'    Dim strClip As String
'
'    With tblResult
'        .Row = OldRow
'        .Col = 2: strClip = .Value
'        .Col = 4: strClip = strClip & " ; " & .Value & " : "
'        For I = 5 To 14
'            .Row = OldRow
'            .Col = I
'            If I >= .SelBlockCol And I <= .SelBlockCol2 Then
'                If Trim(.Value) <> "" Then
'                    strClip = strClip & .Value
'                    .Row = 0
'                    .Col = I: strClip = strClip & "(" & Mid(.Value, 4, 5) & ")" & Space(2)
'                End If
'            End If
'        Next
'        .Row = OldRow
'        .Col = 15: strClip = strClip & Space(3) & "단위(" & .Value & ")"
'        .Col = 16: strClip = strClip & Space(3) & "기준치(" & .Value & ")"
'    End With
'
'    Clipboard.Clear
'    Clipboard.SetText strClip
'
'End Sub

Private Sub optCumCd_Click(Index As Integer)

    If Index = 0 Then
        Call LoadCumList(lstCumList, "")
    Else
        Call LoadCumList(lstCumList, "dept")   'mDeptCd)
        If lstCumList.ListCount <= 0 Then
            MsgBox "부서코드 " & Chr(34) & gDeptCd & Chr(34) & " 에 등록된 누적코드가 없습니다.", vbExclamation, "메세지"
            optCumCd(0).Value = True
        ElseIf lstCumList.ListCount = 1 Then
            lstCumList.ListIndex = 0
            Call lstCumList_MouseDown(1, 0, 0, 0)
            DoEvents
            Call cmdQuery_Click
            DoEvents
        End If
    End If

End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim I As Integer
    Dim sDPfg As String
    If Row = 0 Then Exit Sub
    If Row = OldRow Then GoTo Skip1
    With tblResult
        .ReDraw = False
        If OldRow > 0 Then
            .Col = 2: .Col2 = .MaxCols
            .Row = OldRow: .Row2 = OldRow
            .BlockMode = True
            '.FontSize = 10
            .FontBold = False
            .BackColor = OldColor
            .CellBorderType = 0
            .Action = ActionSetCellBorder
            .BlockMode = False
            .Col = 2: .BackColor = &HE2E8E9
            .Col = 4: .BackColor = &HEEF4F4  '&HF9FBFB     '&HE7EFEF
            .RowHeight(OldRow) = 12
            
            .Col = 17
            sDPfg = .Value
            For I = 1 To 10
                If medGetP(sDPfg, I, ":") <> "" Then
                    .Col = I + 4
                    .BackColor = &HC0FFFF     '&HFFF7FF
                End If
            Next
            
        End If
        .Row = Row: .Col = 1
        OldColor = .BackColor
        
        .Col = 2: .Col2 = .MaxCols
        .Row = Row:  .Row2 = Row
        .BlockMode = True
        '.FontSize = 11
        .FontBold = True
        .BackColor = &HC0FFFF
        .CellBorderColor = &H80
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderType = 16
        .Action = ActionSetCellBorder
        .BlockMode = False
        .RowHeight(Row) = 12
        OldRow = Row
        .ReDraw = True
    End With
Skip1:
    If chkGraph.Value = 1 Then Call ShowGraph(Row)
End Sub

Private Sub tblResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    Call tblResult_Click(Col, Row)
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_COPY, "CLIPBOARD로 복사"
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuCopy = frmControls.mnuSub
'    frmControls.mnuSub1.Visible = False
'    mnuCopy.Caption = "Clipboard로 복사"
'
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuCopy = Nothing
'    Unload frmControls
'    Set frmControls = Nothing
End Sub

Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    If Row = 0 Then Exit Sub
    If Col = 2 Or Col = 4 Or Col = 15 Then
        tblResult.Row = Row
        tblResult.Col = Col
        MultiLine = 1
        TipText = "  " & Trim(tblResult.Value)
        TipWidth = Len(TipText) * 150  '3000
        tblResult.TextTipDelay = 200
        'Call spdStat.SetTextTipAppearance("굴림", 9, False, False, &HEEFDF2, vbBlue)    '&H996666)
        Call tblResult.SetTextTipAppearance("Arial", 11, False, False, vbWhite, vbBlue)    '&H996666)
        ShowTip = True
    ElseIf Col >= 5 Then
        tblResult.Row = Row
        tblResult.Col = Col
        If Len(tblResult.Value) > 9 Then
            MultiLine = 1
            TipText = "  " & Trim(tblResult.Value)
            TipWidth = Len(TipText) * 150  '3000
            tblResult.TextTipDelay = 200
            'Call spdStat.SetTextTipAppearance("굴림", 9, False, False, &HEEFDF2, vbBlue)    '&H996666)
            Call tblResult.SetTextTipAppearance("Arial", 11, False, False, vbWhite, vbBlue)    '&H996666)
            ShowTip = True
        Else
            ShowTip = False
        End If
    End If
End Sub

Private Sub txtCumCd_Change()
    lstCumList.ListIndex = medListFind(lstCumList, txtCumCd.Text)
End Sub

Private Sub txtCumCd_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown And lstCumList.ListCount > 0 Then
        lstCumList.SetFocus
    End If

End Sub

Private Sub txtCumCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        If lstCumList.ListIndex >= 0 Then
            Call lstCumList_MouseDown(1, 0, 0, 0)
            dtpFromDt.SetFocus
        Else
            txtCumCd.SetFocus
        End If
        Exit Sub
    End If
    If lstCumList.ListCount > 0 Then
        Call medCodeHelp(KeyAscii, lstCumList, txtCumCd.Text, txtCumCd, dtpFromDt)
    End If
End Sub

Private Sub txtCumCd_GotFocus()
   With txtCumCd
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

'% 환자ID가 변경되면 화면Clear
Private Sub txtPtId_Change()
    If Not ClearFg Then
        lblPtNm.Caption = ""
        lblPtSex.Caption = ""
        lblPtAge.Caption = ""
        lblAgeDiv.Caption = ""
        lblPtDob.Caption = ""
        lblDeptNm.Caption = ""
        lblWardId.Caption = ""
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

   If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub


Private Sub txtPtId_LostFocus()
    
    If Not gUsingInWardMenu Then
'        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'        If ActiveControl.Name = cmdExit.Name Then Exit Sub
'        If ActiveControl.Name = cmdClear.Name Then Exit Sub
'        If ActiveControl.Name = cmdCumItem.Name Then Exit Sub
'        If MsgFg Then Exit Sub
        
'        If txtPtId.Text = "" Then
'            txtPtId.SetFocus
'            Exit Sub
'        End If
    
    End If
    
    If txtPtId.Text = "" Then
        'txtPtId.SetFocus
        Exit Sub
    End If

    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    With MyPatient
        If .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm
            lblPtSex.Caption = .SEXNM
            lblPtAge.Caption = .Age
            lblAgeDiv.Caption = .AGEDIV
            lblDeptNm.Caption = .DeptNm
            lblPtDob.Caption = Format(.Dob, CS_DateMask)
            'lblBedinDt.Caption = Format(.BedInDt, CS_DateMask)
            'lblBedoutDt.Caption = Format(.BedOutDt, CS_DateMask)
            If .BEDOUTDT <> "" Then
                Dim strTmp1 As String
'            '최근의 처방과를 가지고 온다.
                strTmp1 = objRstSQL.GetDeptInfo(txtPtId.Text)
                If strTmp1 <> "" Then
                    lblLocation.Caption = ""
                    lblDeptNm.Caption = medGetP(strTmp1, 1, COL_DIV)
                    'lblDoctNm.Caption = medGetP(strTmp1, 2, COL_DIV)
                End If
            End If
            cmdItemAdd.Enabled = True
            gPatientId = txtPtId.Text
            ClearFg = False
            PtFg = True
        Else
            MsgFg = True
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
            txtPtId.SetFocus
            PtFg = False
            MsgFg = False
            Call txtPtId_GotFocus
            Exit Sub
        End If
    End With
    'If ActiveControl.Name <> cmdRefresh.Name Then dtpFromDate.SetFocus
End Sub


Private Sub txtRstCmt_DblClick()
   fraTextResult.Top = (Me.Height - fraTextResult.Height) / 2
   fraTextResult.Left = (Me.Width - fraTextResult.Width) / 2
   txtSamCmt1.Text = txtSamCmt.Text
   txtRstCmt1.Text = txtRstCmt.Text
   fraTextResult.Visible = True
   fraTextResult.ZOrder 0
End Sub

Private Sub txtRstCmt_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Name = "txtRstCmt" Then
        txtSamCmt.Height = txtSamCmt.Height + Y
        txtRstCmt.Height = txtRstCmt.Height - Y
        txtRstCmt.Top = txtRstCmt.Top + Y
    End If
    txtRstCmt.DragMode = 0
End Sub

Private Sub txtRstCmt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Y <= 50 Then
      txtRstCmt.DragMode = 1
      txtRstCmt.Drag vbBeginDrag  '1
   Else
      txtRstCmt.DragMode = 0
   End If
End Sub

Private Sub txtRstCmt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Y <= 30 Then
      txtRstCmt.MousePointer = 99
   Else
      txtRstCmt.MousePointer = rtfDefault
   End If
End Sub

Private Sub txtRstCmt1_DblClick()
   fraTextResult.Visible = False
End Sub

Private Sub txtSamCmt_DblClick()
   Call txtRstCmt_DblClick
End Sub

Private Sub txtSamCmt_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Name = "txtRstCmt" Then
        txtRstCmt.Height = txtRstCmt.Height + txtSamCmt.Height - Y
        txtSamCmt.Height = Y
        txtRstCmt.Top = txtSamCmt.Top + Y
    End If
    txtRstCmt.DragMode = 0
End Sub

Private Sub txtSamCmt1_DblClick()
   fraTextResult.Visible = False
End Sub



Public Sub DisplayItem(ByVal objTestCd As clsDictionary, ByVal strPtId As String)
    Dim RS          As Recordset
    Dim rsRef       As Recordset
    Dim objMyRst    As New clsLISSqlReview
    Dim blnDupChk   As Boolean
    Dim strDt       As String
    Dim SqlStmt     As String
    Dim RefF        As Double
    Dim RefT        As Double
    
    Dim I          As Integer
    
    strDt = Format(GetSystemDate, CS_DateDbFormat)
    
    txtPtId.Text = strPtId
    
    Call txtPtId_LostFocus
     
    Erase MyItem
    mItemCount = 0
    ReDim MyItem(mItemCount)
    
    tblResult.ReDraw = False
    objTestCd.MoveFirst
    Do Until objTestCd.EOF
        SqlStmt = MySql.GetCumulative(objTestCd.Fields("testcd"), objTestCd.Fields("spccd"))
        Set RS = Nothing
        Set RS = New Recordset
        RS.Open SqlStmt, DBConn
        
        blnDupChk = False
        For I = 1 To tblResult.DataRowCnt
            tblResult.Row = I
            tblResult.Col = 1
            If objTestCd.Fields("testcd") = tblResult.Value Then
                blnDupChk = True
                Exit For
            End If
        Next
        
        If Not blnDupChk Then
            mItemCount = mItemCount + 1
            If tblResult.DataRowCnt + 1 > tblResult.MaxRows Then
                tblResult.MaxRows = tblResult.MaxRows + 1
            End If
            tblResult.Row = tblResult.DataRowCnt + 1
            tblResult.Col = 1: tblResult.Value = "" & RS.Fields("TestCd").Value
            tblResult.Col = 2: tblResult.Value = "" & RS.Fields("TestNm").Value
            tblResult.Col = 3: tblResult.Value = "" & RS.Fields("SpcCd").Value
            tblResult.Col = 4: tblResult.Value = "" & RS.Fields("SpcNm").Value
            'tblResult.Col = 16: tblResult.Value = .RefVal
        Else
            GoTo Skip
        End If
        
        mItemCount = mItemCount + 1
        ReDim Preserve MyItem(mItemCount)
        With MyItem(mItemCount)
            .TestCd = Trim("" & RS.Fields("TestCd").Value)
            .PanelFg = Trim("" & RS.Fields("PanelFg").Value)
            .TestDiv = Trim("" & RS.Fields("TestDiv").Value)
            .SpcCd = Trim("" & RS.Fields("SpcCd").Value)
            .WorkArea = Trim("" & RS.Fields("WorkArea").Value)
            .TestNm = Trim("" & RS.Fields("TestNm").Value)
            .SpcNm = Trim("" & RS.Fields("SpcNm").Value)
            SqlStmt = objMyRst.SqlGetReference(.TestCd, .SpcCd, strDt, "B", _
                                            DateDiff("y", Format(MyPatient.Dob, CS_DateMask), GetSystemDate))
            Set rsRef = Nothing
            Set rsRef = New Recordset
            rsRef.Open SqlStmt, DBConn
            If rsRef.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
               SqlStmt = objMyRst.SqlGetReference(.TestCd, .SpcCd, strDt, MyPatient.Sex, _
                                            DateDiff("y", Format(MyPatient.Dob, CS_DateMask), Now))
                Set rsRef = Nothing
                Set rsRef = New Recordset
                rsRef.Open SqlStmt, DBConn
            End If
            If rsRef.EOF Then
               .RefVal = ""
            Else
               RefF = Val("" & rsRef.Fields("RefValFrom").Value)
               RefT = Val("" & rsRef.Fields("RefValTo").Value)
               .RefVal = Trim("" & rsRef.Fields("RefCd").Value)
               If RefF <> 0 Or RefT <> 0 Then .RefVal = RefF & " - " & RefT
            End If
            Set rsRef = Nothing
            
'            tblResult.MaxRows = mItemCount
'            tblResult.Row = mItemCount
'            tblResult.Col = 1: tblResult.Value = .TestCd
'            tblResult.Col = 2: tblResult.Value = .TestNm
'            tblResult.Col = 3: tblResult.Value = .SpcCd
'            tblResult.Col = 4: tblResult.Value = .SpcNm
'            tblResult.Col = 16: tblResult.Value = .RefVal
            
        End With
        If MyItem(mItemCount).PanelFg = PN_Detail Then _
            Call DisplayDetail(MyItem(mItemCount).TestCd, MyItem(mItemCount).SpcCd, MyItem(mItemCount).SpcNm)
                
                
        Set RS = Nothing
Skip:
        objTestCd.MoveNext
    Loop
   
    Call SetTable
    cmdQuery.Enabled = True
    
    Call cmdQuery_Click
    
NoData:
    tblResult.ReDraw = True
    
'    RS.Close
    Set RS = Nothing
    Set rsRef = Nothing
    Set objMyRst = Nothing
End Sub

Private Sub SetTable()

    With tblResult
        .Row = -1
        .Col = 2: .Col2 = 2
        .BlockMode = True
        .ForeColor = &H864B24
        .BackColor = &HE2E8E9
        .BlockMode = False
        .Col = 4: .Col2 = 4
        .BlockMode = True
        .ForeColor = &H808080     '&H404040        '&H864B24
        .BackColor = &HEEF4F4     '&HF9FBFB  '&HE7EFEF
        .BlockMode = False
        .Col = 15: .Col2 = 15
        .BlockMode = True
        .ForeColor = &H80 ' &HDF6A3E
        'tblResult.BackColor = &HFBFFFF
        .BlockMode = False
        .Col = 16: .Col2 = 16
        .BlockMode = True
        .ForeColor = &H136604 ' &HDF6A3E
        'tblResult.BackColor = &HFBFFFF
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
End Sub

Private Sub DisplayDetail(ByVal pTestCd As String, ByVal pSpcCd As String, ByVal pSpcNm As String)
    Dim RS          As Recordset
    Dim rsRef       As Recordset
    Dim objMyRst    As New clsLISSqlReview
    
    Dim SqlStmt     As String
    Dim strDt       As String
    Dim blnDupChk   As Boolean
    Dim RefF        As Double
    Dim RefT        As Double
    
    Dim I          As Integer
    
    strDt = Format(Now, CS_DateDbFormat)
    SqlStmt = objMyRst.SqlGetCumDetail(pTestCd)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    While (Not RS.EOF)
        blnDupChk = False
        For I = 1 To tblResult.DataRowCnt
            tblResult.Row = I
            tblResult.Col = 1
            If "" & RS.Fields("TestCd").Value = tblResult.Value Then
                blnDupChk = True
                Exit For
            End If
        Next
        
        If Not blnDupChk Then
            mItemCount = mItemCount + 1
            If tblResult.DataRowCnt + 1 > tblResult.MaxRows Then
                tblResult.MaxRows = tblResult.MaxRows + 1
            End If
            tblResult.Row = tblResult.DataRowCnt + 1
            tblResult.Col = 1: tblResult.Value = "" & RS.Fields("TestCd").Value
            tblResult.Col = 2: tblResult.Value = "    " & "" & RS.Fields("TestNm").Value
            tblResult.Col = 3: tblResult.Value = pSpcCd
            tblResult.Col = 4: tblResult.Value = pSpcNm
'            tblResult.Col = 16: tblResult.Value = .RefVal
        Else
            GoTo Skip
        End If
        
        'mItemCount = mItemCount + 1
        ReDim Preserve MyItem(mItemCount)
        With MyItem(mItemCount)
            .TestCd = Trim("" & RS.Fields("TestCd").Value)
            .PanelFg = Trim("" & RS.Fields("PanelFg").Value)
            .TestDiv = Trim("" & RS.Fields("TestDiv").Value)
            .SpcCd = pSpcCd
            .SpcNm = pSpcNm
            .WorkArea = Trim("" & RS.Fields("WorkArea").Value)
            .TestNm = "    " & Trim("" & RS.Fields("TestNm").Value)
            SqlStmt = objMyRst.SqlGetReference(.TestCd, .SpcCd, strDt, "B", _
                                            DateDiff("y", Format(MyPatient.Dob, CS_DateMask), Now))
            Set rsRef = Nothing
            Set rsRef = New Recordset
            rsRef.Open SqlStmt, DBConn
            If rsRef.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
               SqlStmt = objMyRst.SqlGetReference(.TestCd, .SpcCd, strDt, MyPatient.Sex, _
                                            DateDiff("y", Format(MyPatient.Dob, CS_DateMask), Now))
                Set rsRef = Nothing
                Set rsRef = New Recordset
                rsRef.Open SqlStmt, DBConn
            End If
            If rsRef.EOF Then
               .RefVal = ""
            Else
               RefF = Val("" & rsRef.Fields("RefValFrom").Value)
               RefT = Val("" & rsRef.Fields("RefValTo").Value)
               .RefVal = Trim("" & rsRef.Fields("RefCd").Value)
               If RefF <> 0 Or RefT <> 0 Then .RefVal = RefF & " - " & RefT
            End If
            Set rsRef = Nothing
            
'            tblResult.MaxRows = mItemCount
'            tblResult.Row = mItemCount
'            tblResult.Col = 1: tblResult.Value = .TestCd
'            tblResult.Col = 2: tblResult.Value = .TestNm
'            tblResult.Col = 3: tblResult.Value = .SpcCd
'            tblResult.Col = 4: tblResult.Value = .SpcNm
'            tblResult.Col = 16: tblResult.Value = .RefVal
            
        End With
Skip:
        RS.MoveNext
    Wend

    RS.Close
    Set RS = Nothing
    Set rsRef = Nothing
    Set objMyRst = Nothing
End Sub

Private Sub DisplayResult(ByRef objPrgBar As Object)

    iPageNo = 0
    iPageCnt = 0
    Call ReadData(objPrgBar)
    If lstDtTm.ListCount = 0 Then Exit Sub
    
    Call DisplayRemark
    Call DisplayTxtRst
    iPageCnt = (lstDtTm.ListCount + 9) \ 10
    Call cmdNext_Click(1)
    
End Sub

Private Sub DisplayRemark()

    Dim I As Integer
    Dim strDtTm As String, J As Integer
    
    txtSamCmt.Text = ""
    txtSamCmt1.Text = ""
    For I = 0 To lstRemark.ListCount - 1
        strDtTm = medGetP(lstRemark.List(I), 1, vbTab)
        J = Val(medGetP(lstRemark.List(I), 2, vbTab))
        txtSamCmt.Text = txtSamCmt.Text & "▶ " & strDtTm & vbCrLf
        If mCumCol.Item(J).RmkCd <> "" Then _
            txtSamCmt.Text = txtSamCmt.Text & "* Remark : " & mCumCol.Item(J).Remark & vbCrLf
        If mCumCol.Item(J).FootNoteFg <> "0" Then _
            txtSamCmt.Text = txtSamCmt.Text & mCumCol.Item(J).FootNote & vbCrLf
    Next
    
End Sub

Private Sub DisplayTxtRst()
    
    Dim I As Integer
    Dim J As Integer
    Dim clsData As clsCumResult
    
    On Error Resume Next
    txtRstCmt.Text = ""
    txtRstCmt1.Text = ""
    For I = 0 To lstDtTm.ListCount - 1
        For J = 1 To mItemCount
            Set clsData = mCumCol.Item(lstDtTm.List(I) & ":" & MyItem(J).TestCd)
            If clsData.RstText <> "" Then
                txtRstCmt.Text = txtRstCmt.Text & "▶ " & lstDtTm.List(I) & " : " & clsData.TestNm & vbCrLf
                txtRstCmt.Text = txtRstCmt.Text & clsData.RstText & vbCrLf
            End If
        Next
    Next
    Set clsData = Nothing

End Sub

Private Sub ReadData(ByRef barStatus As Object)

    Dim strFromDt As String, strToDt As String
    Dim strSpcCd As String, strWorkArea As String, strTestNm As String
    Dim strTestcd As String, strPanelFg As String, strTestDiv As String
    Dim strDtTm As String
    Dim iSeq  As Integer
    Dim strList As String
    
    Dim I As Integer, J As Integer
    Dim SqlStmt As String
    Dim RS As Recordset
    Dim ObjDic As New clsDictionary
    
    Dim clsNewData As clsCumResult
    Dim MyResult As New clsLISResultReview
    
    ObjDic.Clear
    ObjDic.FieldInialize "strDtTm, strTestCd, iSeq", "seq"
    
    strFromDt = Format(dtpFromDt.Value, CS_DateDbFormat)
    strToDt = Format(dtpToDt.Value, CS_DateDbFormat)
    
    lstDtTm.Clear
    lstRemark.Clear
    Set mCumCol = New Collection
    For I = 1 To mItemCount
    
        barStatus.Value = (I - 1) * 10 + 1
        barStatus.Msg = MyItem(I).TestNm & " 항목을 검색중입니다."
        DoEvents
        
        strTestcd = MyItem(I).TestCd
        strPanelFg = MyItem(I).PanelFg
        strTestDiv = MyItem(I).TestDiv
        strSpcCd = MyItem(I).SpcCd
        strWorkArea = MyItem(I).WorkArea
        strTestNm = MyItem(I).TestNm
        SqlStmt = objRstSQL.SqlCumResult(txtPtId.Text, strFromDt, strToDt, strTestcd, _
                                     strPanelFg, strTestDiv, strSpcCd, strWorkArea)
        If Trim(SqlStmt) <> "" Then
        Set RS = Nothing
        Set RS = New Recordset
        RS.Open SqlStmt, DBConn
                
        While (Not RS.EOF)
        
            If barStatus.Value < (I * 10) Then barStatus.Value = barStatus.Value + 1
            DoEvents
            
            iSeq = 0
            strDtTm = Format("" & RS.Fields("ColDt").Value, CS_DateMask) & "  " & _
                      Format("" & RS.Fields("ColTm").Value, CS_TimeShortMask)
            strList = strDtTm & vbTab & Trim(CStr(iSeq))
            If medListFind(lstDtTm, strList) < 0 Then lstDtTm.AddItem strList
            
            
            Set clsNewData = New clsCumResult
            With clsNewData
                .ColDt = Format("" & RS.Fields("ColDt").Value, CS_DateShortMask)
                .ColTm = Format("" & RS.Fields("ColTm").Value, CS_TimeShortMask)
                .DeptCd = "" & RS.Fields("DeptCd").Value
                .WardId = "" & RS.Fields("WardId").Value
                .HosilId = "" & RS.Fields("HosilId").Value
                .TestCd = "" & RS.Fields("TestCd").Value
                .TestNm = strTestNm
                .SpcCd = "" & RS.Fields("SpcCd").Value
                .SpcNm = "" & RS.Fields("TestCd").Value
                .DPDiv = "" & RS.Fields("DpDiv").Value
                .HLDiv = "" & RS.Fields("HlDiv").Value
                .RstDiv = "" & RS.Fields("RstDiv").Value
                .WorkArea = "" & RS.Fields("WorkArea").Value
                .AccDt = "" & RS.Fields("AccDt").Value
                .AccSeq = "" & RS.Fields("AccSeq").Value
                .FootNoteFg = "" & RS.Fields("FootNoteFg").Value
                .RmkCd = "" & RS.Fields("RmkCd").Value
                'If Trim("" & rs.Fields("RstType").Value) = "N" Then
                '    .RstCd = "" & rs.Fields("RstCd").Value
                'Else
                    If Trim("" & RS.Fields("RstCdNm").Value) <> "" Then
                        .RstCd = "" & RS.Fields("RstCdNm").Value
                    Else
                        .RstCd = "" & RS.Fields("RstCd").Value
                    End If
                'End If
                .RstUnit = "" & RS.Fields("RstUnit").Value
                '.TestDiv = "" & rs.Fields("TestDiv").Value
                .TxtFg = "" & RS.Fields("TxtFg").Value
                .Remark = ""
                If Trim(.RmkCd) <> "" Then
                    .Remark = MyResult.ReadRemark(.RmkCd)
                    If medListFind(lstRemark, strDtTm) < 0 Then lstRemark.AddItem strDtTm & vbTab & CStr(mCumCol.Count + 1)
                    'txtSamCmt.Text = txtSamCmt.Text & "-- " & strDtTm & " --" & vbCRLF & "remark : " & .Remark & vbCRLF
                End If
                .FootNote = ""
                If Trim(.FootNoteFg) <> "0" Then
                    .FootNote = MyResult.ReadFootNote(.WorkArea, .AccDt, .AccSeq)
                    If medListFind(lstRemark, strDtTm) < 0 Then lstRemark.AddItem strDtTm & vbTab & CStr(mCumCol.Count + 1)
                    'If .Remark = "" Then txtSamCmt.Text = txtSamCmt.Text & "-- " & strDtTm & " --" & vbCRLF
                    'txtSamCmt.Text = txtSamCmt.Text & .FootNote & vbCRLF
                End If
                    
                .RstText = "" & RS.Fields("TextResult").Value
                'If .RstText <> "" Then
                '    txtRstCmt.Text = txtRstCmt.Text & "-- " & strDtTm & " -- " & .TestNm & vbCRLF & .RstText & vbCRLF
                'End If
                
                On Error GoTo Dup_Err
                             
                If ObjDic.Exists(strDtTm & COL_DIV & .TestCd & Trim(CStr(iSeq))) = False Then
                    ObjDic.AddNew strDtTm & COL_DIV & .TestCd & Trim(CStr(iSeq)), Trim(CStr(iSeq))
                    
                    mCumCol.Add clsNewData, strDtTm & ":" & .TestCd & ":" & Trim(CStr(iSeq))
                End If
                                
            End With
            RS.MoveNext
        Wend
        End If
        Set RS = Nothing
    Next
    Set ObjDic = Nothing
    Exit Sub
    
Dup_Err:
    If Err.Number = 457 Then
        iSeq = iSeq + 1
        strList = strDtTm & vbTab & Trim(CStr(iSeq))
        If medListFind(lstDtTm, strList) < 0 Then lstDtTm.AddItem strList
        Resume
    Else
        MsgBox Err.Number & "  " & Err.Description, vbCritical, "Error"
        Set RS = Nothing
    End If
    Set ObjDic = Nothing
    
End Sub

Private Sub DisplayOnePage(ByVal iCurPage As Integer)
    Dim I As Integer
    Dim J As Integer
    Dim iListIndex As Integer
    Dim sDtTm As String
    Dim sSeq As String
    Dim sDPfg As String
    Dim clsData As clsCumResult
    Dim ErrFg As Boolean
    Dim EvenBkColor As Long, OddBkColor As Long
    
    EvenBkColor = &HF9FBFA
    OddBkColor = &HFFFFFF
    
    With tblResult
        .Row = 0: .Row2 = .MaxRows
        .Col = 5: .Col2 = 14
        .BlockMode = True
        .Text = ""
        .BlockMode = False
        
        '.ReDraw = False
        
        For I = 1 To .MaxRows
            .Row = I: .Row2 = I
            .Col = 5: .Col2 = 14
            .BlockMode = True
            If I <> OldRow Then
                .BackColor = IIf((I Mod 2) = 0, EvenBkColor, OddBkColor)
            End If
            .ForeColor = vbBlack
            .BlockMode = False
            .Col = 17: .Value = ""
        Next
        
        For I = (iCurPage - 1) * 10 To iCurPage * 10 - 1
            .Row = 0
            If I >= lstDtTm.ListCount Then Exit For
            
            '가장 최근날짜부터 Display하기 위해 Index계산...
            iListIndex = lstDtTm.ListCount - I - 1
            
            .Col = I - ((iCurPage - 1) * 10) + 5
            .Text = Format(medGetP(lstDtTm.List(iListIndex), 1, vbTab), CS_DateShortFormat & "  " & CS_TimeShortFormat)
            sDtTm = medGetP(lstDtTm.List(iListIndex), 1, vbTab)
            sSeq = medGetP(lstDtTm.List(iListIndex), 2, vbTab)
            For J = 1 To mItemCount
                
                sDPfg = ""
                ErrFg = False
                
                On Error GoTo Err_Trap
    
                .Row = J
                .Col = I - ((iCurPage - 1) * 10) + 5
                Set clsData = mCumCol.Item(sDtTm & ":" & MyItem(J).TestCd & ":" & sSeq)
                If ErrFg Then GoTo Skip
                .Value = clsData.RstCd
                If clsData.HLDiv = "H" Then
                    .ForeColor = &H7477EF   'vbRed
                ElseIf clsData.HLDiv = "L" Then
                    .ForeColor = &HDF6A3E  'vbBlue
                End If
                If clsData.DPDiv <> "" Then
                    sDPfg = clsData.DPDiv
                    .Value = .Value & " " & clsData.DPDiv
                    '.FontBold = True
                    .ForeColor = vbRed
                    .BackColor = &HC0FFFF     '&HFFF7FF
'                    .CellBorderStyle = CellBorderStyleSolid
'                    .CellBorderType = 16
'                    .CellBorderColor = &H7477EF
'                    .Action = ActionSetCellBorder
                End If
                .Col = 15
                If .Value = "" Then .Value = clsData.RstUnit
Skip:
                .Col = 17
                .Value = .Value & sDPfg & ":"
                
            Next
            
            If I = (iCurPage - 1) * 10 Then
'                Dim objDept As clsBasisData
                Dim strDept As String
'                Set objDept = Nothing
'                Set objDept = New clsBasisData
                strDept = GetDeptNm(clsData.DeptCd)
'                Set objDept = Nothing
                
'                If ObjLISComCode.DeptCd.Exists(clsData.DeptCd) Then
'                    ObjLISComCode.DeptCd.KeyChange (clsData.DeptCd)
                    lblDeptNm.Caption = strDept 'ObjLISComCode.DeptCd.Fields("deptnm")
'                    'lblDeptNm.Caption = MyPatient.GetDeptNm(clsData.DeptCd)
'                End If
                lblWardId.Caption = clsData.WardId & " - " & clsData.HosilId
            End If
                
        Next
        '.ReDraw = True
    End With
    Set clsData = Nothing
    Exit Sub
    
Err_Trap:
    ErrFg = True
    Resume Next
End Sub

Private Sub SetStartDt()
    
    Dim sDate As Date
    
    sDate = objResult.SetStartDt(gDeptCd)   'mDeptCd)
    dtpFromDt.Value = Format(sDate, "YY-MM") & "-01"
    dtpToDt.Value = Format(GetSystemDate, "yyyy-mm-dd")

End Sub

Private Sub ClearRtn()

    ClearFg = True

    txtCumCd.Text = ""
    optCumCd(0).Value = True
    cmdItemAdd.Enabled = False
    txtSamCmt.Text = ""
    txtRstCmt.Text = ""
    txtSamCmt1.Text = ""
    txtRstCmt1.Text = ""
    cmdQuery.Enabled = True
    cmdPrint.Enabled = False
    cmdExcel.Enabled = False
    cmdNext(0).Enabled = False
    cmdNext(1).Enabled = False
    Call ClearTable
    Call ClearGraph
    grpResult.BorderStyle = 1
    OldRow = -1
    chkGraph.Value = 0
    Erase MyItem
    mItemCount = 0
    ReDim MyItem(mItemCount)
    
    iPageNo = 0
    iPageCnt = 0
    grpResult.Visible = False
    cmdPrintGraph.Visible = False
    Call chkGraph_Click
    Set mCumCol = New Collection
    
    lstDtTm.Clear
    lstRemark.Clear

'    txtPtId.Locked = False
'    txtPtId.BackColor = vbWhite
    txtCumCd.Locked = False
    txtCumCd.BackColor = vbWhite
    lstCumList.Enabled = True
    optCumCd(0).Enabled = True
    optCumCd(1).Enabled = True
    cmdItemAdd.Enabled = True
    dtpFromDt.Enabled = True

    cboWorkArea.ListIndex = -1
    txtSpcCd.Text = ""
    txtSpcCd.Enabled = True
    txtSpcCd.BackColor = vbWhite
    lblSpcNm.Caption = ""
    Call cmdReset_Click
    lstSelList.Clear

End Sub

Private Sub ClearTable()
    With tblResult
        .MaxRows = 0
        .Col = 5: .Col2 = 14
        .Row = 0: .Row2 = 0
        .BlockMode = True
        .Text = ""
        .BlockMode = False
        
'        .Row = 0: .Col = 2
'        .FontBold = True
'        .FontUnderline = True
    End With
End Sub

Private Sub ClearGraph()
    With grpResult
        .ClearData CD_VALUES
        .ClearLegend CHART_LEGEND
    End With
End Sub


Private Sub ShowGraph(ByVal iGrpRow As Integer)

    Dim I As Integer, J As Integer
    Dim FirstFg As Boolean
    Dim iSeries As Integer, iPoints As Integer
    Dim iMaxValue As Double, iMinValue As Double
    Dim iFromRef As Double, iToRef As Double
    Dim sPnt As Integer, ePnt As Integer
    Dim sXVal As Integer, eXVal As Integer
    Dim tmpStr As String
    Dim clsData As clsCumResult
    Dim ErrFg As Boolean
    Dim sDtTm  As String, sSeq As String
    
    
    FirstFg = True
    
    iSeries = 1
    iPoints = 0
    
    'Call InitDraw(iSeries, iPoints)
    Call SetDateRange(sPnt, ePnt)
    Call ClearGraph
    
    With tblResult
        .Row = iGrpRow: .Col = 2
        grpResult.Title(CHART_TOPTIT) = .Value
        grpResult.ClearData CD_VALUES
        'grpResult.ClearLegend CHART_LEGEND
        
        grpResult.RealTimeStyle = CRT_LOOPPOS Or CRT_NOWAITARROW
        grpResult.OpenDataEx COD_VALUES, iSeries, lstDtTm.ListCount
        
        grpResult.TopGap = 20
        grpResult.BottomGap = 25
        grpResult.FixedGap = 33
        grpResult.Grid = CHART_NOGRID
        grpResult.Scrollable = True
        
        .Row = iGrpRow: .Col = 16
        iFromRef = Val(medGetP(.Value, 1, "-"))
        iToRef = Val(medGetP(.Value, 2, "-"))
        iMinValue = iFromRef '- (iFromRef / 50) '2
        iMaxValue = iToRef '+ (iFromRef / 50) '2
        
        'grpResult.ItemStyle(CI_HORZGRID) = CHART_SOLID
         
        grpResult.Scrollable = True
        grpResult.PointLabels = True
        grpResult.RGBFont(CHART_POINTFT) = vbBlue
        grpResult.Axis(AXIS_X).STEP = 1
        'grpResult.Axis(AXIS_X).Decimals = 0
        'grpresult.PointLabelsFont.Bold = False
        
        'Call SetSerLeg
        'Call SetLegend
        'Call chkTable_Click
        
        'For i = 0 To iSeries - 1
        '    grpResult.Series(i).COLOR = GrpColor(i)
        'Next
        
        '-- Stripe Color & Series Color
        'grpResult.COLOR(1) = COLOR(2) 'QBColor(I)
        
        grpResult.ThisSerie = 0
        For I = lstDtTm.ListCount - 1 To 0 Step -1
            
            sDtTm = medGetP(lstDtTm.List(I), 1, vbTab)
            sSeq = medGetP(lstDtTm.List(I), 2, vbTab)
            
            ErrFg = False
                
            On Error GoTo Err_Trap
    
            Set clsData = mCumCol.Item(sDtTm & ":" & MyItem(iGrpRow).TestCd & ":" & sSeq)
            If ErrFg Then GoTo Skip
            If Not IsNumeric(clsData.RstCd) Then GoTo Skip
            
            grpResult.KeyLeg(iPoints) = Format(sDtTm, "MM-DD")
            grpResult.Value(iPoints) = Val(clsData.RstCd)
'            grpResult
            iPoints = iPoints + 1
            
            If I = sPnt Then sXVal = iPoints
            If I = ePnt Then eXVal = iPoints
            
            If iMinValue > Val(clsData.RstCd) Then iMinValue = Val(clsData.RstCd)
            If iMaxValue < Val(clsData.RstCd) Then iMaxValue = Val(clsData.RstCd)
                    
Skip:
        Next
        
        If iPoints = 0 Then
            Call ClearGraph
            Exit Sub
        End If
        
        grpResult.CloseData COD_VALUES
        
        grpResult.OpenDataEx COD_STRIPES, 2, 0
        '참고치 구간 표시...
        grpResult.Stripe(0).Axis = AXIS_Y
        grpResult.Stripe(0).COLOR = &HC0FFFF
        grpResult.Stripe(0).From = iFromRef
        grpResult.Stripe(0).To = iToRef
        'Spread에 보여지고 있는 구간 표시...
        grpResult.Stripe(1).Axis = AXIS_X
        grpResult.Stripe(1).COLOR = &HDBF2FD          '&HD6EAFA       ' &HD6EAFA        '&HFFF9F4     '&HF4FEED   '&HD6D7FA     '&HFFF4FF  '&HF7FFFF  '&HEEF4F4  '&HEEEEEE
        grpResult.Stripe(1).From = sXVal
        grpResult.Stripe(1).To = eXVal
        grpResult.CloseData COD_STRIPES
        
        grpResult.OpenDataEx COD_CONSTANTS, 2, 0
        
        grpResult.ConstantLine(0).Value = iFromRef
        grpResult.ConstantLine(0).LineColor = &H808080
        grpResult.ConstantLine(0).Axis = AXIS_Y
        grpResult.ConstantLine(0).Label = CStr(iFromRef)
        grpResult.ConstantLine(0).LineWidth = 1
        grpResult.ConstantLine(0).LineStyle = CHART_DOT
        
        grpResult.ConstantLine(1).Value = iToRef
        grpResult.ConstantLine(1).LineColor = &H808080  '&H80&
        grpResult.ConstantLine(1).Axis = AXIS_Y
        grpResult.ConstantLine(1).Label = CStr(iToRef)
        grpResult.ConstantLine(1).LineWidth = 1
        grpResult.ConstantLine(1).LineStyle = CHART_DOT
        
        grpResult.CloseData COD_CONSTANTS
        
        grpResult.OpenDataEx COD_VALUES, iSeries, iPoints
            
        grpResult.Axis(AXIS_Y).Min = iMinValue - ((iMaxValue - iFromRef) / 10) '1
        grpResult.Axis(AXIS_Y).Max = iMaxValue + ((iMaxValue - iFromRef) / 10) '1
        
        grpResult.Axis(AXIS_Y).STEP = (iMaxValue - iMinValue) / 3
        
        grpResult.CloseData COD_VALUES
    
    End With
    Exit Sub

Err_Trap:
    ErrFg = True
    Resume Next
End Sub


'Private Sub InitDraw(ByVal nSeries As Integer, ByVal nPoints As Integer)
'
'    Dim iMaxValue As Long
'    Dim iSS As Integer, iPT As Integer, iCnt As Long, iVal As Long
'    Dim I As Integer
'
'    With ssDataBuf
'
'        grpResult.ClearData CD_VALUES
'        grpResult.OpenDataEx COD_VALUES, nSeries, nPoints
'
'        For I = 0 To .MaxRows - 1
'            .Row = I + 1
'            .Col = COL_SERIES:  iSS = Val(.Value)
'            .Col = COL_POINTS:  iPT = Val(.Value)
'            .Col = COL_COUNT:   iCnt = Val(.Value)
'
'            iVal = grpResult.ValueEx(iSS - 1, iPT - 1)
'            grpResult.ValueEx(iSS - 1, iPT - 1) = iVal + iCnt
'
'            iVal = grpResult.ValueEx(iSS - 1, iPT - 1)
'            If iMaxValue < iVal Then iMaxValue = iVal
'
'            .Col = cboXVal.ItemData(cboXVal.ListIndex)
'            'grpresult.Axis(AXIS_X).Label(iPT - 1) = .Value
'            grpResult.Legend(iPT - 1) = .Value
'        Next I
'
'        grpResult.Axis(AXIS_Y).Max = iMaxValue + 1
'
'    End With
'
'End Sub



Private Sub txtSpcCd_Change()
    lstSpcList.ListIndex = medListFind(lstSpcList, txtSpcCd.Text)
    lstItemList.Clear
    'lstSelList.Clear
End Sub

Private Sub txtSpcCd_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown And lstSpcList.ListCount > 0 Then
        lstSpcList.Visible = True
        lstSpcList.ZOrder 0
        lstSpcList.SetFocus
    End If

End Sub

Private Sub txtSpcCd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lstSpcList_MouseDown(1, 0, 0, 0)
        lstItemList.SetFocus
        Exit Sub
    End If
    If lstSpcList.ListCount > 0 Then
        lstSpcList.Visible = True
        lstSpcList.ZOrder 0
        Call medCodeHelp(KeyAscii, lstSpcList, txtSpcCd.Text, txtSpcCd, lstItemList)
    End If
End Sub

Private Sub lstSpcList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call lstSpcList_MouseDown(1, 0, 0, 0)
        lstItemList.SetFocus
    End If
End Sub

Private Sub lstSpcList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        lstSpcList.Visible = False
        txtSpcCd.Text = medGetP(lstSpcList.Text, 1, vbTab)
        lblSpcNm.Caption = medGetP(lstSpcList.Text, 2, vbTab)
        DoEvents
        Call objResult.LoadItem(lstItemList, medGetP(cboWorkArea.Text, 1, " "), txtSpcCd.Text)
    End If
    
End Sub

Private Function SpcCheck(ByVal pSpcCd As String) As Boolean
    SpcCheck = True
    If mItemCount = 0 Then Exit Function
    If MyItem(mItemCount).SpcCd <> pSpcCd Then SpcCheck = False
End Function

Private Sub SetDateRange(sPnt As Integer, ePnt As Integer)

    Dim I As Integer
    Dim sDt As String, eDt As String
    
    With tblResult
        For I = 1 To 10
            .Row = OldRow
            .Col = I + 4
            If IsNumeric(.Value) Then
                sPnt = lstDtTm.ListCount - ((iPageNo - 1) * 10) - I
                Exit For
            End If
        Next
        For I = 10 To 1 Step -1
            .Row = OldRow
            .Col = I + 4
            If IsNumeric(.Value) Then
                ePnt = lstDtTm.ListCount - ((iPageNo - 1) * 10) - I
                Exit For
            End If
        Next
    End With
    
End Sub


Private Sub PrintGraph()

    With grpResult
        .Printer.TopMargin = 2
        .Printer.LeftMargin = 0
        .Printer.RightMargin = 1
        .Printer.BottomMargin = 2
        .Printer.Compress = True
        .Printer.Orientation = ORIENTATION_LANDSCAPE
        .Printer.ForceColors = True
        .PrintIt 0, 0
    End With
    
End Sub


Public Sub LoadCumList(ByRef lstList As ListBox, Optional ByVal pDeptCd As String = "ALL")
    Dim SqlStmt As String
    Dim RS As Recordset
    
    SqlStmt = MySql.SqlGetCumList(pDeptCd)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    lstList.Clear
    If pDeptCd = "" Then
        While (Not RS.EOF)
            lstList.AddItem Trim(RS.Fields("CumCd").Value) & vbTab & _
                            Trim(RS.Fields("CumNm").Value) & vbTab & Space(20) & Trim(RS.Fields("CumCd").Value & "")
            RS.MoveNext
        Wend
    Else
        While (Not RS.EOF)
            lstList.AddItem Trim(RS.Fields("field2").Value) & vbTab & _
                            Trim(RS.Fields("CumNm").Value) & vbTab & Space(20) & _
                            Trim(RS.Fields("CumCd").Value)
            RS.MoveNext
        Wend
    End If
    Set RS = Nothing
    Set MySql = Nothing
    DoEvents
End Sub

Public Sub LoadSpcList(ByRef lstList As ListBox)

'    lstList.Clear
'
'    ObjLISComCode.LisSpc.MoveFirst
'    While (Not ObjLISComCode.LisSpc.EOF)
'        lstList.AddItem ObjLISComCode.LisSpc.Fields("spccd") & vbTab & _
'                        ObjLISComCode.LisSpc.Fields("spcnm")
'        ObjLISComCode.LisSpc.MoveNext
'    Wend
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm, " & _
             "       a.field1 multifg, a.field2 spcgrp, b.field2 labrange  " & _
             "FROM " & T_LAB032 & " b, " & T_LAB032 & " a " & _
             "WHERE  a.cdindex = 'C215' " & _
             "AND    " & DBJ("b.cdindex = 'C217'") & _
             "AND    " & DBJ("b.cdval1  =* a.field2")
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    lstList.Clear
    Do Until RS.EOF
        lstList.AddItem RS.Fields("spccd").Value & "" & vbTab & _
                        RS.Fields("spcnm").Value & ""
        RS.MoveNext
    Loop
    
    Set RS = Nothing
End Sub

Public Sub LoadSpcItem(ByRef lstList As ListBox, ByRef lstList1 As ListBox, ByVal pSpcCd As String)

    Dim SqlStmt As String
    Dim RS As Recordset
    Dim tmpStr As String
    Dim I%
    
    '상세항목 제외...
    SqlStmt = MySql.SqlLoadSpcItem(pSpcCd)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    lstList.Clear
    lstList1.Clear
    If RS.EOF Then GoTo NoData
    
    For I = 1 To RS.RecordCount
        tmpStr = RS.Fields("TestCd").Value & Space(9)
        lstList.AddItem Mid(tmpStr, 1, 10) & _
                        RS.Fields("TestNm").Value
        lstList1.AddItem RS.Fields("TestNm").Value & vbTab & RS.Fields("TestCd").Value
        RS.MoveNext
    Next I
    
NoData:
    Set RS = Nothing
End Sub


Public Sub Call_PtId_KeyPress()

   Call txtPtId_LostFocus

End Sub





