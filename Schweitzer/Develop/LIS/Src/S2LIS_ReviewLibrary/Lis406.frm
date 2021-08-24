VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Begin VB.Form frm406OldCumulative 
   Caption         =   "Cumulative Result ( Old Data )"
   ClientHeight    =   9450
   ClientLeft      =   -120
   ClientTop       =   165
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   14670
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdPrintGraph 
      BackColor       =   &H00DBF2FD&
      Caption         =   "Print"
      Height          =   315
      Left            =   13725
      Style           =   1  '그래픽
      TabIndex        =   62
      Top             =   6870
      Width           =   885
   End
   Begin DRcontrol1.DrFrame fraAddItem 
      Height          =   5355
      Left            =   6180
      TabIndex        =   48
      Top             =   1725
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   9446
      Title           =   "검사항목 추가"
      BackColor       =   13753559
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboWorkArea 
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
         TabIndex        =   59
         Top             =   570
         Width           =   3480
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
         TabIndex        =   56
         Top             =   2670
         Width           =   4140
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "확인"
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
         Left            =   2790
         TabIndex        =   55
         Top             =   4830
         Width           =   765
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "취소"
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
         Left            =   3585
         TabIndex        =   54
         Top             =   4830
         Width           =   765
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
         TabIndex        =   53
         Top             =   1335
         Width           =   4140
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "지움"
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
         Left            =   2010
         TabIndex        =   52
         Top             =   4830
         Width           =   750
      End
      Begin VB.TextBox txtSpcCd 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   51
         Top             =   930
         Width           =   945
      End
      Begin VB.CommandButton cmdSpcList 
         BackColor       =   &H00D1DCD7&
         Caption         =   "▼"
         Height          =   300
         Left            =   1800
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   49
         Top             =   945
         Width           =   285
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   2100
         TabIndex        =   50
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
      Begin VB.ListBox lstSpcList 
         BackColor       =   &H00FCF8FB&
         Height          =   2205
         Left            =   825
         TabIndex        =   60
         Top             =   1245
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "Work Area"
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
         Left            =   255
         TabIndex        =   58
         Tag             =   "40202"
         Top             =   540
         Width           =   465
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 체"
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
         Left            =   255
         TabIndex        =   57
         Tag             =   "40202"
         Top             =   1020
         Width           =   420
      End
   End
   Begin MedControls1.LisLabel lblMsg 
      Height          =   495
      Left            =   3990
      TabIndex        =   46
      Top             =   3600
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
   Begin VB.CommandButton cmdItemAdd 
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
      Left            =   5085
      TabIndex        =   45
      Top             =   1725
      Width           =   1080
   End
   Begin MSComctlLib.TabStrip tabCumCd 
      Height          =   315
      Left            =   3180
      TabIndex        =   44
      Top             =   1755
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      Style           =   2
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Default"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "과별"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin VB.ListBox lstRemark 
      Height          =   2595
      Left            =   2805
      Sorted          =   -1  'True
      TabIndex        =   43
      Top             =   4275
      Visible         =   0   'False
      Width           =   2625
   End
   Begin ChartfxLibCtl.ChartFX grpResult 
      Height          =   2145
      Left            =   165
      TabIndex        =   11
      Top             =   6870
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
      RGBBk           =   14739427
      RGB2DBk         =   16777215
      RGB3DBk         =   13753559
      nColors         =   1
      Colors          =   "Lis406.frx":0000
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
      _Data_          =   "Lis406.frx":0028
   End
   Begin VB.ListBox lstDtTm 
      Height          =   2595
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   42
      Top             =   4275
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H00FFF4FF&
      BorderStyle     =   0  '없음
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
      Height          =   1320
      Left            =   4665
      TabIndex        =   39
      Top             =   3030
      Visible         =   0   'False
      Width           =   6585
      Begin MSComctlLib.ProgressBar barStatus 
         Height          =   195
         Left            =   1410
         TabIndex        =   40
         Top             =   390
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   405
         TabIndex        =   41
         Top             =   885
         Width           =   6045
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   495
         Picture         =   "Lis406.frx":01C1
         Top             =   255
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   480
         Picture         =   "Lis406.frx":0603
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   60
         Top             =   45
         Width           =   6465
      End
   End
   Begin VB.CommandButton cmdCumItem 
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
      Left            =   4800
      TabIndex        =   38
      Top             =   2235
      Width           =   1380
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   4140
      Left            =   150
      TabIndex        =   37
      Top             =   2715
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
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   17
      MaxRows         =   50
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14739427
      ShadowDark      =   13753559
      SpreadDesigner  =   "Lis406.frx":0A45
      TextTip         =   4
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료(&X)"
      Height          =   570
      Left            =   13335
      TabIndex        =   10
      Tag             =   "128"
      Top             =   2085
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   570
      Left            =   13335
      TabIndex        =   9
      Tag             =   "124"
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력(&P)"
      Height          =   570
      Left            =   13335
      TabIndex        =   8
      Tag             =   "132"
      Top             =   855
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFromDt 
      Height          =   360
      Left            =   7155
      TabIndex        =   3
      Top             =   60
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   64028672
      CurrentDate     =   36567
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00D1DCD7&
      Caption         =   "조회(&Q)"
      Height          =   660
      Left            =   13320
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "133"
      Top             =   90
      Width           =   1230
   End
   Begin RichTextLib.RichTextBox txtSamCmt 
      Height          =   735
      Left            =   7125
      TabIndex        =   23
      Top             =   450
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   15658734
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Lis406.frx":17B8
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
   Begin VB.CommandButton cmdNext 
      Caption         =   "(&N) >>"
      Height          =   465
      Index           =   1
      Left            =   2850
      TabIndex        =   6
      Top             =   2235
      Width           =   1380
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "<< (&B)"
      Height          =   465
      Index           =   0
      Left            =   1425
      TabIndex        =   5
      Top             =   2235
      Width           =   1380
   End
   Begin VB.CheckBox chkGraph 
      Caption         =   "그래프(&G)"
      Height          =   270
      Left            =   165
      TabIndex        =   7
      Tag             =   "40201"
      Top             =   2370
      Width           =   1260
   End
   Begin VB.TextBox txtCumCd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4035
      TabIndex        =   1
      Top             =   60
      Width           =   2130
   End
   Begin VB.TextBox txtPtId 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1125
      TabIndex        =   0
      Top             =   90
      Width           =   1965
   End
   Begin RichTextLib.RichTextBox txtRstCmt 
      Height          =   1515
      Left            =   7125
      TabIndex        =   24
      Top             =   1155
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   2672
      _Version        =   393217
      BackColor       =   16252927
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Lis406.frx":185D
      MouseIcon       =   "Lis406.frx":1902
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
   Begin VB.ListBox lstCumList 
      BackColor       =   &H00EEF4F4&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   3180
      TabIndex        =   2
      Top             =   450
      Width           =   2970
   End
   Begin MedControls1.LisLabel lblPtNm 
      Height          =   330
      Left            =   1125
      TabIndex        =   33
      Top             =   450
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      BackColor       =   16052466
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
      Caption         =   ""
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel lblDeptNm 
      Height          =   330
      Left            =   1125
      TabIndex        =   36
      Top             =   1440
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      BackColor       =   16052466
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
      Caption         =   ""
      LeftGab         =   100
   End
   Begin DRcontrol1.DrFrame fraTextResult 
      Height          =   8040
      Left            =   3345
      TabIndex        =   28
      Top             =   1140
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
         TabIndex        =   47
         Top             =   135
         Width           =   285
      End
      Begin RichTextLib.RichTextBox txtSamCmt1 
         Height          =   2265
         Left            =   165
         TabIndex        =   29
         Top             =   450
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   3995
         _Version        =   393217
         BackColor       =   16252927
         Enabled         =   -1  'True
         TextRTF         =   $"Lis406.frx":1A64
      End
      Begin RichTextLib.RichTextBox txtRstCmt1 
         Height          =   4815
         Left            =   165
         TabIndex        =   30
         Top             =   3000
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   8493
         _Version        =   393217
         BackColor       =   16710910
         Enabled         =   -1  'True
         TextRTF         =   $"Lis406.frx":1B01
         MouseIcon       =   "Lis406.frx":1B9E
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
         TabIndex        =   32
         Tag             =   "40205"
         Top             =   180
         Width           =   2370
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
         TabIndex        =   31
         Tag             =   "40204"
         Top             =   2790
         Width           =   2205
      End
   End
   Begin VB.Label Label4 
      Caption         =   "☞ 2000년 2월 18일 종결된 결과까지    조회 가능합니다..."
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00864B24&
      Height          =   405
      Left            =   9930
      TabIndex        =   61
      Top             =   75
      Width           =   3135
   End
   Begin VB.Label lblAgeDiv 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
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
      Left            =   2685
      TabIndex        =   35
      Top             =   810
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "진료부서"
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
      Left            =   255
      TabIndex        =   34
      Tag             =   "102"
      Top             =   1515
      Width           =   720
   End
   Begin VB.Label lblLocation 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "병      실"
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
      Left            =   255
      TabIndex        =   27
      Tag             =   "102"
      Top             =   1845
      Width           =   720
   End
   Begin VB.Label lblWardId 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00F4F0F2&
      BorderStyle     =   1  '단일 고정
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
      Left            =   1125
      TabIndex        =   26
      Top             =   1770
      Width           =   1965
   End
   Begin VB.Label lblFrDt 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "시  작  일"
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
      Left            =   6300
      TabIndex        =   25
      Tag             =   "154"
      Top             =   135
      Width           =   795
   End
   Begin VB.Label lblRstCmt 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      Caption         =   "검사소견"
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
      Left            =   6330
      TabIndex        =   22
      Tag             =   "40204"
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label lblSamCmt 
      AutoSize        =   -1  'True
      Caption         =   "Remark"
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
      Left            =   6375
      TabIndex        =   21
      Tag             =   "40205"
      Top             =   540
      Width           =   645
   End
   Begin VB.Label lblRptNm 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "누적코드"
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
      Left            =   3225
      TabIndex        =   20
      Tag             =   "40202"
      Top             =   165
      Width           =   720
   End
   Begin VB.Label lblPtDob 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00F4F0F2&
      BorderStyle     =   1  '단일 고정
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
      Left            =   1125
      TabIndex        =   18
      Top             =   1110
      Width           =   1965
   End
   Begin VB.Label lblDOB 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "생년월일"
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
      Left            =   255
      TabIndex        =   17
      Tag             =   "101"
      Top             =   1185
      Width           =   720
   End
   Begin VB.Label lblPtAge 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
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
      Left            =   2040
      TabIndex        =   16
      Top             =   810
      Width           =   390
   End
   Begin VB.Label lblPtSex 
      BackStyle       =   0  '투명
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
      Left            =   1380
      TabIndex        =   15
      Top             =   855
      Width           =   480
   End
   Begin VB.Label lblSexAge 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "성별/연령"
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
      Left            =   255
      TabIndex        =   14
      Tag             =   "108"
      Top             =   855
      Width           =   810
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "성      명"
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
      Left            =   255
      TabIndex        =   13
      Tag             =   "103"
      Top             =   525
      Width           =   720
   End
   Begin VB.Label lblPtId 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "환자  I D"
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
      Left            =   255
      TabIndex        =   12
      Tag             =   "105"
      Top             =   180
      Width           =   705
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F4F0F2&
      BorderStyle     =   1  '단일 고정
      Caption         =   "             /"
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
      Left            =   1125
      TabIndex        =   19
      Top             =   780
      Width           =   1965
   End
End
Attribute VB_Name = "frm406OldCumulative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private MyPatient As New clsPatient   '환자 클래스
'Private MySql As New clsSqlStatements   'Sql문 클래스
'
'Private Type MyItem
'    TestCd As String
'    PanelFg As String
'    TestDiv As String
'    SpcCd As String
'    WorkArea As String
'    testnm As String
'    SpcNm As String
'    RefVal As String
'End Type
'
'Private MyItem() As MyItem
'Private mItemCount As Integer
'Private iPageNo As Integer
'Private iPageCnt As Integer
'Private OldRow As Integer
'Private OldColor As Long
'
'Const iColPerPage = 10
'
'Private mCumCol As Collection
'Private mDeptCd As String
'
'Public PtFg As Boolean
'Public QueryFg As Boolean
'
'Private ClearFg As Boolean
'
'
'Public Property Get DeptCd() As String
'    DeptCd = mDeptCd
'End Property
'
'Public Property Let DeptCd(ByVal vNewValue As String)
'    mDeptCd = vNewValue
'End Property
'
'
'Private Sub cboWorkArea_Click()
'    Call LoadItem(lstItemList, medGetP(cboWorkArea.Text, 1, " "), txtSpcCd.Text)
'End Sub
'
'Private Sub chkGraph_Click()
'    If chkGraph.Value = 1 Then
'        'If Not grpResult.Visible Then
'            grpResult.Visible = True
'            cmdPrintGraph.Visible = True
'            tblResult.Height = grpResult.Top - tblResult.Top
'            If OldRow > 0 Then
'                Call tblResult_Click(2, OldRow)
'                tblResult.TopRow = OldRow
'            Else
'                Call tblResult_Click(2, 1)
'                tblResult.TopRow = 1
'            End If
'        'End If
'    Else
'        'If grpResult.Visible Then
'            grpResult.Visible = False
'            cmdPrintGraph.Visible = False
'            tblResult.Height = grpResult.Top - tblResult.Top + grpResult.Height + 20
'        'End If
'    End If
'End Sub
'
'Private Sub cmdCANCEL_Click()
'    txtCumCd.SetFocus
'    fraAddItem.Visible = False
'End Sub
'
'Private Sub cmdClear_Click()
'    Call ClearRtn
'    'txtPtId.Text = ""
'    txtCumCd.Text = ""
'    'Call SetStartDt
'    txtPtId.SetFocus
'End Sub
'
'Private Sub cmdClose_Click()
'    fraTextResult.Visible = False
'End Sub
'
'Private Sub cmdCumItem_Click()
'    lblMsg.Caption = "누적코드를 로딩중입니다. 잠시만 기다리세요...."
'    lblMsg.Visible = True
'    lblMsg.ZOrder 0
'    DoEvents
'    frm4021CumCdSet.DeptCd = mDeptCd
'    frm4021CumCdSet.IsManager = False
'    frm4021CumCdSet.Show 1
'    lblMsg.Visible = False
'End Sub
'
'Private Sub cmdExit_Click()
'   Unload Me
'   Set frm402Cumulative = Nothing
'End Sub
'
'Private Sub cmdItemAdd_Click()
'    lblMsg.Caption = "검사항목 리스트를 로드하고 있습니다. 잠시만 기다리세요...."
'    lblMsg.Visible = True
'    DoEvents
'    cboWorkArea.ListIndex = -1
''    If mItemCount > 0 Then
''        txtSpcCd.Text = MyItem(mItemCount).SpcCd
''        tblResult.Row = 0: tblResult.Col = 2
''        lblSpcNm.Caption = tblResult.Text
''    Else
'    txtSpcCd.Text = ""
'    lblSpcNm.Caption = ""
''    End If
'    If lstSpcList.ListCount = 0 Then Call LoadSpcList(lstSpcList)
'    Call cmdReset_Click
'    lstSelList.Clear
'    lstSpcList.Visible = False
'    fraAddItem.Visible = True
'    lblMsg.Visible = False
'End Sub
'
'Private Sub cmdNext_Click(Index As Integer)
'
'    Select Case Index
'    Case 0:
'        iPageNo = iPageNo - 1
'        If iPageCnt > 1 Then cmdNext(1).Enabled = True
'    Case 1:
'        iPageNo = iPageNo + 1
'        If iPageCnt > 1 Then cmdNext(0).Enabled = True
'    End Select
'    Call DisplayOnePage(iPageNo)
'    If chkGraph.Value = 1 Then Call ShowGraph(OldRow)
'
'    If iPageNo = 1 Then
'        cmdNext(0).Enabled = False
'        If iPageCnt > 1 Then cmdNext(1).Enabled = True
'    End If
'    If iPageNo = iPageCnt Then
'        cmdNext(1).Enabled = False
'        If iPageCnt > 1 Then cmdNext(0).Enabled = True
'    End If
'    tblResult.SetFocus
'
'End Sub
'
'Private Sub cmdOK_Click()
'
'    Dim i As Integer
'    Dim SqlStmt As String
'    Dim rsRef As DrSqlOcx.Recordset
'    Dim RefF As Double, RefT As Double
'
'    strDt = Format(Now, CS_DateDbFormat)
'
'    tblResult.ReDraw = False
'
'    For i = 0 To lstSelList.ListCount - 1
'
'        mItemCount = mItemCount + 1
'        ReDim Preserve MyItem(mItemCount)
'        With MyItem(mItemCount)
'            .TestCd = medGetP(lstSelList.List(i), 1, " ")
'            .testnm = Trim(Mid(medGetP(lstSelList.List(i), 1, vbTab), 10))
'            .TestDiv = medGetP(lstSelList.List(i), 2, vbTab)
'            .WorkArea = medGetP(lstSelList.List(i), 3, vbTab)
'            .PanelFg = medGetP(lstSelList.List(i), 4, vbTab)
'            .SpcCd = medGetP(lstSelList.List(i), 5, vbTab)
'            .SpcNm = medGetP(lstSelList.List(i), 6, vbTab)
'            SqlStmt = MySql.SqlGetReference(.TestCd, .SpcCd, strDt, "B", _
'                                            DateDiff("y", Format(MyPatient.DOB, CS_DateMask), Now))
'            Set rsRef = new recordset
'            If rsRef.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
'               rsRef.Close
'               SqlStmt = MySql.SqlGetReference(.TestCd, .SpcCd, strDt, MyPatient.Sex, _
'                                            DateDiff("y", Format(MyPatient.DOB, CS_DateMask), Now))
'               Set rsRef = new recordset
'            End If
'            If rsRef.EOF Then
'               .RefVal = ""
'            Else
'               RefF = Val("" & rsRef.Fields("RefValFrom").Value)
'               RefT = Val("" & rsRef.Fields("RefValTo").Value)
'               .RefVal = Trim("" & rsRef.Fields("RefCd").Value)
'               If RefF <> 0 Or RefT <> 0 Then .RefVal = RefF & " - " & RefT
'            End If
'            rsRef.Close
'
'            tblResult.MaxRows = mItemCount
'            tblResult.Row = mItemCount
'            tblResult.Col = 1: tblResult.Value = .TestCd
'            tblResult.Col = 2: tblResult.Value = .testnm
'            tblResult.Col = 3: tblResult.Value = .SpcCd
'            tblResult.Col = 4: tblResult.Value = .SpcNm
'            tblResult.Col = 16: tblResult.Value = .RefVal
'        End With
'    Next
'
'    Call SetTable
'    tblResult.ReDraw = True
'    txtCumCd.SetFocus
'    fraAddItem.Visible = False
'
'End Sub
'
'Private Sub cmdPrint_Click()
'
'    With tblResult
'        .PrintMarginTop = 100
'        .PrintMarginLeft = 400
'        .PrintJobName = "누적결과레포트 출력"
'
'        .PrintAbortMsg = "누적결과지를 출력중입니다. "
'
'        .PrintOrientation = PrintOrientationLandscape
'        .PrintColor = False
'        .PrintFirstPageNumber = 1
'
'        .PrintHeader = "/n/n/l/fb1 " & "♧ 누적결과 - " & txtPtId.Text & "  " & lblPtNm.Caption & "   " & _
'                                        lblPtSex.Caption & "/" & lblPtAge.Caption & " " & lblAgeDiv.Caption & " /c/fb1/n/n"
'
'        .PrintFooter = "/c/p/fb1"
'
'        .PrintGrid = False
'        .PrintMarginBottom = 500
'        .PrintMarginLeft = 250
'        .PrintMarginRight = 100
'        .PrintShadows = False
'        .PrintMarginTop = 300
'        .PrintNextPageBreakCol = 1
'        .PrintNextPageBreakRow = 1
'        .PrintPageEnd = 2
'        .PrintRowHeaders = False
'        .PrintColHeaders = True
'        .PrintBorder = True
'        '.PrintGrid = True
'        .PrintGrid = True
'        .GridSolid = False
'        .PrintType = PrintTypeAll
'
'        .Action = ActionPrint
'        .GridSolid = True
'    End With
'
'    'If chkGraph.Value = 1 Then Call PrintGraph
'
'End Sub
'
'Private Sub cmdPrintGraph_Click()
'    Call PrintGraph
'End Sub
'
'Private Sub cmdQuery_Click()
'
'    Image1.Visible = True
'    fraStatus.Visible = True
'    fraStatus.ZOrder 0
'    lblStatus.Caption = "해당 환자의 결과내역을 검색중입니다."
'    barStatus.Max = mItemCount * 10
'    barStatus.Min = 0
'    barStatus.Value = 0
'    DoEvents
'
'    Screen.MousePointer = vbArrowHourglass
'
''    txtPtId.Locked = True
''    txtPtId.BackColor = &HE0E0E0
'    txtCumCd.Locked = True
'    txtCumCd.BackColor = &HE0E0E0
'    lstCumList.Enabled = False
'    tabCumCd.Enabled = False
'    cmdItemAdd.Enabled = False
'    dtpFromDt.Enabled = False
'    cmdQuery.Enabled = False
'    cmdPrint.Enabled = True
'
'    Call DisplayResult
'
'    Screen.MousePointer = vbDefault
'    fraStatus.Visible = False
'
'    If lstDtTm.ListCount <= 0 Then
'        MsgBox "해당 환자의 누적결과가 없습니다.", vbInformation, "메세지"
'        Call cmdClear_Click
'        Exit Sub
'    End If
'
'End Sub
'
'Private Sub cmdSpcList_Click()
'
'    lstSpcList.Visible = True
'    lstSpcList.ZOrder 0
'
'End Sub
'
'Private Sub cmdReset_Click()
'
'    Dim i As Integer
'
'    For i = 0 To lstItemList.ListCount - 1
'        lstItemList.Selected(i) = False
'    Next
'    lstSelList.Clear
'
'End Sub
'
'Private Sub dtpFromDt_Click()
'
'    If dtpFromDt.Value > Now Then
'        MsgBox "시작일이 현재날짜보다 큽니다. 다시 설정하십시오.", vbExclamation, "메세지"
'        dtpFromDt.SetFocus
'    End If
'
'End Sub
'
'Private Sub dtpFromDt_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then cmdQuery.SetFocus
'End Sub
'
'Private Sub Form_Load()
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
'    Me.Show
'    DoEvents
'    PtFg = False
'    Call ClearRtn
'    Call LoadCumList(lstCumList, "0")
'    Call LoadWorkArea
'    'Call SetStartDt
'    dtpFromDt.Value = DateAdd("yyyy", -1, Now)
'    'Call GrpDraw
'End Sub
'
'Private Sub lstCumList_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Call lstCumList_MouseDown(1, 0, 0, 0)
'        dtpFromDt.SetFocus
'    End If
'End Sub
'
'Private Sub lstCumList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If Button = 1 Then
'        txtCumCd.Text = medGetP(lstCumList.Text, 1, vbTab)
'        Screen.MousePointer = vbArrowHourglass
'        Call DisplayItem(txtCumCd.Text)
'        Screen.MousePointer = vbDefault
'        blnNewFg = False
'        blnChanged = False
'    End If
'
'End Sub
'
'Private Sub lstItemList_Click()
'    Dim i As Integer
'    i = medListFind(lstSelList, lstItemList.Text)
'    If lstItemList.Selected(lstItemList.ListIndex) Then
'        If i < 0 Then
'            If ItemCheck(medGetP(lstItemList.Text, 1, " ")) Then lstSelList.AddItem lstItemList.Text
'        End If
'    Else
'        If i >= 0 Then lstSelList.RemoveItem (i)
'    End If
'End Sub
'
'Private Function ItemCheck(ByVal pTestCd As String) As Boolean
'    Dim i As Integer
'    ItemCheck = True
'    For i = 1 To mItemCount
'        If MyItem(i).TestCd = pTestCd Then
'            ItemCheck = False
'            Exit For
'        End If
'    Next
'End Function
'
'Private Sub tabCumCd_Click()
'    If tabCumCd.SelectedItem.Index = 1 Then
'        Call LoadCumList(lstCumList, "0")
'    Else
'        Call LoadCumList(lstCumList, mDeptCd)
'        If lstCumList.ListCount <= 0 Then
'            MsgBox "부서코드 " & Chr(34) & mDeptCd & Chr(34) & " 에 등록된 누적코드가 없습니다.", vbExclamation, "메세지"
'            tabCumCd.Tabs(1).Selected = True
'        ElseIf lstCumList.ListCount = 1 Then
'            lstCumList.ListIndex = 0
'            Call lstCumList_MouseDown(1, 0, 0, 0)
'            DoEvents
'            Call cmdQuery_Click
'            DoEvents
'        End If
'    End If
'End Sub
'
'Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim i As Integer
'    Dim sDPfg As String
'    If Row = 0 Then Exit Sub
'    If Row = OldRow Then GoTo Skip1
'    With tblResult
'        .ReDraw = False
'        If OldRow > 0 Then
'            .Col = 2: .Col2 = .MaxCols
'            .Row = OldRow: .Row2 = OldRow
'            .BlockMode = True
'            '.FontSize = 10
'            .FontBold = False
'            .BackColor = OldColor
'            .CellBorderType = 0
'            .Action = ActionSetCellBorder
'            .BlockMode = False
'            .Col = 2: .BackColor = &HE2E8E9
'            .Col = 4: .BackColor = &HEEF4F4  '&HF9FBFB     '&HE7EFEF
'            .RowHeight(OldRow) = 12
'
'            .Col = 17
'            sDPfg = .Value
'            For i = 1 To 10
'                If medGetP(sDPfg, i, ":") <> "" Then
'                    .Col = i + 4
'                    .BackColor = &HC0FFFF     '&HFFF7FF
'                End If
'            Next
'
'        End If
'        .Row = Row: .Col = 1
'        OldColor = .BackColor
'
'        .Col = 2: .Col2 = .MaxCols
'        .Row = Row:  .Row2 = Row
'        .BlockMode = True
'        '.FontSize = 11
'        .FontBold = True
'        .BackColor = &HC0FFFF
'        .CellBorderColor = &H80
'        .CellBorderStyle = CellBorderStyleSolid
'        .CellBorderType = 16
'        .Action = ActionSetCellBorder
'        .BlockMode = False
'        .RowHeight(Row) = 12
'        OldRow = Row
'        .ReDraw = True
'    End With
'Skip1:
'    If chkGraph.Value = 1 Then Call ShowGraph(Row)
'End Sub
'
'Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
'
'    If Row = 0 Then Exit Sub
'    'If Col = 2 Or Col = 4 Or Col = 15 Then
'        tblResult.Row = Row
'        tblResult.Col = Col
'    If Trim(tblResult.Value) <> "" Then
'        MultiLine = 1
'        TipText = "  " & Trim(tblResult.Value)
'        TipWidth = Len(TipText) * 150  '3000
'        tblResult.TextTipDelay = 200
'        'Call spdStat.SetTextTipAppearance("굴림", 9, False, False, &HEEFDF2, vbBlue)    '&H996666)
'        Call tblResult.SetTextTipAppearance("Arial", 11, False, False, vbWhite, vbBlue)    '&H996666)
'        ShowTip = True
'    Else
'        ShowTip = False
'    End If
'End Sub
'
'Private Sub txtCumCd_Change()
'    lstCumList.ListIndex = medListFind(lstCumList, txtCumCd.Text)
'End Sub
'
'Private Sub txtCumCd_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyDown And lstCumList.ListCount > 0 Then
'        lstCumList.SetFocus
'    End If
'
'End Sub
'
'Private Sub txtCumCd_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = vbKeyReturn Then
'        If lstCumList.ListIndex >= 0 Then
'            Call lstCumList_MouseDown(1, 0, 0, 0)
'            dtpFromDt.SetFocus
'        Else
'            txtCumCd.SetFocus
'        End If
'        Exit Sub
'    End If
'    If lstCumList.ListCount > 0 Then
'        Call CodeHelp(KeyAscii, lstCumList, txtCumCd.Text, txtCumCd, dtpFromDt)
'    End If
'End Sub
'
'Private Sub txtCumCd_GotFocus()
'   With txtCumCd
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
'End Sub
'
''% 환자ID가 변경되면 화면Clear
'Private Sub txtPtId_Change()
'    If Not ClearFg Then
'        lblPtNm.Caption = ""
'        lblPtSex.Caption = ""
'        lblPtAge.Caption = ""
'        lblAgeDiv.Caption = ""
'        lblPtDob.Caption = ""
'        lblDeptNm.Caption = ""
'        lblWardId.Caption = ""
'        Call ClearRtn
'    End If
'End Sub
'
''% 환자 ID
'Private Sub txtPtId_GotFocus()
'   With txtPtId
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
'End Sub
'
''% 환자정보 검색
'Private Sub txtPtId_KeyPress(KeyAscii As Integer)
'   If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'
'Private Sub txtPtId_LostFocus()
'
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'    If ActiveControl.Name = cmdExit.Name Then Exit Sub
'    If ActiveControl.Name = cmdClear.Name Then Exit Sub
'    If ActiveControl.Name = cmdCumItem.Name Then Exit Sub
'    If MsgFg Then Exit Sub
'
'    If txtPtId.Text = "" Then
'        txtPtId.SetFocus
'        Exit Sub
'    End If
'
'      With MyPatient
'         If .PtntQuery(txtPtId.Text) Then
'            lblPtNm.Caption = .PtNm
'            lblPtSex.Caption = .SexNm
'            lblPtAge.Caption = .Age
'            lblAgeDiv.Caption = .AgeDiv
'            lblDeptNm.Caption = .DeptNm
'            lblPtDob.Caption = Format(.DOB, CS_DateMask)
'            'lblBedinDt.Caption = Format(.BedInDt, CS_DateMask)
'            'lblBedoutDt.Caption = Format(.BedOutDt, CS_DateMask)
'            cmdItemAdd.Enabled = True
'            ClearFg = False
'            PtFg = True
'            'Call ClearRtn
'            'txtCumCd.Locked = False
'            'txtCumCd.BackColor = vbWhite
'            'lstCumList.Enabled = True
'            'tabCumCd.Enabled = True
'            'cmdItemAdd.Enabled = True
'         Else
'            MsgFg = True
'            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
'            txtPtId.SetFocus
'            PtFg = False
'            MsgFg = False
'            Call txtPtId_GotFocus
'            Exit Sub
'         End If
'      End With
'      'If ActiveControl.Name <> cmdRefresh.Name Then dtpFromDate.SetFocus
'End Sub
'
'
'Private Sub txtRstCmt_DblClick()
'   fraTextResult.Top = (Me.Height - fraTextResult.Height) / 2
'   fraTextResult.Left = (Me.Width - fraTextResult.Width) / 2
'   txtSamCmt1.Text = txtSamCmt.Text
'   txtRstCmt1.Text = txtRstCmt.Text
'   fraTextResult.Visible = True
'   fraTextResult.ZOrder 0
'End Sub
'
'Private Sub txtRstCmt_DragDrop(Source As Control, X As Single, Y As Single)
'    If Source.Name = "txtRstCmt" Then
'        txtSamCmt.Height = txtSamCmt.Height + Y
'        txtRstCmt.Height = txtRstCmt.Height - Y
'        txtRstCmt.Top = txtRstCmt.Top + Y
'    End If
'    txtRstCmt.DragMode = 0
'End Sub
'
'Private Sub txtRstCmt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If Y <= 50 Then
'      txtRstCmt.DragMode = 1
'      txtRstCmt.Drag vbBeginDrag  '1
'   Else
'      txtRstCmt.DragMode = 0
'   End If
'End Sub
'
'Private Sub txtRstCmt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If Y <= 30 Then
'      txtRstCmt.MousePointer = 99
'   Else
'      txtRstCmt.MousePointer = rtfDefault
'   End If
'End Sub
'
'Private Sub txtRstCmt1_DblClick()
'   fraTextResult.Visible = False
'End Sub
'
'Private Sub txtSamCmt_DblClick()
'   Call txtRstCmt_DblClick
'End Sub
'
'Private Sub txtSamCmt_DragDrop(Source As Control, X As Single, Y As Single)
'    If Source.Name = "txtRstCmt" Then
'        txtRstCmt.Height = txtRstCmt.Height + txtSamCmt.Height - Y
'        txtSamCmt.Height = Y
'        txtRstCmt.Top = txtSamCmt.Top + Y
'    End If
'    txtRstCmt.DragMode = 0
'End Sub
'
'Private Sub txtSamCmt1_DblClick()
'   fraTextResult.Visible = False
'End Sub
'
'Private Sub GrpDraw()
'
'Dim Series As Integer
'Dim Points As Integer
'Dim pConstOpt As String
'Dim i As Integer
'Dim j As Integer
'
'   Series = 1
'   Points = 20
'   pConstOpt = "OK"
'   Call grpResult.CloseData(COD_REMOVE)
'   grpResult.TypeEx = CTE_ACTMINMAX
'
'    grpResult.Type = grpResult.Type Or CT_EVENSPACING
'    grpResult.MaxValues = 20                              '//Points
'    grpResult.RealTimeStyle = CRT_LOOPPOS Or CRT_NOWAITARROW
'    grpResult.OpenDataEx COD_VALUES Or COD_ADDPOINTS, Series, Points
'    '
'   'CHART_SOLID 0 Solid Pen.
'   'CHART_DASH 1 Dashed Pen.
'   'CHART_DOT 2 Dotted Pen.
'   'CHART_DASHDOT 3 Dash-Dotted Pen.
'   'CHART_DASHDOTDOT 4 Dash-Dot-Dotted Pen
'    '
'    aryXleg = Array("3/1", "3/4", "3/10", "3/11", "3/20", "3/25", "3/26", "3/28", _
'                          "4/1", "4/2", "4/4", "4/6", "4/8", "4/10", "4/15", "4/20", "4/28", "5/2", "5/5", "5/10")
'    GrpVal = Array(3.85, 3.91, 3.9, 4.11, 2.31, 2.18, 2.77, 2.83, _
'                          2.73, 2.67, 2.6, 2.82, 2.6, 2.94, 3.45, 3.6, 4.28, 4.21, 4.51, 4.55)
'    For i = 0 To Series - 1
'        grpResult.ThisSerie = i
'        grpResult.DecimalsNum(i) = 2
'        '-- 결과값 (X-Axis)
'        For j = 0 To Points - 1
'            grpResult.KeyLeg(j) = aryXleg(j)
'            grpResult.Value(j) = GrpVal(j)
'        Next j
'    Next i
'    '
'   '
'   grpResult.PointLabels = True
'   '
''   grpResult.CloseDataEx COD_VALUES Or COD_ADDPOINTS
'   grpResult.Visible = True
'   cmdPrintGraph.Visible = True
'   '
'End Sub
'
'
'Private Sub DisplayItem(ByVal pCumCd As String)
'
'    Dim i As Integer
'    Dim SqlStmt As String
'    Dim rs As DrSqlOcx.Recordset
'    Dim rsRef As DrSqlOcx.Recordset
'
'    Dim RefF As Double, RefT As Double
'    Dim strDt As String
'
'    strDt = Format(Now, CS_DateDbFormat)
'    SqlStmt = MySql.SqlGetCumItem(pCumCd, 1)
'    Set rs = new recordset
'
''    tblResult.MaxRows = 0
''    tblResult.Row = 0: tblResult.Row2 = 0
''    tblResult.Col = 1: tblResult.COL2 = 12
''    tblResult.BlockMode = True
''    tblResult.Text = ""
''    tblResult.BlockMode = False
'
'    Erase MyItem
'    mItemCount = 0
'    ReDim MyItem(mItemCount)
'
'    tblResult.ReDraw = False
'
'    If Not rs.EOF Then
''        If Not SpcCheck(Trim(rs.Fields("SpcCd").Value)) Then
''            MsgBox "이미 선택된 검체코드와 일치하지 않습니다."
''            GoTo NoData
''        End If
'
''        tblResult.Row = 0: tblResult.Col = 2
''        tblResult.Text = Trim(rs.Fields("SpcNm").Value)
'
''        lblSpcNm.Caption = Trim(rs.Fields("SpcNm").Value)
''        txtSpcCd.Text = Trim(rs.Fields("SpcCd").Value)
''        txtSpcCd.Enabled = False
''        txtSpcCd.BackColor = vbLockColor
'
'        While (Not rs.EOF)
'            mItemCount = mItemCount + 1
'            ReDim Preserve MyItem(mItemCount)
'            With MyItem(mItemCount)
'                .TestCd = Trim(rs.Fields("TestCd").Value)
'                .PanelFg = Trim(rs.Fields("PanelFg").Value)
'                .TestDiv = Trim(rs.Fields("TestDiv").Value)
'                .SpcCd = Trim(rs.Fields("SpcCd").Value)
'                .WorkArea = Trim(rs.Fields("WorkArea").Value)
'                .testnm = Trim(rs.Fields("TestNm").Value)
'                .SpcNm = Trim(rs.Fields("SpcNm").Value)
'                SqlStmt = MySql.SqlGetReference(.TestCd, .SpcCd, strDt, "B", _
'                                                DateDiff("y", Format(MyPatient.DOB, CS_DateMask), Now))
'                Set rsRef = new recordset
'                If rsRef.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
'                   rsRef.Close
'                   SqlStmt = MySql.SqlGetReference(.TestCd, .SpcCd, strDt, MyPatient.Sex, _
'                                                DateDiff("y", Format(MyPatient.DOB, CS_DateMask), Now))
'                   Set rsRef = new recordset
'                End If
'                If rsRef.EOF Then
'                   .RefVal = ""
'                Else
'                   RefF = Val("" & rsRef.Fields("RefValFrom").Value)
'                   RefT = Val("" & rsRef.Fields("RefValTo").Value)
'                   .RefVal = Trim("" & rsRef.Fields("RefCd").Value)
'                   If RefF <> 0 Or RefT <> 0 Then .RefVal = RefF & " - " & RefT
'                End If
'                rsRef.Close
'
'                tblResult.MaxRows = mItemCount
'                tblResult.Row = mItemCount
'                tblResult.Col = 1: tblResult.Value = .TestCd
'                tblResult.Col = 2: tblResult.Value = .testnm
'                tblResult.Col = 3: tblResult.Value = .SpcCd
'                tblResult.Col = 4: tblResult.Value = .SpcNm
'                tblResult.Col = 16: tblResult.Value = .RefVal
'
'            End With
'            If MyItem(mItemCount).PanelFg = PN_Detail Then _
'                Call DisplayDetail(MyItem(mItemCount).TestCd, MyItem(mItemCount).SpcCd, MyItem(mItemCount).SpcNm)
'
'            rs.MoveNext
'        Wend
'        '
'        Call SetTable
'        cmdQuery.Enabled = True
'    End If
'
'NoData:
'    tblResult.ReDraw = True
'
'    rs.Close
'    Set rs = Nothing
'    Set rsRef = Nothing
'
'End Sub
'
'Private Sub SetTable()
'
'    With tblResult
'        .Row = -1
'        .Col = 2: .Col2 = 2
'        .BlockMode = True
'        .ForeColor = &H864B24
'        .BackColor = &HE2E8E9
'        .BlockMode = False
'        .Col = 4: .Col2 = 4
'        .BlockMode = True
'        .ForeColor = &H808080     '&H404040        '&H864B24
'        .BackColor = &HEEF4F4     '&HF9FBFB  '&HE7EFEF
'        .BlockMode = False
'        .Col = 15: .Col2 = 15
'        .BlockMode = True
'        .ForeColor = &H80 ' &HDF6A3E
'        'tblResult.BackColor = &HFBFFFF
'        .BlockMode = False
'        .Col = 16: .Col2 = 16
'        .BlockMode = True
'        .ForeColor = &H136604 ' &HDF6A3E
'        'tblResult.BackColor = &HFBFFFF
'        .BlockMode = False
'        .RowHeight(-1) = 12
'    End With
'
'End Sub
'
'Private Sub DisplayDetail(ByVal pTestCd As String, ByVal pSpcCd As String, ByVal pSpcNm As String)
'
'    Dim i As Integer
'    Dim SqlStmt As String
'    Dim rs As DrSqlOcx.Recordset
'    Dim rsRef As DrSqlOcx.Recordset
'
'    Dim RefF As Double, RefT As Double
'    Dim strDt As String
'
'    strDt = Format(Now, CS_DateDbFormat)
'    SqlStmt = MySql.SqlGetCumDetail(pTestCd)
'    Set rs = new recordset
'
'    While (Not rs.EOF)
'        mItemCount = mItemCount + 1
'        ReDim Preserve MyItem(mItemCount)
'        With MyItem(mItemCount)
'            .TestCd = Trim(rs.Fields("TestCd").Value)
'            .PanelFg = Trim(rs.Fields("PanelFg").Value)
'            .TestDiv = Trim(rs.Fields("TestDiv").Value)
'            .SpcCd = pSpcCd
'            .SpcNm = pSpcNm
'            .WorkArea = Trim(rs.Fields("WorkArea").Value)
'            .testnm = "    " & Trim(rs.Fields("TestNm").Value)
'            SqlStmt = MySql.SqlGetReference(.TestCd, .SpcCd, strDt, "B", _
'                                            DateDiff("y", Format(MyPatient.DOB, CS_DateMask), Now))
'            Set rsRef = new recordset
'            If rsRef.EOF Then  '환자성별에 해당하는 기준치가 없는 경우 "B"(Both)에 해당하는 데이타 검색
'               rsRef.Close
'               SqlStmt = MySql.SqlGetReference(.TestCd, .SpcCd, strDt, MyPatient.Sex, _
'                                            DateDiff("y", Format(MyPatient.DOB, CS_DateMask), Now))
'               Set rsRef = new recordset
'            End If
'            If rsRef.EOF Then
'               .RefVal = ""
'            Else
'               RefF = Val("" & rsRef.Fields("RefValFrom").Value)
'               RefT = Val("" & rsRef.Fields("RefValTo").Value)
'               .RefVal = Trim("" & rsRef.Fields("RefCd").Value)
'               If RefF <> 0 Or RefT <> 0 Then .RefVal = RefF & " - " & RefT
'            End If
'            rsRef.Close
'
'            tblResult.MaxRows = mItemCount
'            tblResult.Row = mItemCount
'            tblResult.Col = 1: tblResult.Value = .TestCd
'            tblResult.Col = 2: tblResult.Value = .testnm
'            tblResult.Col = 3: tblResult.Value = .SpcCd
'            tblResult.Col = 4: tblResult.Value = .SpcNm
'            tblResult.Col = 16: tblResult.Value = .RefVal
'
'        End With
'        rs.MoveNext
'    Wend
'
'    rs.Close
'    Set rs = Nothing
'    Set rsRef = Nothing
'
'End Sub
'
'Private Sub DisplayResult()
'
'    iPageNo = 0
'    iPageCnt = 0
'    Call ReadData
'    If lstDtTm.ListCount = 0 Then Exit Sub
'
'    Call DisplayRemark
'    Call DisplayTxtRst
'    iPageCnt = (lstDtTm.ListCount + 9) \ 10
'    Call cmdNext_Click(1)
'
'End Sub
'
'Private Sub DisplayRemark()
'
'    Dim i As Integer
'    Dim strDtTm As String, j As Integer
'
'    txtSamCmt.Text = ""
'    txtSamCmt1.Text = ""
'    For i = 0 To lstRemark.ListCount - 1
'        strDtTm = medGetP(lstRemark.List(i), 1, vbTab)
'        j = Val(medGetP(lstRemark.List(i), 2, vbTab))
'        txtSamCmt.Text = txtSamCmt.Text & "▶ " & strDtTm & vbCrLf
'        If mCumCol.Item(j).RmkCd <> "" Then _
'            txtSamCmt.Text = txtSamCmt.Text & "* Remark : " & mCumCol.Item(j).Remark & vbCrLf
'        If mCumCol.Item(j).FootNoteFg <> "0" Then _
'            txtSamCmt.Text = txtSamCmt.Text & mCumCol.Item(j).FootNote & vbCrLf
'    Next
'
'End Sub
'
'Private Sub DisplayTxtRst()
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim clsData As clsCumResult
'
'    On Error Resume Next
'    txtRstCmt.Text = ""
'    txtRstCmt1.Text = ""
'    For i = 0 To lstDtTm.ListCount - 1
'        For j = 1 To mItemCount
'            Set clsData = mCumCol.Item(lstDtTm.List(i) & ":" & MyItem(j).TestCd)
'            If clsData.RstText <> "" Then
'                txtRstCmt.Text = txtRstCmt.Text & "▶ " & lstDtTm.List(i) & " : " & clsData.testnm & vbCrLf
'                txtRstCmt.Text = txtRstCmt.Text & clsData.RstText & vbCrLf
'            End If
'        Next
'    Next
'    Set clsData = Nothing
'
'End Sub
'
'Private Sub ReadData()
'
'    Dim strFromDt As String, strToDt As String
'    Dim strSpcCd As String, strWorkArea As String, strTestNm As String
'    Dim strTestCd As String, strPanelFg As String, strTestDiv As String
'    Dim strDtTm As String
'    Dim iSeq  As Integer
'    Dim strList As String
'
'    Dim i As Integer, j As Integer
'    Dim SqlStmt As String
'    Dim rs As DrSqlOcx.Recordset
'
'    Dim clsNewData As clsCumResult
'    Dim MyResult As New clsResultReview
'
'    strFromDt = Format(dtpFromDt.Value, CS_DateDbFormat)
'    strToDt = Format(Now, CS_DateDbFormat)
'
'    lstDtTm.Clear
'    lstRemark.Clear
'    Set mCumCol = New Collection
'    For i = 1 To mItemCount
'
'        barStatus.Value = (i - 1) * 10 + 1
'        lblStatus.Caption = MyItem(i).testnm & " 항목을 검색중입니다."
'        DoEvents
'
'        strTestCd = MyItem(i).TestCd
'        strPanelFg = MyItem(i).PanelFg
'        strTestDiv = MyItem(i).TestDiv
'        strSpcCd = MyItem(i).SpcCd
'        strWorkArea = MyItem(i).WorkArea
'        strTestNm = MyItem(i).testnm
'        SqlStmt = MySql.SqlCumOldResult(txtPtId.Text, strFromDt, strToDt, strTestCd, _
'                                        strPanelFg, strTestDiv, strSpcCd, strWorkArea)
'        Set rs = new recordset
'
'        While (Not rs.EOF)
'
'            If barStatus.Value < (i * 10) Then barStatus.Value = barStatus.Value + 1
'            DoEvents
'
'            iSeq = 0
'            strDtTm = Format(rs.Fields("ColDt").Value, CS_DateMask) & "  " & _
'                      Format(rs.Fields("ColTm").Value, CS_TimeSMask)
'            strList = strDtTm & vbTab & Trim(CStr(iSeq))
'            If medListFind(lstDtTm, strList) < 0 Then lstDtTm.AddItem strList
'
'
'            Set clsNewData = New clsCumResult
'            With clsNewData
'                .ColDt = Format(rs.Fields("ColDt").Value, CS_DateSMask)
'                .ColTm = Format(rs.Fields("ColTm").Value, CS_TimeSMask)
'                .DeptCd = rs.Fields("DeptCd").Value
'                .WardId = rs.Fields("WardId").Value
'                .HosilID = rs.Fields("HosilId").Value
'                .TestCd = rs.Fields("TestCd").Value
'                .testnm = strTestNm
'                .SpcCd = rs.Fields("SpcCd").Value
'                .SpcNm = rs.Fields("TestCd").Value
'                .DPDiv = rs.Fields("DpDiv").Value
'                .HLDiv = rs.Fields("HlDiv").Value
'                .RstDiv = rs.Fields("RstDiv").Value
'                .WorkArea = rs.Fields("WorkArea").Value
'                .AccDt = rs.Fields("AccDt").Value
'                .AccSeq = rs.Fields("AccSeq").Value
'                If Trim(rs.Fields("FootNote").Value) <> "" Then
'                    .FootNoteFg = "Y"    'rs.Fields("FootNote").Value
'                    If medListFind(lstRemark, strDtTm) < 0 Then lstRemark.AddItem strDtTm & vbTab & CStr(mCumCol.Count + 1)
'                End If
'                .RmkCd = ""         'rs.Fields("RmkCd").Value
'                .RstCd = rs.Fields("RstCd").Value
'                .RstUnit = rs.Fields("RstUnit").Value
'                '.TestDiv = rs.Fields("TestDiv").Value
'                .TxtFg = rs.Fields("TxtFg").Value
'                .Remark = ""
'                'If Trim(.RmkCd) <> "" Then
'                '    .Remark = MyResult.ReadRemark(.RmkCd)
'                '    If medListFind(lstRemark, strDtTm) < 0 Then lstRemark.AddItem strDtTm & vbTab & CStr(mCumCol.Count + 1)
'                    'txtSamCmt.Text = txtSamCmt.Text & "-- " & strDtTm & " --" & vbCRLF & "remark : " & .Remark & vbCRLF
'                'End If
'                .FootNote = rs.Fields("FootNote").Value
'                'If Trim(.FootNoteFg) <> "0" Then
'                '    .FootNote = MyResult.ReadFootNote(.WorkArea, .AccDt, .AccSeq)
'                '    If medListFind(lstRemark, strDtTm) < 0 Then lstRemark.AddItem strDtTm & vbTab & CStr(mCumCol.Count + 1)
'                    'If .Remark = "" Then txtSamCmt.Text = txtSamCmt.Text & "-- " & strDtTm & " --" & vbCRLF
'                    'txtSamCmt.Text = txtSamCmt.Text & .FootNote & vbCRLF
'                'End If
'
'                .RstText = rs.Fields("TextResult").Value
'                'If .RstText <> "" Then
'                '    txtRstCmt.Text = txtRstCmt.Text & "-- " & strDtTm & " -- " & .TestNm & vbCRLF & .RstText & vbCRLF
'                'End If
'
'                On Error GoTo Dup_Err
'
'                mCumCol.Add clsNewData, strDtTm & ":" & .TestCd & ":" & Trim(CStr(iSeq))
'            End With
'            rs.MoveNext
'        Wend
'        rs.Close
'        Set rs = Nothing
'    Next
'    Exit Sub
'
'Dup_Err:
'    If Err.Number = 457 Then
'        iSeq = iSeq + 1
'        strList = strDtTm & vbTab & Trim(CStr(iSeq))
'        If medListFind(lstDtTm, strList) < 0 Then lstDtTm.AddItem strList
'        Resume
'    Else
'        MsgBox Err.Number & "  " & Err.Description, vbCritical, "Error"
'        Set rs = Nothing
'    End If
'
'End Sub
'
'Private Sub DisplayOnePage(ByVal iCurPage As Integer)
'    Dim i As Integer
'    Dim j As Integer
'    Dim iListIndex As Integer
'    Dim sDtTm As String
'    Dim sSeq As String
'    Dim sDPfg As String
'    Dim clsData As clsCumResult
'    Dim ErrFg As Boolean
'    Dim EvenBkColor As Long, OddBkColor As Long
'
'    EvenBkColor = &HF9FBFA
'    OddBkColor = &HFFFFFF
'
'    With tblResult
'        .Row = 0: .Row2 = .MaxRows
'        .Col = 5: .Col2 = 14
'        .BlockMode = True
'        .Text = ""
'        .BlockMode = False
'
'        '.ReDraw = False
'
'        For i = 1 To .MaxRows
'            .Row = i: .Row2 = i
'            .Col = 5: .Col2 = 14
'            .BlockMode = True
'            If i <> OldRow Then
'                .BackColor = IIf((i Mod 2) = 0, EvenBkColor, OddBkColor)
'            End If
'            .ForeColor = vbBlack
'            .BlockMode = False
'            .Col = 17: .Value = ""
'        Next
'
'        For i = (iCurPage - 1) * 10 To iCurPage * 10 - 1
'            .Row = 0
'            If i >= lstDtTm.ListCount Then Exit For
'
'            '가장 최근날짜부터 Display하기 위해 Index계산...
'            iListIndex = lstDtTm.ListCount - i - 1
'
'            .Col = i - ((iCurPage - 1) * 10) + 5
'            .Text = Format(medGetP(lstDtTm.List(iListIndex), 1, vbTab), CS_DateSFormat & "  " & CS_TimeSFormat)
'            sDtTm = medGetP(lstDtTm.List(iListIndex), 1, vbTab)
'            sSeq = medGetP(lstDtTm.List(iListIndex), 2, vbTab)
'            For j = 1 To mItemCount
'
'                sDPfg = ""
'                ErrFg = False
'
'                On Error GoTo Err_Trap
'
'                .Row = j
'                .Col = i - ((iCurPage - 1) * 10) + 5
'                Set clsData = mCumCol.Item(sDtTm & ":" & MyItem(j).TestCd & ":" & sSeq)
'                If ErrFg Then GoTo Skip
'                .Value = clsData.RstCd
'                If clsData.HLDiv = "H" Then
'                    .ForeColor = &H7477EF   'vbRed
'                ElseIf clsData.HLDiv = "L" Then
'                    .ForeColor = &HDF6A3E  'vbBlue
'                End If
'                If clsData.DPDiv <> "" Then
'                    sDPfg = clsData.DPDiv
'                    .Value = .Value & " " & clsData.DPDiv
'                    '.FontBold = True
'                    .ForeColor = vbRed
'                    .BackColor = &HC0FFFF     '&HFFF7FF
''                    .CellBorderStyle = CellBorderStyleSolid
''                    .CellBorderType = 16
''                    .CellBorderColor = &H7477EF
''                    .Action = ActionSetCellBorder
'                End If
'                .Col = 15
'                If .Value = "" Then .Value = clsData.RstUnit
'Skip:
'                .Col = 17
'                .Value = .Value & sDPfg & ":"
'
'            Next
'
'            If i = (iCurPage - 1) * 10 Then
'                lblDeptNm.Caption = MyPatient.GetDeptNm(clsData.DeptCd)
'                lblWardId.Caption = clsData.WardId & " - " & clsData.HosilID
'            End If
'
'        Next
'        '.ReDraw = True
'    End With
'    Set clsData = Nothing
'    Exit Sub
'
'Err_Trap:
'    ErrFg = True
'    Resume Next
'End Sub
'
'Private Sub SetStartDt()
'    Dim rs As DrSqlOcx.Recordset
'    Dim SqlStmt As String
'    Dim sDate As Date
'
'    SqlStmt = "select field1 from " & TB_LAB032 & " where cdindex = '" & CD2_StartDate & "' " & _
'              "and cdval1 = '" & mDeptCd & "' "
'    Set rs = new recordset
'    If Not rs.EOF Then
'        sDate = DateAdd("m", Val(rs.Fields("Field1").Value) * (-1), Now)
'    Else
'        sDate = DateAdd("m", -1, Now)
'    End If
'    dtpFromDt.Value = Format(sDate, "YY-MM") & "-01"
'    rs.Close
'    Set rs = Nothing
'End Sub
'
'Private Sub ClearRtn()
'
'    ClearFg = True
'
'    txtCumCd.Text = ""
'    tabCumCd.Tabs(1).Selected = True
'    cmdItemAdd.Enabled = False
'    txtSamCmt.Text = ""
'    txtRstCmt.Text = ""
'    txtSamCmt1.Text = ""
'    txtRstCmt1.Text = ""
'    cmdQuery.Enabled = False
'    cmdPrint.Enabled = False
'    cmdNext(0).Enabled = False
'    cmdNext(1).Enabled = False
'    Call ClearTable
'    Call ClearGraph
'    grpResult.BorderStyle = 1
'    OldRow = -1
'    chkGraph.Value = 0
'    Erase MyItem
'    mItemCount = 0
'    ReDim MyItem(mItemCount)
'
'    iPageNo = 0
'    iPageCnt = 0
'    grpResult.Visible = False
'    cmdPrintGraph.Visible = False
'    Call chkGraph_Click
'    Set mCumCol = New Collection
'
'    lstDtTm.Clear
'    lstRemark.Clear
'
''    txtPtId.Locked = False
''    txtPtId.BackColor = vbWhite
'    txtCumCd.Locked = False
'    txtCumCd.BackColor = vbWhite
'    lstCumList.Enabled = True
'    tabCumCd.Enabled = True
'    cmdItemAdd.Enabled = True
'    dtpFromDt.Enabled = True
'
'    cboWorkArea.ListIndex = -1
'    txtSpcCd.Text = ""
'    txtSpcCd.Enabled = True
'    txtSpcCd.BackColor = vbWhite
'    lblSpcNm.Caption = ""
'    Call cmdReset_Click
'    lstSelList.Clear
'
'End Sub
'
'Private Sub ClearTable()
'    With tblResult
'        .MaxRows = 0
'        .Col = 5: .Col2 = 14
'        .Row = 0: .Row2 = 0
'        .BlockMode = True
'        .Text = ""
'        .BlockMode = False
'
''        .Row = 0: .Col = 2
''        .FontBold = True
''        .FontUnderline = True
'    End With
'End Sub
'
'Private Sub ClearGraph()
'    With grpResult
'        .ClearData CD_VALUES
'        .ClearLegend CHART_LEGEND
'    End With
'End Sub
'
'
'Private Sub ShowGraph(ByVal iGrpRow As Integer)
'
'    Dim i As Integer, j As Integer
'    Dim FirstFg As Boolean
'    Dim iSeries As Integer, iPoints As Integer
'    Dim iMaxValue As Double, iMinValue As Double
'    Dim iFromRef As Double, iToRef As Double
'    Dim sPnt As Integer, ePnt As Integer
'    Dim sXVal As Integer, eXVal As Integer
'    Dim tmpStr As String
'    Dim clsData As clsCumResult
'    Dim ErrFg As Boolean
'
'
'    FirstFg = True
'
'    iSeries = 1
'    iPoints = 0
'
'    'Call InitDraw(iSeries, iPoints)
'    Call SetDateRange(sPnt, ePnt)
'    Call ClearGraph
'
'    With tblResult
'        .Row = iGrpRow: .Col = 2
'        grpResult.Title(CHART_TOPTIT) = .Value
'        grpResult.ClearData CD_VALUES
'        'grpResult.ClearLegend CHART_LEGEND
'
'        grpResult.RealTimeStyle = CRT_LOOPPOS Or CRT_NOWAITARROW
'        grpResult.OpenDataEx COD_VALUES, iSeries, lstDtTm.ListCount
'
'        grpResult.TopGap = 20
'        grpResult.BottomGap = 25
'        grpResult.FixedGap = 33
'        grpResult.Grid = CHART_NOGRID
'        grpResult.Scrollable = True
'
'        .Row = iGrpRow: .Col = 16
'        iFromRef = Val(medGetP(.Value, 1, "-"))
'        iToRef = Val(medGetP(.Value, 2, "-"))
'        iMinValue = iFromRef '- (iFromRef / 50) '2
'        iMaxValue = iToRef '+ (iFromRef / 50) '2
'
'        'grpResult.ItemStyle(CI_HORZGRID) = CHART_SOLID
'
'        grpResult.Scrollable = True
'        grpResult.PointLabels = True
'        grpResult.RGBFont(CHART_POINTFT) = vbBlue
'        grpResult.Axis(AXIS_X).STEP = 1
'        'grpResult.Axis(AXIS_X).Decimals = 0
'        'grpresult.PointLabelsFont.Bold = False
'
'        'Call SetSerLeg
'        'Call SetLegend
'        'Call chkTable_Click
'
'        'For i = 0 To iSeries - 1
'        '    grpResult.Series(i).COLOR = GrpColor(i)
'        'Next
'
'        '-- Stripe Color & Series Color
'        'grpResult.COLOR(1) = COLOR(2) 'QBColor(I)
'
'        grpResult.ThisSerie = 0
'        For i = lstDtTm.ListCount - 1 To 0 Step -1
'
'            sDtTm = medGetP(lstDtTm.List(i), 1, vbTab)
'            sSeq = medGetP(lstDtTm.List(i), 2, vbTab)
'
'            ErrFg = False
'
'            On Error GoTo Err_Trap
'
'            Set clsData = mCumCol.Item(sDtTm & ":" & MyItem(iGrpRow).TestCd & ":" & sSeq)
'            If ErrFg Then GoTo Skip
'            If Not IsNumeric(clsData.RstCd) Then GoTo Skip
'
'            grpResult.KeyLeg(iPoints) = Format(sDtTm, "MM-DD")
'            grpResult.Value(iPoints) = Val(clsData.RstCd)
''            grpResult
'            iPoints = iPoints + 1
'
'            If i = sPnt Then sXVal = iPoints
'            If i = ePnt Then eXVal = iPoints
'
'            If iMinValue > Val(clsData.RstCd) Then iMinValue = Val(clsData.RstCd)
'            If iMaxValue < Val(clsData.RstCd) Then iMaxValue = Val(clsData.RstCd)
'
'Skip:
'        Next
'
'        If iPoints = 0 Then
'            Call ClearGraph
'            Exit Sub
'        End If
'
'        grpResult.CloseData COD_VALUES
'
'        grpResult.OpenDataEx COD_STRIPES, 2, 0
'        '참고치 구간 표시...
'        grpResult.Stripe(0).Axis = AXIS_Y
'        grpResult.Stripe(0).COLOR = &HC0FFFF
'        grpResult.Stripe(0).From = iFromRef
'        grpResult.Stripe(0).To = iToRef
'        'Spread에 보여지고 있는 구간 표시...
'        grpResult.Stripe(1).Axis = AXIS_X
'        grpResult.Stripe(1).COLOR = &HDBF2FD          '&HD6EAFA       ' &HD6EAFA        '&HFFF9F4     '&HF4FEED   '&HD6D7FA     '&HFFF4FF  '&HF7FFFF  '&HEEF4F4  '&HEEEEEE
'        grpResult.Stripe(1).From = sXVal
'        grpResult.Stripe(1).To = eXVal
'        grpResult.CloseData COD_STRIPES
'
'        grpResult.OpenDataEx COD_CONSTANTS, 2, 0
'
'        grpResult.ConstantLine(0).Value = iFromRef
'        grpResult.ConstantLine(0).LineColor = &H808080
'        grpResult.ConstantLine(0).Axis = AXIS_Y
'        grpResult.ConstantLine(0).Label = CStr(iFromRef)
'        grpResult.ConstantLine(0).LineWidth = 1
'        grpResult.ConstantLine(0).LineStyle = CHART_DOT
'
'        grpResult.ConstantLine(1).Value = iToRef
'        grpResult.ConstantLine(1).LineColor = &H808080  '&H80&
'        grpResult.ConstantLine(1).Axis = AXIS_Y
'        grpResult.ConstantLine(1).Label = CStr(iToRef)
'        grpResult.ConstantLine(1).LineWidth = 1
'        grpResult.ConstantLine(1).LineStyle = CHART_DOT
'
'        grpResult.CloseData COD_CONSTANTS
'
'        grpResult.OpenDataEx COD_VALUES, iSeries, iPoints
'
'        grpResult.Axis(AXIS_Y).Min = iMinValue - ((iMaxValue - iFromRef) / 10) '1
'        grpResult.Axis(AXIS_Y).Max = iMaxValue + ((iMaxValue - iFromRef) / 10) '1
'
'        grpResult.Axis(AXIS_Y).STEP = (iMaxValue - iMinValue) / 3
'
'        grpResult.CloseData COD_VALUES
'
'    End With
'    Exit Sub
'
'Err_Trap:
'    ErrFg = True
'    Resume Next
'End Sub
'
'
'Private Sub InitDraw(ByVal nSeries As Integer, ByVal nPoints As Integer)
'
'    Dim iMaxValue As Long
'    Dim iSS As Integer, iPT As Integer, iCnt As Long, iVal As Long
'    Dim i As Integer
'
'    With ssDataBuf
'
'        grpResult.ClearData CD_VALUES
'        grpResult.OpenDataEx COD_VALUES, nSeries, nPoints
'
'        For i = 0 To .MaxRows - 1
'            .Row = i + 1
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
'        Next i
'
'        grpResult.Axis(AXIS_Y).Max = iMaxValue + 1
'
'    End With
'
'End Sub
'
'
'Private Sub LoadWorkArea()
'
'    Dim i%
'    Dim SqlStmt As String
'    Dim tmpRs As DrSqlOcx.Recordset
'
'    SqlStmt = " SELECT cdval1, field1 " & _
'              " FROM " & TB_LAB032 & _
'              " WHERE cdindex = '" & CD2_WorkArea & "'"
'
'    Set tmpRs = new recordset
'
'    If tmpRs.EOF = True Then ' record가 존재하지 않을 경우
'        tmpRs.Close
'        Set tmpRs = Nothing
'        Exit Sub
'    End If
'
'    cboWorkArea.Clear
'    With tmpRs
'        .MoveFirst
'        For i = 1 To .RecordCount
'            cboWorkArea.AddItem .Fields("cdval1").Value & "  " & .Fields("field1").Value
'            .MoveNext
'        Next i
'    End With
'    cboWorkArea.ListIndex = 0
'
'    tmpRs.Close
'    Set tmpRs = Nothing
'
'End Sub
'
'Private Sub txtSpcCd_Change()
'    lstSpcList.ListIndex = medListFind(lstSpcList, txtSpcCd.Text)
'    lstItemList.Clear
'    'lstSelList.Clear
'End Sub
'
'Private Sub txtSpcCd_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyDown And lstSpcList.ListCount > 0 Then
'        lstSpcList.Visible = True
'        lstSpcList.ZOrder 0
'        lstSpcList.SetFocus
'    End If
'
'End Sub
'
'Private Sub txtSpcCd_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        Call lstSpcList_MouseDown(1, 0, 0, 0)
'        lstItemList.SetFocus
'        Exit Sub
'    End If
'    If lstSpcList.ListCount > 0 Then
'        lstSpcList.Visible = True
'        lstSpcList.ZOrder 0
'        Call CodeHelp(KeyAscii, lstSpcList, txtSpcCd.Text, txtSpcCd, lstItemList)
'    End If
'End Sub
'
'Private Sub lstSpcList_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Call lstSpcList_MouseDown(1, 0, 0, 0)
'        lstItemList.SetFocus
'    End If
'End Sub
'
'Private Sub lstSpcList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If Button = 1 Then
'        lstSpcList.Visible = False
'        txtSpcCd.Text = medGetP(lstSpcList.Text, 1, vbTab)
'        lblSpcNm.Caption = medGetP(lstSpcList.Text, 2, vbTab)
'        DoEvents
'        Call LoadItem(lstItemList, medGetP(cboWorkArea.Text, 1, " "), txtSpcCd.Text)
'    End If
'
'End Sub
'
'Private Function SpcCheck(ByVal pSpcCd As String) As Boolean
'    SpcCheck = True
'    If mItemCount = 0 Then Exit Function
'    If MyItem(mItemCount).SpcCd <> pSpcCd Then SpcCheck = False
'End Function
'
'Public Sub LoadItem(ByRef lstList As ListBox, ByVal pWorkArea As String, ByVal pSpcCd As String)
'
'    Dim SqlStmt As String
'    Dim rs As DrSqlOcx.Recordset
'    Dim tmpStr As String
'    Dim i%
'
'    '상세항목 제외...
'    SqlStmt = "select a.testcd, a.testnm, a.rptseq, " & _
'                     "a.testdiv, a.workarea, a.panelfg, c.field3 as SpcNm " & _
'              "from " & TB_LAB001 & " a, " & TB_LAB004 & " b, " & TB_LAB032 & " c " & _
'              "where a.workarea = '" & pWorkArea & "' " & _
'              "and   a.detailfg = '' " & _
'              "and   a.panelfg = '' " & _
'              "and   b.testcd = a.testcd " & _
'              "and   b.spccd = '" & pSpcCd & "' " & _
'              "and   c.cdindex = '" & CD2_Specimen & "' " & _
'              "and   c.cdval1 = '" & pSpcCd & "' " & _
'              "order by a.rptseq"
'    Set rs = new recordset
'
'    If rs.EOF Then GoTo NoData
'
'    lstList.Clear
'    For i = 1 To rs.RecordCount
'        tmpStr = rs.Fields("TestCd").Value & Space(9)
'        lstList.AddItem Mid(tmpStr, 1, 10) & _
'                        rs.Fields("TestNm").Value & Space(30) & vbTab & _
'                        rs.Fields("TestDiv").Value & vbTab & _
'                        rs.Fields("WorkArea").Value & vbTab & _
'                        rs.Fields("PanelFg").Value & vbTab & _
'                        pSpcCd & vbTab & _
'                        rs.Fields("SpcNm").Value
'        rs.MoveNext
'    Next i
'
'NoData:
'    rs.Close
'    Set rs = Nothing
'
'End Sub
'
'Private Sub SetDateRange(sPnt As Integer, ePnt As Integer)
'
'    Dim i As Integer
'    Dim sDt As String, eDt As String
'
'    With tblResult
'        For i = 1 To 10
'            .Row = OldRow
'            .Col = i + 4
'            If IsNumeric(.Value) Then
'                sPnt = lstDtTm.ListCount - ((iPageNo - 1) * 10) - i
'                Exit For
'            End If
'        Next
'        For i = 10 To 1 Step -1
'            .Row = OldRow
'            .Col = i + 4
'            If IsNumeric(.Value) Then
'                ePnt = lstDtTm.ListCount - ((iPageNo - 1) * 10) - i
'                Exit For
'            End If
'        Next
'    End With
'
'End Sub
'
'
'Private Sub PrintGraph()
'
'    With grpResult
'        .Printer.TopMargin = 2
'        .Printer.LeftMargin = 0
'        .Printer.RightMargin = 1
'        .Printer.BottomMargin = 2
'        .Printer.Compress = True
'        .Printer.Orientation = ORIENTATION_LANDSCAPE
'        .Printer.ForceColors = True
'        .PrintIt 0, 0
'    End With
'
'End Sub
'
'
'Public Sub Call_PtId_LostFocus()
'    Call txtPtId_LostFocus
'End Sub
'
'
'Public Sub Call_CumList_MouseDown(ByVal pCumCd As String)
'    Call DisplayItem(pCumCd)
'End Sub
'
'Public Sub Call_Query_Click()
'    Call cmdQuery_Click
'End Sub
