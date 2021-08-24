VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm205ItemDataEntry 
   BackColor       =   &H00DBE6E6&
   Caption         =   "아이템별 결과등록"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis205_Wm.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14625
   Tag             =   "20500"
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "확인(&S)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   43
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00EFE9E4&
      Caption         =   "출력(&P)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   12660
      Style           =   1  '그래픽
      TabIndex        =   42
      Tag             =   "135"
      Top             =   975
      Width           =   1575
   End
   Begin VB.CheckBox chkBatch 
      BackColor       =   &H00DBE6E6&
      Caption         =   "배치 결과 (&B)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   12675
      TabIndex        =   39
      Tag             =   "20507"
      Top             =   210
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.PictureBox picRst 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   75
      ScaleHeight     =   4755
      ScaleWidth      =   14340
      TabIndex        =   10
      Top             =   1590
      Width           =   14400
      Begin MSComctlLib.ProgressBar prgRst 
         Height          =   240
         Left            =   0
         TabIndex        =   22
         ToolTipText     =   "자료를 가져오고 있읍니다."
         Top             =   4530
         Visible         =   0   'False
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread ssRst 
         CausesValidation=   0   'False
         Height          =   4740
         Left            =   0
         TabIndex        =   11
         Tag             =   "20001"
         Top             =   0
         Width           =   14340
         _Version        =   196608
         _ExtentX        =   25294
         _ExtentY        =   8361
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   8
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
         GridColor       =   13816530
         MaxCols         =   19
         MaxRows         =   18
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         SpreadDesigner  =   "Lis205_Wm.frx":038A
         VisibleCols     =   10
         VisibleRows     =   13
      End
      Begin VB.Label lblSpreadLoading 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "잠시 기다려 주세요. 결과 데이터를 로딩하고 있읍니다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3330
         TabIndex        =   23
         Top             =   2520
         Width           =   6675
      End
   End
   Begin VB.Frame fraRst 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10080
      TabIndex        =   26
      Top             =   6330
      Width           =   4395
      Begin VB.CommandButton cmdResult 
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
         Left            =   990
         Style           =   1  '그래픽
         TabIndex        =   34
         Top             =   165
         Width           =   300
      End
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
         Left            =   1305
         TabIndex        =   12
         Tag             =   "opt"
         Top             =   180
         Width           =   1125
      End
      Begin MedControls1.LisLabel lblRstNm 
         Height          =   330
         Left            =   2445
         TabIndex        =   35
         Top             =   180
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         BackColor       =   15265000
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   45
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
         _ExtentX        =   1640
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
         Caption         =   "배치결과"
         Appearance      =   0
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
      Height          =   1620
      Left            =   7170
      TabIndex        =   20
      Tag             =   "20002"
      Top             =   6885
      Width           =   7305
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
         Left            =   6570
         Picture         =   "Lis205_Wm.frx":0D1B
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   1215
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   1260
         Left            =   75
         TabIndex        =   14
         Top             =   270
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   2223
         _Version        =   393217
         BackColor       =   15663102
         Enabled         =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis205_Wm.frx":124D
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
      Height          =   1620
      Left            =   75
      TabIndex        =   18
      Tag             =   "20003"
      Top             =   6885
      Width           =   7080
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
         Left            =   6705
         Picture         =   "Lis205_Wm.frx":14C0
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   1170
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
         Left            =   6705
         Picture         =   "Lis205_Wm.frx":19F2
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   570
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfComment 
         Height          =   630
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   1111
         _Version        =   393217
         BackColor       =   15857140
         ScrollBars      =   2
         TextRTF         =   $"Lis205_Wm.frx":1F24
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
         TabIndex        =   41
         Top             =   1155
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   635
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"Lis205_Wm.frx":2156
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
      Begin VB.Label lblCapRemark 
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
         Left            =   180
         TabIndex        =   29
         Top             =   900
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   75
      TabIndex        =   15
      Top             =   -45
      Width           =   12225
      Begin MSComctlLib.TabStrip tabModeSelect 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   180
         TabIndex        =   27
         Tag             =   "20508"
         Top             =   240
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   582
         Style           =   1
         TabFixedWidth   =   3528
         TabFixedHeight  =   616
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   " By &Work Sheet "
               Key             =   "WorkSheet"
               Object.Tag             =   "WorkSheet"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   " By &Lab No."
               Key             =   "Accession"
               Object.Tag             =   "Accession"
               ImageVarType    =   2
            EndProperty
         EndProperty
         Enabled         =   0   'False
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
      Begin VB.Frame fraWS 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   90
         TabIndex        =   16
         Top             =   630
         Width           =   11925
         Begin VB.CheckBox chkResult 
            BackColor       =   &H00DBE6E6&
            Caption         =   "결과 확인"
            Height          =   255
            Left            =   10560
            TabIndex        =   55
            Top             =   90
            Width           =   1395
         End
         Begin VB.CommandButton cmdHelpList 
            BackColor       =   &H00DEDBDD&
            Caption         =   "▼"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   3120
            MaskColor       =   &H00F4F0F2&
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   54
            Tag             =   "DeptCd"
            Top             =   510
            Width           =   285
         End
         Begin VB.ComboBox cboPoct 
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
            Left            =   5310
            Style           =   2  '드롭다운 목록
            TabIndex        =   53
            Top             =   90
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton cmdTestList 
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
            Height          =   375
            Left            =   6525
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   38
            Top             =   750
            Visible         =   0   'False
            Width           =   270
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
            Height          =   375
            Left            =   3150
            MousePointer    =   14  '화살표와 물음표
            Style           =   1  '그래픽
            TabIndex        =   37
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmdQuery 
            BackColor       =   &H00F4F0F2&
            Caption         =   "조회"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   0
            Left            =   10560
            MaskColor       =   &H00808080&
            Style           =   1  '그래픽
            TabIndex        =   5
            Top             =   390
            Width           =   1320
         End
         Begin VB.TextBox txtTestCd 
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
            Height          =   360
            Index           =   0
            Left            =   1995
            TabIndex        =   1
            Top             =   510
            Width           =   1125
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
            Height          =   360
            Left            =   1995
            TabIndex        =   0
            Top             =   60
            Width           =   1125
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
            Height          =   345
            Left            =   8115
            TabIndex        =   3
            Top             =   510
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            BackColor       =   15857140
            Enabled         =   0   'False
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
            Height          =   345
            Left            =   8670
            TabIndex        =   24
            Top             =   510
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   345
            Left            =   9750
            TabIndex        =   25
            Top             =   510
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
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
            Height          =   345
            Left            =   9195
            TabIndex        =   4
            Top             =   510
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            BackColor       =   15857140
            Enabled         =   0   'False
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
         Begin MSComCtl2.DTPicker dptWorkDt 
            Height          =   330
            Left            =   8100
            TabIndex        =   2
            Top             =   60
            Width           =   1950
            _ExtentX        =   3440
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
            Format          =   64487427
            CurrentDate     =   36287
         End
         Begin MedControls1.LisLabel lblWorkCdNm 
            Height          =   345
            Left            =   3420
            TabIndex        =   31
            Top             =   60
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   609
            BackColor       =   15265000
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
         Begin MedControls1.LisLabel lblTestNm 
            Height          =   360
            Index           =   0
            Left            =   3420
            TabIndex        =   32
            Top             =   510
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   635
            BackColor       =   15265000
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
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   6
            Left            =   105
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
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
            Caption         =   "병동코드  "
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   0
            Left            =   105
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   540
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BackColor       =   10392451
            ForeColor       =   -2147483643
            Enabled         =   0   'False
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
            Caption         =   "검사코드"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   4
            Left            =   6345
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   60
            Width           =   1695
            _ExtentX        =   2990
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
            Caption         =   "작업일자  :"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   345
            Index           =   1
            Left            =   6345
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   510
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            BackColor       =   10392451
            ForeColor       =   -2147483643
            Enabled         =   0   'False
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
            Caption         =   "작업번호(from-to)"
            Appearance      =   0
         End
         Begin VB.Line Line1 
            X1              =   9015
            X2              =   9090
            Y1              =   690
            Y2              =   690
         End
      End
      Begin VB.Frame fraAccession 
         BackColor       =   &H00DBE6E6&
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
         Height          =   930
         Left            =   75
         TabIndex        =   17
         Top             =   630
         Width           =   11925
         Begin VB.CommandButton cmdQuery 
            BackColor       =   &H00F4F0F2&
            Caption         =   "&Query"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   1
            Left            =   10560
            MaskColor       =   &H00808080&
            Style           =   1  '그래픽
            TabIndex        =   9
            Top             =   375
            Width           =   1320
         End
         Begin VB.TextBox txtTestCd 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1995
            TabIndex        =   8
            Top             =   510
            Width           =   1125
         End
         Begin MSMask.MaskEdBox mskAccNo 
            Height          =   330
            Index           =   0
            Left            =   1995
            TabIndex        =   6
            Top             =   60
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "&&-######-#####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskAccNo 
            Height          =   330
            Index           =   1
            Left            =   3840
            TabIndex        =   7
            Top             =   60
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "&&-######-#####"
            PromptChar      =   "_"
         End
         Begin MedControls1.LisLabel lblTestNm 
            Height          =   330
            Index           =   1
            Left            =   3135
            TabIndex        =   30
            Top             =   510
            Width           =   3120
            _ExtentX        =   5503
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
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   2
            Left            =   105
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   60
            Width           =   1845
            _ExtentX        =   3254
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
            Caption         =   "접수번호(From-To)"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   3
            Left            =   105
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   510
            Width           =   1845
            _ExtentX        =   3254
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
            Caption         =   "검사코드"
            Appearance      =   0
         End
         Begin VB.Line Line2 
            X1              =   3705
            X2              =   3780
            Y1              =   195
            Y2              =   195
         End
      End
   End
   Begin VB.ListBox lstTestCode 
      Appearance      =   0  '평면
      BackColor       =   &H00FFF9F7&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   2145
      TabIndex        =   36
      Top             =   1470
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.ListBox lstWSCode 
      Appearance      =   0  '평면
      BackColor       =   &H00F7FBF7&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      ItemData        =   "Lis205_Wm.frx":238A
      Left            =   2160
      List            =   "Lis205_Wm.frx":238C
      TabIndex        =   33
      Top             =   1020
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackColor       =   &H00DDF0F5&
      BackStyle       =   0  '투명
      Caption         =   "오류가 발생했다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00313D46&
      Height          =   180
      Left            =   255
      TabIndex        =   40
      Top             =   8715
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF9F7&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   8640
      Width           =   9870
   End
End
Attribute VB_Name = "frm205ItemDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private insForm         As Form
Private gintTemplete    As Integer

Private WithEvents clsTemplete  As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1

Private objLab032       As clsComcode032
Private objLab301       As clsWSList
Private objPtInfo       As clsPatientInfo
Private objItem         As clsPatientInfo

Private gblnNewObj      As Boolean
Private blnFirst        As Boolean
Private gstrAccMsk      As String
Private RstValidate     As Boolean

Private MsgFg           As Boolean
Private blnRstChange    As Boolean
Private LeaveCellFg     As Boolean
Private blnExpect       As Boolean
Private objSQL  As New clsLISSqlStatistic

Private Sub chkBatch_Click()

    Dim ii As Integer
    
    If chkBatch.Value = 1 Then
        picRst.Height = 4815
        ssRst.Height = 4740
        prgRst.Top = 4530
        fraRst.Visible = True
    Else
        picRst.Height = 5175
        ssRst.Height = 5100
        prgRst.Top = 4890
        txtBatchRst.Text = ""
        lblRstNm.Caption = ""
        Set objCodeList = Nothing
        'If ssRst.MaxRows > 0 Then Call txtBatchRst_LostFocus
        fraRst.Visible = False
    End If
   '
   '
    If ssRst.MaxRows > 0 Then
        EditData
        With objPtInfo
            .ColDt = .SSCol("RESULT")
            If chkBatch.Value = 1 Then
                SpreadLock .SSCol("RESULT"), .SSCol("RESULT"), False
                For ii = 1 To ssRst.MaxRows
                    ssRst.Row = ii
                    ssRst.Col = .SSCol("MAXCOL")
                    ssRst.Value = ""
                    ssRst.Col = .SSCol("RESULT")
                    ssRst.BackColor = &HE0E0E0
                    ssRst.Value = ""
                    If .Result.Item(ii).RstType = "N" Or .Result.Item(ii).RstType = "A" Then
                        .Result.Item(ii).RstType = ""
                    End If
                    .Result.Item(ii).HLDiv = "": .Result.Item(ii).DPDiv = ""
                    .Result.Item(ii).RstVal = "": .Result.Item(ii).RstCd = ""
                Next ii
            Else
                SpreadLock .SSCol("RESULT"), .SSCol("RESULT"), True
                For ii = 1 To ssRst.MaxRows
                    ssRst.Row = ii
                    ssRst.Col = .SSCol("RESULT")
                    ssRst.BackColor = vbWhite
                Next ii
            End If
        End With
    End If

End Sub


Private Sub chkResult_Click()
    If chkResult.Value = 1 Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub

Private Sub cmdHelpList_Click(Index As Integer)
'    Dim objData As clsBasisData
    
'    Set objData = New clsBasisData
    Set objCodeList = New clsPopUpList
    objCodeList.Connection = DBConn
    With objCodeList
        Select Case 1
            Case 0:
'                If optInOut(0).Value Then
'                    .FormCaption = "병동 조회"
'                    .ColumnHeaderText = "병동코드;병동명"
'                    Call .LoadPopUp(GetSQLWardList) ', 3400, 6500) ', ObjLISComCode.WardId)
'                Else
'                    .FormCaption = "진료과 조회"
'                    .ColumnHeaderText = "진료과코드;진료과명"
'                    Call .LoadPopUp(GetSQLDeptList) ', 3400, 6500) ', ObjLISComCode.DeptCd)
'                End If
'
'                txtDeptCd.Text = medGetP(.SelectedString, 1, ";")
'                lblDeptNm.Caption = medGetP(.SelectedString, 2, ";")
            Case 1:
                .FormCaption = "검사항목 조회"
                .ColumnHeaderText = "검사항목코드;검사명"
                Call .LoadPopUp(objSQL.GetAccTest_Wm) ', 3400, 9800)
                txtTestCd(0).Text = medGetP(.SelectedString, 1, ";")
                lblTestNm(0).Caption = medGetP(.SelectedString, 2, ";")
        End Select
    End With
'    Set objData = Nothing
    Set objCodeList = Nothing
End Sub

'2001-10-26 김미경수정 : 스프레트의 내용을 그대로 출력한다.

Private Sub cmdPrint_Click()

    With ssRst
    
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .FontBold = False
        .FontSize = 9
        .BlockMode = False
        .Col = 11: .ColHidden = True
        .Col = 12: .ColHidden = True
        .Col = 13: .ColHidden = True
        .Col = 8: .ColHidden = True
        
        .PrintJobName = "Item별 Worksheet 출력"

        .PrintAbortMsg = "Item별 Worksheet을 출력중입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1

        If tabModeSelect.SelectedItem.Index = 1 Then
            .PrintHeader = "/n/n/l/fb1 " & "♧ 임상병리 Item별 Worksheet /n" & _
                           " -. Worksheet코드 : " & txtWorkCd.Text & " /n" & _
                           " -. 작업일 : " & Format(dptWorkDt.Value, CS_DateShortFormat) & " /n" & _
                           " -. 작업번호 : " & mskFrWorkNo.ClipText & " - " & mskToWorkNo.ClipText & " /n" & _
                           " -. 검사항목 : " & txtTestCd(0).Text & "    " & lblTestNm(0).Caption & ") /c/fb1/n/n"
        Else
            .PrintHeader = "/n/n/l/fb1 " & "♧ 임상병리 Item별 Worksheet /n" & _
                           " -. 검사항목 : " & txtTestCd(1).Text & "    " & lblTestNm(1).Caption & " /n" & _
                           " -. 접수일 : " & mskAccNo(0).Text & " - " & mskAccNo(1).Text & ") /c/fb1/n/n"
        End If
        .PrintFooter = " /l " & String(116, Chr(6)) & "/n/l " & P_HOSPITALNAME & "/c/p/fb1"
     
        .PrintMarginBottom = 100
        .PrintMarginLeft = 200
        .PrintMarginRight = 100
        .PrintShadows = False
        .PrintMarginTop = 300
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint

        .GridSolid = True
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .FontSize = 10
        .FontBold = False
        .BlockMode = False
        .Col = 11: .ColHidden = False
        .ColWidth(11) = 9
        .Col = 12: .ColHidden = False
        .ColWidth(12) = 9
        .Col = 13: .ColHidden = False
        .ColWidth(13) = 11
        .Col = 8: .ColHidden = False
        .ColWidth(8) = 2.75
        
    End With

End Sub

Private Sub cmdResult_Click()
   '
    Dim strSQL As String
    Dim strTestCd  As String
    
    cmdResult.Enabled = False
    If tabModeSelect.SelectedItem.Key = "WorkSheet" Then
        strTestCd = txtTestCd(0).Text
    Else
        strTestCd = txtTestCd(1).Text
    End If
    
    If objPtInfo.GetSqlRstCd(strTestCd) = "" Then
        MsgBox "설정된 결과코드가 없습니다.", vbExclamation
        cmdResult.Enabled = True
        Exit Sub
    End If
    
    Set objCodeList = New clsPopUpList
    
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Result Code List"
        .ColumnHeaderText = "Code;Name"
        .ColumnHeaderWidth = "840.189;2310.236"
        .Tag = "ResultCode"
        .HideSearchTool = True
'        .HideColumnHeaders = True
        .SelectByClick = True
        .FormHeight = 2490 '2250
        .LoadPopUp objPtInfo.GetSqlRstCd(strTestCd)
        
'        .ListCaption = "Result Code List"
'        .ListColHeader = "Name" & vbTab & "Code"
'        .Top = 8620
'        .Left = 11230
'        .Width = 3000
'        .Height = 2000
'        .Tag = "ResultCode"
'        .CaptionOn = False
'        .MultiSel = False
'        strSQL = objPtInfo.GetSqlRstCd(strTestCd)
'        .PopupList strSQL, 2
    End With
    cmdResult.Enabled = True
   '
End Sub

Private Sub cmdTestList_Click()

    If lstTestCode.ListCount = 0 Then
        MsgBox "조건에 맞는 검사코드가 없습니다.", vbExclamation, "메세지"
        Exit Sub
    End If
    lstTestCode.Visible = True
    lstWSCode.Visible = False
    lstTestCode.ZOrder 0
    Call medCodeHelp(0, lstTestCode, txtTestCd(0).Text, txtTestCd(0), cmdQuery(0))

End Sub

Private Sub cmdWSList_Click()
    Dim mvarWardId As String
'    If lstWSCode.ListCount = 0 Then
'        MsgBox "등록된 Worksheet 코드가 없습니다.", vbExclamation, "메세지"
'        Exit Sub
'    End If
'    lstWSCode.Visible = False
'    lstTestCode.Visible = False
'    lstWSCode.ZOrder 0
'    Call medCodeHelp(0, lstWSCode, txtWorkCd.Text, txtWorkCd, txtTestCd(0))

    Dim objDeptHelp As clsPopUpList
    
    Set objDeptHelp = New clsPopUpList
       
    With objDeptHelp
        .Connection = DBConn
        .FormCaption = "병동리스트"
        .ColumnHeaderText = "병동;병동명"
        .LoadPopUp GetSQLWard ', 2000, 1500 ', ObjLISComCode.WardId
        
        mvarWardId = medGetP(.SelectedString, 1, ";")
        txtWorkCd.Text = mvarWardId
        lblWorkCdNm.Caption = medGetP(.SelectedString, 2, ";")
        
        If Trim(mvarWardId) = "" Then
            txtWorkCd.Text = "병동없슴"
        End If
    End With
    
'    Call medCodeHelp(0, lstWSCode, "POCT", txtWorkCd, txtTestCd(0))
    Call medCodeHelp(0, lstWSCode, txtWorkCd.Text, txtWorkCd, txtTestCd(0))
    
    Set objDeptHelp = Nothing
    
End Sub

Private Sub dptWorkDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dptWorkDt_Validate(Cancel As Boolean)
    
    Dim strMaxSeq As String

    mskFrWorkNo.Text = "1___"
    
    If txtWorkCd.Text <> "" Then
        strMaxSeq = CStr(objPtInfo.GetMaxWorkSeq(txtWorkCd.Text, Format(dptWorkDt.Value, CS_DateDbFormat)))
        If Val(strMaxSeq) > 0 Then
            mskToWorkNo.Text = strMaxSeq & String(4 - Len(strMaxSeq), "_")
        Else
            mskToWorkNo.Text = "1___"
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If lstWSCode.Visible Then lstWSCode.Visible = False
        If lstTestCode.Visible Then lstTestCode.Visible = False
        Set objCodeList = Nothing
    End If
End Sub


Private Sub lstTestCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If tabModeSelect.SelectedItem.Key = "WorkSheet" Then
            txtTestCd(0).Text = medGetP(lstTestCode.Text, 1, vbTab)
            lblTestNm(0).Caption = medGetP(lstTestCode.Text, 2, vbTab)
            lstTestCode.Visible = False
            dptWorkDt.SetFocus
        Else
            txtTestCd(1).Text = medGetP(lstTestCode.Text, 1, vbTab)
            lblTestNm(1).Caption = medGetP(lstTestCode.Text, 2, vbTab)
            lstTestCode.Visible = False
            cmdQuery(1).SetFocus
        End If
    End If

End Sub

Private Sub lstTestCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call lstTestCode_KeyDown(vbKeyReturn, 0)

End Sub

Private Sub lstTestCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lstTestCode.SetFocus
End Sub

Private Sub lstWSCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn And lstWSCode.ListIndex >= 0 Then
        txtWorkCd.Text = Trim(Mid(lstWSCode.Text, 1, _
                 InStr(1, lstWSCode.Text, vbTab) - 1))
        lblWorkCdNm.Caption = medGetP(lstWSCode.Text, 2, vbTab)
        lstWSCode.Visible = False
        Call objItem.LoadTestCode(lstTestCode, txtWorkCd.Text, "1")
        txtTestCd(0).SetFocus
    End If
    
        
    Select Case txtWorkCd
     Case "POCT"
          cboPoct.Visible = True
     Case Else
          cboPoct.Visible = False
    End Select
    cboPoct.ListIndex = "0"
End Sub

Private Sub lstWSCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call lstWSCode_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstWSCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lstWSCode.SetFocus
End Sub

Private Sub mskAccNo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub mskFrWorkNo_GotFocus()
    FocusMe Me.mskFrWorkNo
End Sub

Private Sub mskFrWorkNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub mskToWorkNo_GotFocus()
   '
    FocusMe Me.mskToWorkNo
   '
End Sub

Private Sub mskToWorkNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
    Dim ii As Integer
    Dim strValue As String
   '
    If Not IsNull(pSelectedItem) Then
        Select Case objCodeList.Tag
            Case "ResultCode":
                If chkBatch.Value = 1 Then
                    ssRst.Col = objPtInfo.SSCol("RESULT")
                    strValue = medShift(pSelectedItem, ";")
                    txtBatchRst.Text = strValue
                    lblRstNm.Caption = medShift(pSelectedItem, ";")
                    Call txtBatchRst_LostFocus
                End If
            Case "Remark":
                objPtInfo.Result.Item(ssRst.ActiveRow).OTmpCd = medGetP(pSelectedItem, 1, ";")
                objPtInfo.RmkNm = objPtInfo.GetRmkCdNm(objPtInfo.Result.Item(ssRst.ActiveRow).OTmpCd)
                If medGetP(pSelectedItem, 1, ";") <> "" Then
                    rtfRemark.Text = medGetP(pSelectedItem, 2, ";")
                Else
                    rtfRemark.Text = ""
                End If
        End Select
    End If
    Set objCodeList = Nothing
End Sub

Private Sub rtfComment_LostFocus()
    '
    If ssRst.ActiveRow < 0 Then Exit Sub
    objPtInfo.Result.Item(ssRst.ActiveRow).FootNote = rtfComment.Text
    '
End Sub

Private Sub ssRst_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    If Col <> objPtInfo.SSCol("EC") Then Exit Sub
'
    If chkBatch.Value = 1 Then  '배치결과일 경우만...
        ssRst.Row = Row
        ssRst.Col = Col
'        If ssRst.Value = "1" Then  '제외
'            ssRst.Col = objPtInfo.SSCol("RESULT")
'            ssRst.Value = ""
''            ssRst.Col = 19
'            ssRst.Col = 18
'            ssRst.Value = ""
'        Else
'
'            '결과코드값을 보여주기 위해서 수정함
''            ssRst.Col = objPtInfo.SSCol("RESULT")
''            ssRst.Value = txtBatchRst.Text
'
'
'            ssRst.Col = objPtInfo.SSCol("RESULT")
'            ssRst.Value = lblRstNm.Caption
''            ssRst.Col = 19
'            ssRst.Col = 18
'            ssRst.Value = txtBatchRst.Text
'
'        End If
        CheckBatchRst (Row)
    End If
    
End Sub

Private Sub ssRst_EditChange(ByVal Col As Long, ByVal Row As Long)
    ssRst.Row = Row
    ssRst.Col = objPtInfo.SSCol("MAXCOL")
    ssRst.Value = ""
End Sub

Private Sub tabModeSELECT_Click()
    If tabModeSelect.SelectedItem.Key = "WorkSheet" Then
        fraAccession.Visible = False
        fraWS.Visible = True
        lstTestCode.Clear
        mskAccNo(0).Text = gstrAccMsk
        mskAccNo(1).Text = gstrAccMsk
        txtTestCd(1).Text = ""
        lblTestNm(1).Caption = ""
        fraWS.ZOrder 0
        txtWorkCd.SetFocus
    Else
        lstTestCode.Clear
        lstWSCode.Visible = False
        lstTestCode.Visible = False
        fraAccession.Visible = True
        txtWorkCd.Text = ""
        txtTestCd(0).Text = ""
        dptWorkDt.Value = Date
        mskFrWorkNo.Text = "___1"
        mskToWorkNo.Text = "___1"
        lblTestNm(0).Caption = ""
'        lblWorkCdNm.Caption = ""
        fraWS.Visible = False
        fraAccession.ZOrder 0
        mskAccNo(0).SetFocus
    End If
End Sub

Private Sub mskAccNo_Validate(Index As Integer, Cancel As Boolean)
    
    Dim ii As Integer
    Dim strTmp As String
'
    If Trim(mskAccNo(Index).ClipText) = "" Then
        Cancel = True
        Exit Sub
    End If
   '
    If Index = 0 Then
        strTmp = mskAccNo(0).Text
'접수Seq 자릿수 증가로 인한 수정
'2003/12/02 Modify By legends
'        mskAccNo(1).Text = medGetP(strTmp, 1, "-") & "-" & _
'                           medGetP(strTmp, 2, "-") & "-____"
'        mskAccNo(1).SetFocus
'        mskAccNo(1).SelStart = 10
'        mskAccNo(1).SelLength = 4
        mskAccNo(1).Text = medGetP(strTmp, 1, "-") & "-" & _
                           medGetP(strTmp, 2, "-") & "-_____"
        mskAccNo(1).SetFocus
        mskAccNo(1).SelStart = 10
        mskAccNo(1).SelLength = 5
    End If
    '
    If Index = 1 Then
        strTmp = mskAccNo(0).Text
        If (medGetP(mskAccNo(1).Text, 1, "-") <> medGetP(strTmp, 1, "-")) _
            Or (medGetP(mskAccNo(1).Text, 2, "-") <> medGetP(strTmp, 2, "-")) Then
'접수Seq 자릿수 증가로 인한 수정
'2003/12/02 Modify By legends
'            mskAccNo(1).Text = medGetP(strTmp, 1, "-") & "-" & _
'            medGetP(strTmp, 2, "-") & "-____"
            mskAccNo(1).Text = medGetP(strTmp, 1, "-") & "-" & _
            medGetP(strTmp, 2, "-") & "-_____"
        End If
    End If
    '
    If Index = 0 Then
        strTmp = mskAccNo(0).Text
'접수Seq 자릿수 증가로 인한 수정
'2003/12/02 Modify By legends
'        If medGetP(mskAccNo(0).Text, 3, "-") = "____" Then
'            mskAccNo(0).Text = medGetP(strTmp, 1, "-") & "-" & _
'            medGetP(strTmp, 2, "-") & "-1___"
'        End If
        If medGetP(mskAccNo(0).Text, 3, "-") = "_____" Then
            mskAccNo(0).Text = medGetP(strTmp, 1, "-") & "-" & _
            medGetP(strTmp, 2, "-") & "-1____"
        End If
    Else
        strTmp = mskAccNo(1).Text
'접수Seq 자릿수 증가로 인한 수정
'2003/12/02 Modify By legends
'        If medGetP(mskAccNo(1).Text, 3, "-") = "____" Then
'            mskAccNo(1).Text = medGetP(strTmp, 1, "-") & "-" & _
'            medGetP(strTmp, 2, "-") & "-9999"
'        End If
        If medGetP(mskAccNo(1).Text, 3, "-") = "_____" Then
            mskAccNo(1).Text = medGetP(strTmp, 1, "-") & "-" & _
            medGetP(strTmp, 2, "-") & "-99999"
        End If
    End If

   '
End Sub

Private Sub cmdClear_Click()
    ClearData
End Sub

Private Sub cmdExit_Click()
    Set clsTemplete = Nothing
    Set objCodeList = Nothing
    Set objLab301 = Nothing
    Set objPtInfo = Nothing
    Set objItem = Nothing
    
    Unload Me
    Set frm205ItemDataEntry = Nothing
End Sub

Private Sub cmdQuery_Click(Index As Integer)
    
    Dim strCurrentData As String
    Dim ii As Integer, jj As Integer
    Dim SvVal As String
    Dim SvKey As String
   '
   If txtWorkCd.Text = "" Then
        MsgBox "병동을 선택하세요.", vbOKOnly + vbCritical, "병동선택"
        Exit Sub
   End If
   
   If txtTestCd(0).Text = "" Then
        MsgBox "검사항목을 선택하세요.", vbOKOnly + vbCritical, "검사항목선택"
        Exit Sub
   End If
   
    ssRst.MaxRows = 0
    rtfText.Text = ""
    rtfComment.Text = ""
    rtfRemark.Text = ""
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
    MouseRunning

    lblSpreadLoading.Visible = True
    'If Index = 0 Then
    '접수번호별 아이템 결과등록
    ItemResultLoad "AccNo"
    '
    If objPtInfo.TestCount > 0 Then
        Call EditData
        With ssRst
            .ReDraw = False
            .Row = 1: .Row2 = .MaxRows
            If chkBatch.Value = 1 Then
                .Col = objPtInfo.SSCol("RESULT"): .COL2 = objPtInfo.SSCol("RESULT")
                .BlockMode = True
                .BackColor = &HE0E0E0
                .BlockMode = False
    
                SpreadLock objPtInfo.SSCol("RESULT"), objPtInfo.SSCol("RESULT")
    '            For ii = 1 To ssRst.MaxRows
    '               ssRst.Row = ii
    '               ssRst.BackColor = &HE0E0E0
    '            Next ii
            End If
            .Col = objPtInfo.SSCol("SEQ"): .COL2 = objPtInfo.SSCol("SEQ")
            .BlockMode = True
            .FontBold = True
            .ForeColor = &HE48372   '&H7477EF
            .BlockMode = False
            .RowHeight(-1) = 12.5
            .ReDraw = True

            '** 중복되는 접수번호 제거  --> Worksheet별 일때만...
    
            If tabModeSelect.SelectedItem.Key = "WorkSheet" Then
                For ii = 1 To ssRst.MaxRows
                    ssRst.Col = 2
                    ssRst.Row = ii
                    SvVal = ssRst.Value
                    For jj = ii + 1 To ssRst.MaxRows
                        ssRst.Col = 2
                        ssRst.Row = jj
                        If ssRst.Value = SvVal Then
                            'ssRst.Col = objPtInfo.SSCol("EC")
                            'ssRst.Value = 1
                            'ssRst.RowHidden = True
                            ssRst.Col = objPtInfo.SSCol("SEQ")
                            SvKey = ssRst.Value
                            ssRst.Col = objPtInfo.SSCol("ACCNO")
                            SvKey = SvKey & ssRst.Value
                            SvKey = SvKey & txtTestCd(0).Text
    
                            ssRst.Action = ActionDeleteRow
                            ssRst.MaxRows = ssRst.MaxRows - 1
                            objPtInfo.Result.Remove (SvKey)
                        End If
                    Next
                Next
            End If

        End With
        lblErr.Caption = ""
    Else
        ClearData
        lblErr.Caption = "조회하신 자료가 없읍니다."
    End If
    
    
    
    With ssRst
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 9: .ForeColor = DCM_LightRed: .FontBold = True
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
    
    MouseDefault

End Sub


Private Sub ItemResultLoad(ByVal pOption As String)
    
    Dim strParaInfo As String
'
    MouseRunning
    Set objPtInfo.prgBar = prgRst
    objPtInfo.PrgBarInit
    ssRst.Visible = False
    With objPtInfo
        .PtType = RESULT_BY_ITEM                 '/* 결과등록 유형, 반드시 셋팅 해야 됨./

        strParaInfo = txtWorkCd.Text & vbTab & Replace(dptWorkDt.Value, "-", "") & _
                        vbTab & txtTestCd(0).Text
        
        If chkResult.Value = 0 Then
            .LoadTable_Wm strParaInfo, ObjMyUser.EmpId, "AccNo", , ObjSysInfo.BuildingCd
        Else
            .LoadTable_Wm1 strParaInfo, ObjMyUser.EmpId, "AccNo", , ObjSysInfo.BuildingCd
        End If
        
        If .TestCount > 0 Then
            CmdTemplete True
            rtfRemark.Text = objPtInfo.GetRmkCdNm(objPtInfo.Result.Item(1).OTmpCd)
            rtfComment.Text = objPtInfo.Result.Item(1).FootNote
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
            .GetResultSpread ssRst, RESULT_BY_WORKSHEET
        End If
    End With
    MouseDefault
    objPtInfo.PrgBarClear
    DoEvents
   '
End Sub

Private Sub cmdRemarkTemplete_Click()
    Dim RS          As Recordset
    Dim SqlStmt     As String
    Dim strWorkArea As String
    
    With ssRst
         If .DataRowCnt < 1 Then Exit Sub
         .Row = 1: .Col = 2
         strWorkArea = medGetP(.Value, 1, "-")
    End With
    
    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark) & " and " & DBW("field1=", strWorkArea)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    If RS.EOF Then
        SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark)
    End If
    Set RS = Nothing
    

    Set objCodeList = Nothing
    Set objCodeList = New clsPopUpList

'    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE  " & DBW("cdindex =", LC4_Remark) & " "
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
'        .Top = Me.cmdRemarkTemplete.Top + 6000
'        .Left = Me.cmdRemarkTemplete.Left + 100
'        .Width = 6250
'        .Height = 3000
'        .Tag = "Remark"
'        .CaptionOn = True
'        .MultiSel = False
'        .PopupList SqlStmt, 2
'        .ListAdd vbTab & "< 없 음 > ", 2, 1
    End With

End Sub
Private Function ABORhCheck() As Boolean
    Dim ii As Integer
    
    For ii = 1 To ssRst.MaxRows
        With objPtInfo.Result.Item(ii)
            ssRst.Row = ii
            ssRst.Col = objPtInfo.SSCol("RESULT")
            If .TestCd = P_ABOTestCD Or .TestCd = "LB2000" Then
                If .LastRst <> "" Then
                    If UCase(ssRst.Value) <> .LastRst Then
                        MsgBox "혈액형이 이전결과와 일치하지 않습니다.", vbOKOnly + vbCritical, "혈액형등록"
                        Exit Function
                    End If
                End If
            ElseIf .TestCd = P_RHTestCD Or .TestCd = "LB2021" Then
                If .LastRst <> "" Then
                    If UCase(ssRst.Value) <> .LastRst Then
                        MsgBox "RH가 이전결과와 일치하지 않습니다.", vbOKOnly + vbCritical, "RH등록"
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
    ABORhCheck = True
End Function

Private Sub cmdSave_Click()
    
    Dim ii As Long, jj As Long
    Dim blnDBSuccess As Boolean
    Dim SvVal As String
    Dim strYesNo     As String
    

    
    If chkBatch.Value = 1 Then
        If txtBatchRst.Text = "" Then
            MsgBox "결과가 입력되지 않았습니다.."
            txtBatchRst.SetFocus
            Exit Sub
        End If
        If Not RstValidate Then Call txtBatchRst_LostFocus
    End If
    '혈액형 결과체크
    If P_ABOCHK Then
        If ABORhCheck = False Then
            strYesNo = MsgBox("결과등록을 하시겠습니까?.", vbInformation + vbYesNo, "결과등록")
            If strYesNo = vbNo Then Exit Sub
        End If
    End If
    
    With objPtInfo
        .Result.Item(ssRst.ActiveRow).FootNote = rtfComment.Text
        .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
    End With

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
    blnDBSuccess = objPtInfo.ItemDataEntry                    '결과등록을 수행한다.
    If blnDBSuccess = False Then
        MsgBox objPtInfo.ErrText, vbCritical + vbOKOnly, "결과등록 ERROR"
        Exit Sub
    Else
        lblErr.Caption = "자료가 정상적으로 보관되었읍니다."
        ClearData
    End If
    '
    ssRst.MaxRows = 0
    rtfText.Text = ""
    rtfComment.Text = ""
    rtfRemark.Text = ""
   '
End Sub

Private Sub cmdTextTemplete_Click()
    If rtfText.Enabled = False Then Exit Sub
    Call CallTemplete(2, 0)
End Sub

Private Sub cmdCommentTemplete_Click()
    If ssRst.MaxRows < 1 Then Exit Sub
    Call CallTemplete(3, 0)
End Sub


Private Sub Form_Activate()

    '
    If blnFirst = False Then
       blnFirst = True
       ClearData
    End If
    '
    If objLab301 Is Nothing Then
       Set objLab301 = New clsWSList
       objLab301.LoadTable , , "", ""
    End If
    '
    If objItem Is Nothing Then
       Set objItem = New clsPatientInfo
    End If
    '
End Sub

Private Sub Form_Load()
   '
    Me.Show
'접수Seq 자릿수 증가로 인한 수정
'2003/12/02 Modify By legends
'    gstrAccMsk = "__-______-____"
    gstrAccMsk = "__-______-_____"
    blnFirst = False
    'Set OraSE = CreateObject("OracleInProcServer.XOraSession")
    'Set OraDB = OraSE.OpenDatabase(DatabaseName$, Connect$, 0&)
    dptWorkDt.Value = Date
    dptWorkDt.MaxDate = Date
    '
    prgRst.Align = vbAlignBottom
    prgRst.Visible = False
    ssRst.RowHeight(-1) = 12.5
    '
    Call cmdClear_Click
    
    Set objPtInfo = New clsPatientInfo
    Call objPtInfo.LoadWorkSheetCode(ObjSysInfo.BuildingCd, lstWSCode)
    'Call LoadLstWSCode
    Call tabModeSELECT_Click
    
    KeyPreview = True
    
    cboPoct.Clear
    cboPoct.AddItem "0.전체"
    cboPoct.AddItem "1.재활센타"
    cboPoct.AddItem "2.본관"
    cboPoct.ListIndex = "0"

    Call chkBatch_Click
End Sub

Private Sub clsTemplete_CopyTemplete()
   '
    If ssRst.MaxRows < 1 Then Exit Sub
    With objPtInfo
        Select Case gintTemplete
            Case 1:
                If clsTemplete.rtfText.Text <> "" Then
                    rtfRemark.Text = clsTemplete.rtfText.Text
                    objPtInfo.Result.Item(ssRst.ActiveRow).OTmpCd = frm230TempSearch.lblCode.Caption
                    .RmkNm = objPtInfo.GetRmkCdNm(objPtInfo.Result.Item(ssRst.ActiveRow).OTmpCd)
                Else
                    rtfRemark.Text = ""
                    objPtInfo.Result.Item(ssRst.ActiveRow).OTmpCd = ""
                    .RmkNm = ""
                End If
            Case 2:
                rtfText.Text = clsTemplete.rtfText.Text
                .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
                rtfText.SetFocus
            Case 3:
                rtfComment.Text = clsTemplete.rtfText.Text
                .Result.Item(ssRst.ActiveRow).FootNote = rtfComment.Text
                rtfComment.SetFocus
        End Select
    End With
    Set clsTemplete = Nothing

End Sub

Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
    
    Dim strTitle As String
    Dim strWorkArea As String
    

    With ssRst
         If .DataRowCnt < 1 Then Exit Sub
         .Row = 1: .Col = 2
         strWorkArea = medGetP(.Value, 1, "-")
    End With
    
    Set clsTemplete = New frm230TempSearch
    strTitle = Choose(pintPrg, "Remark", "Text Result", "Foot Note")
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
                .lblCode.Caption = objPtInfo.Result.Item(ssRst.ActiveRow).OTmpCd
                .rtfText = rtfRemark.Text
            Case 2:
                .rtfText = rtfText.Text
            Case 3:
                .rtfText = rtfComment.Text
        End Select
        
    End With
    gintTemplete = pintPrg
   
End Sub

Private Sub rtfText_LostFocus()
    '
    objPtInfo.Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
    '
End Sub

Private Sub txtBatchRst_Change()
    RstValidate = False
    If txtBatchRst.Text = "" Then lblRstNm.Caption = ""
End Sub

Private Sub txtBatchRst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then rtfComment.SetFocus
    If KeyCode = vbKeyDown Then Call cmdResult_Click
End Sub

Private Sub txtBatchRst_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBatchRst_LostFocus()

    
    Dim strValue As String
    Dim ii As Long
    '
    If ActiveControl.Name = cmdResult.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If ActiveControl.Name = chkBatch.Name Then Exit Sub

    If tabModeSelect.SelectedItem.Key = "WorkSheet" Then
        lblRstNm.Caption = objPtInfo.GetRstCdNm(txtTestCd(0).Text, txtBatchRst.Text)
    Else
        lblRstNm.Caption = objPtInfo.GetRstCdNm(txtTestCd(1).Text, txtBatchRst.Text)
    End If

    For ii = 1 To ssRst.MaxRows
        ssRst.Row = ii
        ssRst.Col = objPtInfo.SSCol("EC")
        If ssRst.Value = "0" Then
            '결과코드값을 보여주기 위해서....
            'COL=18에 코드 저장
            'COL=2 에 코드 값 저장
'            ssRst.Col = objPtInfo.SSCol("RESULT")
'            ssRst.Value = txtBatchRst.Text
            
            ssRst.Col = objPtInfo.SSCol("RESULT")
            ssRst.Value = lblRstNm.Caption
'            ssRst.Col = 19
            ssRst.Col = objPtInfo.SSCol("MAXCOL")
            ssRst.Value = txtBatchRst.Text
        Else
            '** 변경 By M.G.Choi 2007.04.03
            '   보류체크 적용
            ssRst.Col = objPtInfo.SSCol("RESULT")
            ssRst.Value = lblRstNm.Caption
'            ssRst.Col = 19
            ssRst.Col = objPtInfo.SSCol("MAXCOL")
            ssRst.Value = txtBatchRst.Text
            
            '-- 원본 -----------------------------
'            ssRst.Col = objPtInfo.SSCol("RESULT")
'            ssRst.Value = ""
            '-------------------------------------
        End If
    Next ii
   '
    If CheckBatchRst(1) = "" Then
        txtBatchRst.SetFocus
        'Cancel = True
        For ii = 1 To ssRst.MaxRows
            ssRst.Row = ii
            ssRst.Col = objPtInfo.SSCol("RESULT")
            ssRst.Value = ""
            With objPtInfo.Result.Item(ii)
                If .RstType = "N" Or .RstType = "A" Then
                    .RstType = ""
                End If
                .HLDiv = "": .DPDiv = "":  .RstVal = "": .RstCd = ""
            End With
        Next ii
        txtBatchRst.Text = ""
    End If
   '
    For ii = 1 To ssRst.MaxRows
        CheckBatchRst (ii)
    Next ii
   '
    RstValidate = True
    'txtBatchRst.SetFocus
   '
End Sub

Private Sub txtTestCd_GotFocus(Index As Integer)
    FocusMe Me.txtTestCd(Index)
End Sub

Private Sub txtTestCd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'If Index <> 0 Then Exit Sub
    If lstTestCode.ListCount = 0 Then
        Call objItem.LoadTestCode(lstTestCode, Mid(mskAccNo(0).Text, 1, 2), "2")
    End If
    If KeyCode = vbKeyDown Then
        lstTestCode.Visible = True
        lstWSCode.Visible = False
        lstTestCode.ListIndex = 0
        lstTestCode.ZOrder 0
        lstTestCode.SetFocus
    End If

End Sub

Private Sub txtTestCd_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim Char As String

    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))

    If Index = 0 Then
        If KeyAscii = vbKeyReturn Then
            Call lstTestCode_KeyDown(vbKeyReturn, 0)
            lstTestCode.Visible = False
            Exit Sub
        End If
    
        lstTestCode.Visible = True
        lstWSCode.Visible = False
        lstTestCode.ZOrder 0
        Call medCodeHelp(KeyAscii, lstTestCode, txtTestCd(0).Text, txtTestCd(0), dptWorkDt)
    Else
        If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    End If

End Sub

Private Sub txtTestCd_Validate(Index As Integer, Cancel As Boolean)
    
    Dim strTestCd As String
    Dim strTestNm As String
    '
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub

    strTestCd = txtTestCd(Index).Text
    If strTestCd = "" Then Exit Sub
    With objItem
        strTestNm = .GetTestItemDataEntry(strTestCd)
    End With
   '
    If strTestNm = "" Then
        MsgBox "등록되지 않았거나 사용할 수 없는 검사코드입니다.", vbCritical
        'txtTestCd(Index).Text = ""
        lblTestNm(Index).Caption = ""
        Cancel = True
        FocusMe txtTestCd(Index)
        Exit Sub
    Else
        lblTestNm(Index).Caption = strTestNm
        If objPtInfo Is Nothing Then
            Set objPtInfo = New clsPatientInfo
        Else
            Set objPtInfo = Nothing
            Set objPtInfo = New clsPatientInfo
        End If
        If objPtInfo.GetSqlRstCd(txtTestCd(Index).Text) = "" Then
            cmdResult.Enabled = False
        Else
            cmdResult.Enabled = True
        End If
    End If
   '
End Sub

Private Sub txtWorkCd_Change()
    txtTestCd(0).Text = ""
    lblTestNm(0).Caption = ""
End Sub

Private Sub txtWorkCd_GotFocus()
    '
    FocusMe Me.txtWorkCd
    '
End Sub

Private Sub txtWorkCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstWSCode.ListCount = 0 Then Exit Sub
    If KeyCode = vbKeyDown Then
        lstWSCode.Visible = True
        lstTestCode.Visible = False
        'lstWSCode.ListIndex = 0
        lstWSCode.ZOrder 0
        lstWSCode.SetFocus
    End If
End Sub

Private Sub txtWorkCd_KeyPress(KeyAscii As Integer)

    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = vbKeyReturn Then
        Call lstWSCode_KeyDown(vbKeyReturn, 0)
        lstWSCode.Visible = False
        Exit Sub
    End If
    
    lstWSCode.Visible = True
    lstTestCode.Visible = False
    lstWSCode.ZOrder 0
    Call medCodeHelp(KeyAscii, lstWSCode, txtWorkCd.Text, txtWorkCd, txtTestCd(0))

End Sub



Private Sub txtWorkCd_LostFocus()

    '
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If ActiveControl.Name = lstWSCode.Name Then Exit Sub

    If txtWorkCd.Text = "" Then Exit Sub

'    lblWorkCdNm.Caption = ""
    If Trim(txtWorkCd.Text) = "" Then
        lblWorkCdNm.Caption = ""
        txtWorkCd.SetFocus
        'Cancel = True
        Exit Sub
    End If
    '
'    If objLab301.IsWorkCd(txtWorkCd.Text) = False Then
'        MsgBox "코드 입력 Error!", vbCritical + vbOKOnly, "결과등록 ERROR"
'        lblWorkCdNm.Caption = ""
'        txtWorkCd.SetFocus
'        'Cancel = True
'        FocusMe Me.txtWorkCd
'        Exit Sub
'    End If
    '
    Set objLab032 = New clsComcode032
    With objLab032
        .LoadTable LC3_WorkSheetName, , txtWorkCd.Text
        .MoveFirst
        If .RecordCount > 0 Then
            If Trim(ObjSysInfo.BuildingCd) <> Trim(.Field2) Then
                MsgBox "현재 건물에서는 사용할 수 없는 코드입니다.", vbCritical
                'Cancel = True
                txtWorkCd.SetFocus
                FocusMe Me.txtWorkCd
                Exit Sub
            End If
            lblWorkCdNm.Caption = .Field1
        End If
    End With
    Set objLab032 = Nothing
    '
    lstWSCode.Visible = False

End Sub

Private Sub ClearData()
    
'    tabModeSelect.Enabled = True
    mskAccNo(0).Text = gstrAccMsk
    mskAccNo(1).Text = gstrAccMsk
    txtTestCd(1).Text = ""
'    txtWorkCd.Text = ""
    txtTestCd(0).Text = ""
    dptWorkDt.Value = Date
    mskFrWorkNo.Text = "1___"
    mskToWorkNo.Text = "1___"
'    lblWorkCdNm.Caption = ""
    lblTestNm(0).Caption = ""
    lblTestNm(1).Caption = ""
    lblErr.Caption = ""
    Set objCodeList = Nothing
   '
    If tabModeSelect.SelectedItem.Key = "WorkSheet" Then
        txtWorkCd.BackColor = vbWhite
        txtTestCd(0).BackColor = vbWhite
        dptWorkDt.Enabled = True
        mskFrWorkNo.BackColor = vbWhite
        mskToWorkNo.BackColor = vbWhite
        cmdQuery(0).Enabled = True
    Else
        mskAccNo(0).BackColor = vbWhite
        mskAccNo(1).BackColor = vbWhite
        txtTestCd(1).BackColor = vbWhite
        cmdQuery(1).Enabled = True
    End If
   '
    cmdSave.Enabled = False
    CmdTemplete False
    '
    rtfComment.BackColor = DCM_LightGray
    rtfText.BackColor = DCM_LightGray
    '
    fraComment.Enabled = False
    fraText.Enabled = False
    '
    rtfComment.Text = ""
    rtfText.Text = ""
    rtfRemark.Text = ""
    '
    chkBatch.Enabled = True
    fraRst.Enabled = False
    txtBatchRst.BackColor = DCM_LightGray
    fraWS.Enabled = True
    fraAccession.Enabled = True
    If blnFirst = True Then
        If tabModeSelect.SelectedItem.Key = "WorkSheet" Then
            txtWorkCd.SetFocus
        Else
            mskAccNo(0).SetFocus
        End If
    End If
    '
    txtBatchRst.Text = ""
    lblSpreadLoading.Visible = False
    ssRst.Visible = True
    ssRst.MaxRows = 0
    
    RstValidate = False
    MsgFg = False

    blnExpect = False
    
End Sub

Private Sub EditData()
   '
    tabModeSelect.Enabled = False
    ssRst.Enabled = True
    '
    If tabModeSelect.SelectedItem.Key = "WorkSheet" Then
        txtWorkCd.BackColor = DCM_LightGray
        txtTestCd(0).BackColor = DCM_LightGray
        dptWorkDt.Enabled = False
        mskFrWorkNo.BackColor = DCM_LightGray
        mskToWorkNo.BackColor = DCM_LightGray
        cmdQuery(0).Enabled = False
    Else
        mskAccNo(0).BackColor = DCM_LightGray
        mskAccNo(1).BackColor = DCM_LightGray
        txtTestCd(1).BackColor = DCM_LightGray
        cmdQuery(1).Enabled = False
    End If
    '
    'chkBatch.Enabled = False
    fraRst.Enabled = True
    txtBatchRst.BackColor = vbWhite
    fraWS.Enabled = False
    fraAccession.Enabled = False
    If chkResult.Value = 0 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    fraComment.Enabled = True
    fraText.Enabled = True
    '
    rtfComment.BackColor = &HF1F5F4     'vbWhite
    'rtfText.BackColor = vbWhite   '
    If chkBatch.Value = 1 Then
        txtBatchRst.SetFocus
    Else
        ssRst.SetFocus
    End If
End Sub

Private Sub UpDown1_UpClick()
    If Val(mskFrWorkNo.ClipText) < Val(mskToWorkNo.ClipText) Then
        mskFrWorkNo.Text = FormatUnder(mskFrWorkNo.ClipText, "+")
        mskFrWorkNo.SetFocus
    End If
End Sub

Private Sub UpDown1_DownClick()
    mskFrWorkNo.Text = FormatUnder(mskFrWorkNo.ClipText, "-")
End Sub

Private Sub UpDown2_DownClick()
    If Val(mskToWorkNo.ClipText) > Val(mskFrWorkNo.ClipText) Then
        mskToWorkNo.Text = FormatUnder(mskToWorkNo.ClipText, "-")
        mskToWorkNo.SetFocus
    End If
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
        FormatUnder = "___1"
    End If

End Function

Private Sub ssRst_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i   As Integer
    
    '## 보류표시 Clear
    If Row = 0 And Col = 8 Then
        With ssRst
            blnExpect = IIf(blnExpect, False, True)
            For i = 1 To .MaxRows
                .Row = i: .Col = 8
                If .CellType = CellTypeCheckBox Then
                    .Value = IIf(blnExpect, 0, 1)
                End If
            Next
        End With
    End If
    
    If Col <> objPtInfo.SSCol("EC") Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    If chkBatch.Value = 1 Then  '배치결과일 경우만...
        ssRst.Row = Row
        ssRst.Col = Col
        If ssRst.Value = "1" Then  '제외
            ssRst.Col = objPtInfo.SSCol("RESULT")
            ssRst.Value = ""
'            ssRst.Col = 19
            ssRst.Col = 18
            ssRst.Value = ""
        Else
            
            '결과코드값을 보여주기 위해서 수정함
'            ssRst.Col = objPtInfo.SSCol("RESULT")
'            ssRst.Value = txtBatchRst.Text
        
        
            ssRst.Col = objPtInfo.SSCol("RESULT")
            ssRst.Value = lblRstNm.Caption
'            ssRst.Col = 19
            ssRst.Col = 18
            ssRst.Value = txtBatchRst.Text
        
        End If
'        CheckBatchRst (Row)
    End If
    
    SpDispRtfText
    
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
'    If chkBatch.Value = 1 Then Exit Sub
    
    MsgFg = True
    If Row <= 0 Then Exit Sub
    objPtInfo.SsTop = picRst.Top
    objPtInfo.SsLeft = picRst.Left
    ssRst.Row = Row
    ssRst.Col = Col
    ssRst.Action = ActionActiveCell
    'If chkBatch.Value = 1 Then
    '   Call objPtInfo.PopUp(False, Col)
    'Else
    objPtInfo.MfyFg = False
    Call objPtInfo.PopUp(, Col)
    'End If
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
'                    ssRst.Col = objPtInfo.SSCol("HLDiv"):   ssRst.Value = "N"
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
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"): ssRst.Value = strTmp
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
    
    Dim Col As Long
    Dim Row As Long
   '
    Row = ssRst.ActiveRow
    If Row < 0 Then Exit Sub
    lngResultCol = objPtInfo.SSCol("RESULT")
    lngMaxCol = objPtInfo.SSCol("MAXCOL")

    Col = ssRst.ActiveCol
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
                lblErr.Caption = ""
                Call objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            If objPtInfo.IsAlphaCd = False Then
                strErr = "ALPHA 결과코드 입력 오류!"
                GoTo ErrLevaeCell
            Else
                lblErr.Caption = ""
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
        ssRst.SetFocus
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
                If strResultChk <> "" Then
                    objPtInfo.Result.Item(Row).DPDiv = ""
                    objPtInfo.Result.Item(Row).HLDiv = ""
                End If
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
    
    Exit Sub
   '
ErrLevaeCell:
    With ssRst
        .Row = Row: .Col = objPtInfo.SSCol("RESULT"): .Value = ""
    End With
    Call objPtInfo.ResultCheck
    
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
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
    If Cancel Then Exit Sub
    
    On Error GoTo ErrLevaeCell

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
    End If
    
    On Error GoTo ErrLevaeCell
    '
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
            End If
        ElseIf strRstType = "F" Then
            If objPtInfo.IsFreeResult = False Then
                strErr = "FREE결과 입력 오류! (10자리이내)"
                GoTo ErrLevaeCell
            End If
            Call objPtInfo.NumValCheck
            lblErr.Caption = ""
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
                        ssRst.Col = objPtInfo.SSCol("HLDiv"):   ssRst.Value = "N"
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
                            ssRst.Col = objPtInfo.SSCol("HLDiv"):   ssRst.Value = "N"
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
    
    Exit Sub
    Resume
ErrLevaeCell:
    With ssRst
        .Row = Row: .Col = objPtInfo.SSCol("RESULT"): .Value = ""
    End With
    Call objPtInfo.ResultCheck
    
    MsgBox strErr, vbCritical + vbOKOnly, "결과입력 확인"
    Cancel = True: ssRst.SetFocus
    
End Sub

Private Sub ssRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    '
    If Row < 1 Then Exit Sub
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
        rtfComment.Text = .FootNote
        rtfRemark.Text = objPtInfo.GetRmkCdNm(.OTmpCd)
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


'Private Sub objCodeList_ListClick(ByVal SelList As String)
'
'    Dim ii As Integer
'    Dim strValue As String
'   '
'    If Not IsNull(SelList) Then
'        Select Case objCodeList.Tag
'            Case "ResultCode":
'                If chkBatch.Value = 1 Then
'                    ssRst.Col = objPtInfo.SSCol("RESULT")
'                    strValue = medShift(SelList, vbTab)
'                    txtBatchRst.Text = strValue
'                    lblRstNm.Caption = medShift(SelList, vbTab)
'                    Call txtBatchRst_LostFocus
'                End If
'            Case "Remark":
'                objPtInfo.Result.Item(ssRst.ActiveRow).OTmpCd = medGetP(SelList, 1, vbTab)
'                objPtInfo.RmkNm = objPtInfo.GetRmkCdNm(objPtInfo.Result.Item(ssRst.ActiveRow).OTmpCd)
'                If medGetP(SelList, 1, vbTab) <> "" Then
'                    rtfRemark.Text = medGetP(SelList, 2, vbTab)
'                Else
'                    rtfRemark.Text = ""
'                End If
'        End Select
'    End If
'    Set objCodeList = Nothing
'    '
'End Sub

Private Sub SpreadLock(ByVal Col As Long, ByVal COL2 As Long, Optional ByVal NoLock As Boolean = False)
    With ssRst
        .Row = 1: .Row2 = .MaxRows
        .Col = Col: .COL2 = COL2
        .BlockMode = True
        If NoLock Then
            .Lock = False
            .Protect = False
        Else
            .Lock = True
            .Protect = True
        End If
        .BlockMode = False
    End With
End Sub

Private Function CheckBatchRst(Row) As String
    
    Dim strErr As String
    Dim strRstType As String
    
    ssRst.Row = Row
    CheckBatchRst = ""
    objPtInfo.ResultCheck (Row)
    strRstType = objPtInfo.Result.Item(Row).RstType
    
    If strRstType = "N" Then
        strErr = objPtInfo.Result.Item(Row).AvalVal
        If objPtInfo.IsAvalVal(Row) = False Then
            If strErr <> "0" Then
                strErr = "유효숫자 입력 오류. (" & objPtInfo.Result.Item(Row).AvalVal & "자리)"
            Else
                strErr = "유효숫자 입력 오류. (정수형만 입력)"
            End If
            GoTo ErrCheckBatchRst
        Else
            lblErr.Caption = ""
            objPtInfo.NumValCheck (Row)
        End If
    ElseIf strRstType = "A" Then
        If objPtInfo.IsAlphaCd(Row) = False Then
            strErr = "ALPHA 결과코드 입력 오류!"
            GoTo ErrCheckBatchRst
        Else
           lblErr.Caption = ""
        End If
    ElseIf strRstType = "R" Then
        If objPtInfo.IsRateCd(Row) = False Then
            strErr = "비율결과 입력 오류!"
            GoTo ErrCheckBatchRst
        Else
           lblErr.Caption = ""
        End If
    ElseIf strRstType = "F" Then
        If objPtInfo.IsFreeResult(Row) = False Then
            strErr = "FREE결과 입력 오류! (10자리이내)"
            GoTo ErrCheckBatchRst
        Else
           lblErr.Caption = ""
        End If
    End If
       CheckBatchRst = "OK"
    Exit Function

ErrCheckBatchRst:
    CheckBatchRst = ""
    objPtInfo.ResultCheck
    If Row = 1 Then
        MsgBox strErr, vbCritical + vbInformation, "배치결과입력 확인"
    End If

End Function



